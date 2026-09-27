"""Einmaliger Live-Test des kompletten Ablaufs: PDF erzeugen + wirklich versenden.

ACHTUNG - dieses Skript verschickt echte Mails ueber Microsoft Graph. Es ist
ein Test-Werkzeug und ausdruecklich NICHT Teil des automatischen Laufs: weder
automatischer_lauf.py noch src/main.py rufen es auf.

Zweck: gegen die echten Azure-Zugangsdaten pruefen, ob Tokenabruf, Anhang und
Direktversand funktionieren, bevor der erste produktive Lauf stattfindet.

Es kann an mehrere Adressen nacheinander senden. Da mail_sender den Empfaenger
aus "test_empfaenger_override" in config/einstellungen.json liest, wird die
Datei je Adresse einmal umgeschrieben und sende_rechnung() erneut aufgerufen.
Die PDF entsteht dabei nur EINMAL und wird fuer alle Sendungen wiederverwendet.

Bewusst unberuehrt bleiben:

* config/kunden.json - der Testkunde ist frei erfunden und steht unten im
  Skript.
* logs/rechnungsnummern_zaehler.json - die Rechnungsnummer ist als fester
  String gesetzt und traegt das Praefix "TEST" statt "RE";
  src/rechnungsnummer.py wird nicht importiert.
* logs/monatssperre.json - src/monatssperre.py wird nicht importiert.
* output/rechnungen/ - die PDF entsteht in einem temporaeren Ordner und wird
  am Ende geloescht.

config/einstellungen.json wird waehrend des Laufs auf "versand_modus":
"direkt" gesetzt und im finally-Block wieder auf "entwurf" zurueckgestellt -
auch wenn ein Versand scheitert. "test_empfaenger_override" wird dabei auf
den Wert vor dem Lauf zurueckgesetzt, nie auf der zuletzt getesteten Adresse.

Aufrufe aus dem Projektroot:

    .\\venv\\Scripts\\python.exe scripts\\probe_mailversand.py --empfaenger "a@x.de"
    .\\venv\\Scripts\\python.exe scripts\\probe_mailversand.py --empfaenger "a@x.de,b@y.de"
"""

import argparse
import json
import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path

from dateutil.relativedelta import relativedelta

PROJEKT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJEKT_DIR))

from src import mail_sender, pdf_generator  # noqa: E402
from src.config_loader import load_config  # noqa: E402
from src.mail_sender import MailSendError, sende_rechnung  # noqa: E402
from src.pdf_generator import erstelle_rechnung  # noqa: E402

EINSTELLUNGEN_DATEI = PROJEKT_DIR / "config" / "einstellungen.json"

# Praefix bewusst nicht "RE", damit diese Nummer nie mit einer echten
# Rechnung verwechselt werden kann.
TEST_RECHNUNGSNUMMER = "TEST202609901"

TESTKUNDE = {
    "kunde_id": "probe_mailversand",
    "firma": "Musterhandwerk Schmidt GmbH",
    "strasse": "Gewerbering 14",
    "plz_ort": "33602 Bielefeld",
    "email_empfaenger": "test@example.invalid",
    "beschreibung": (
        "Smart-Website Servicepauschale: Laufende Betreuung inklusive "
        "technischer Wartung (Updates, Bugfixes), Hosting, "
        "Sicherheitsüberwachung sowie Aktualisierung rechtlicher Inhalte "
        "gemäß Servicevertrag"
    ),
    "anzahl": 1,
    "einzelpreis": 89.00,
    "steuersatz": 19.0,
    "zahlungsziel_tage": 14,
}


def bestimme_leistungszeitraum(heute: date):
    """Liefert (erster_tag, letzter_tag) des zuletzt abgeschlossenen Monats.

    Nachbildung von main.bestimme_leistungszeitraum; ein Import von main wuerde
    unnoetig die ganze Orchestrierung mitziehen.
    """
    von = (heute - relativedelta(months=1)).replace(day=1)
    bis = von + relativedelta(day=31)  # klemmt auf den letzten Tag des Monats
    return von, bis


def zerlege_empfaenger(text: str):
    """Macht aus "a@x.de, b@y.de" die Liste ["a@x.de", "b@y.de"].

    Leere Eintraege werden verworfen, Dubletten in der Reihenfolge des ersten
    Auftretens entfernt - sonst ginge dieselbe Testmail doppelt raus.

    Raises:
        ValueError: Wenn keine einzige brauchbare Adresse uebrig bleibt.
    """
    adressen = []
    for roh in text.split(","):
        adresse = roh.strip()
        if adresse and adresse not in adressen:
            adressen.append(adresse)
    if not adressen:
        raise ValueError(
            f"--empfaenger enthaelt keine Adresse: {text!r}. "
            f"Erwartet wird eine kommagetrennte Liste, z. B. \"a@x.de,b@y.de\"."
        )
    return adressen


def schreibe_einstellungen(daten: dict) -> None:
    """Schreibt config/einstellungen.json mit dem uebergebenen Inhalt."""
    with EINSTELLUNGEN_DATEI.open("w", encoding="utf-8") as f:
        json.dump(daten, f, indent=2, ensure_ascii=False)
        f.write("\n")


class GraphMitschnitt:
    """Protokolliert die tatsaechlich an Graph gesendeten Anfragen.

    mail_sender.sende_rechnung() gibt selbst nichts zurueck. Statt Betreff und
    Empfaenger nachzubauen (und damit womoeglich etwas anderes zu berichten,
    als wirklich rausging), wird hier requests.post umschlossen: die Anfrage
    laeuft unveraendert durch, Nutzlast und Antwort werden nur mitgeschrieben.
    """

    def __init__(self):
        self.aufrufe = []
        self._original = None

    def __enter__(self):
        self._original = mail_sender.requests.post

        def aufzeichnend(url, *args, **kwargs):
            antwort = self._original(url, *args, **kwargs)
            if url.startswith(mail_sender.GRAPH_BASIS):
                self.aufrufe.append(
                    {
                        "url": url,
                        "nutzlast": kwargs.get("json"),
                        "status": antwort.status_code,
                        "grund": antwort.reason,
                        "kopfzeilen": dict(antwort.headers),
                        "rumpf": _rumpf_als_dict(antwort),
                    }
                )
            return antwort

        mail_sender.requests.post = aufzeichnend
        return self

    def __exit__(self, *_):
        mail_sender.requests.post = self._original
        return False


def _rumpf_als_dict(antwort):
    """Liefert die JSON-Antwort als dict; leeres dict, wenn Graph nichts sendet."""
    try:
        return antwort.json() or {}
    except ValueError:
        return {}


def erzeuge_test_pdf(tempordner: Path, heute: date, von: date, bis: date) -> Path:
    """Erzeugt die Test-PDF im temporaeren Ordner statt in output/rechnungen.

    erstelle_rechnung() schreibt immer nach AUSGABE_DIR/{Jahr}/{Monat}; damit
    der echte Ausgabebaum unberuehrt bleibt, zeigt AUSGABE_DIR waehrend des
    Aufrufs auf den temporaeren Ordner.
    """
    try:
        firmendaten = load_config("firmendaten")
    except Exception:
        firmendaten = {}

    original_ausgabe_dir = pdf_generator.AUSGABE_DIR
    try:
        pdf_generator.AUSGABE_DIR = tempordner
        return erstelle_rechnung(
            kunde=TESTKUNDE,
            firmendaten=firmendaten,
            rechnungsnummer=TEST_RECHNUNGSNUMMER,
            rechnungsdatum=heute,
            zeitraum_von=von,
            zeitraum_bis=bis,
        )
    finally:
        pdf_generator.AUSGABE_DIR = original_ausgabe_dir


def berichte_graph_antwort(aufruf: dict, einzug: str = "  ") -> None:
    """Gibt Betreff, Empfaenger und Graph-Antwort einer Anfrage aus."""
    nachricht = (aufruf["nutzlast"] or {}).get("message", aufruf["nutzlast"]) or {}
    empfaenger = [
        e.get("emailAddress", {}).get("address")
        for e in nachricht.get("toRecipients", [])
    ]
    anhaenge = nachricht.get("attachments", [])

    print(f"{einzug}Graph-Endpunkt : {aufruf['url']}")
    print(f"{einzug}Betreff        : {nachricht.get('subject')!r}")
    print(f"{einzug}Empfaenger     : {', '.join(a for a in empfaenger if a)}")
    for anhang in anhaenge:
        groesse = anhang.get("size")
        if groesse is None and anhang.get("contentBytes"):
            # Graph meldet size nicht zurueck; Base64 -> ungefaehre Bytezahl.
            groesse = len(anhang["contentBytes"]) * 3 // 4
        print(f"{einzug}Anhang         : {anhang.get('name')} (~{groesse} Bytes)")
    print(f"{einzug}HTTP-Status    : {aufruf['status']} {aufruf['grund']}")

    rumpf = aufruf["rumpf"]
    if rumpf.get("id"):
        print(f"{einzug}Message-ID     : {rumpf['id']}")
    else:
        print(f"{einzug}Message-ID     : keine - sendMail antwortet mit "
              f"202 Accepted ohne Rumpf")

    # Graph liefert pro Anfrage eine Korrelations-ID; das ist die einzige
    # Kennung, die sich bei sendMail spaeter zuordnen laesst.
    for kopf in ("request-id", "client-request-id"):
        if aufruf["kopfzeilen"].get(kopf):
            print(f"{einzug}{kopf:<15}: {aufruf['kopfzeilen'][kopf]}")


def sende_an_adresse(adresse, pdf_pfad, heute, von, bis, grundeinstellungen):
    """Stellt die Einstellungen auf diese eine Adresse um und versendet.

    Returns:
        Ergebnis-Dictionary mit Adresse, Betreff, Status und ggf. Fehlertext.
    """
    einstellungen = dict(grundeinstellungen)
    einstellungen["versand_modus"] = "direkt"
    einstellungen["test_empfaenger_override"] = adresse
    schreibe_einstellungen(einstellungen)

    # Bewusst aus der Datei nachladen: so wird genau das verschickt, was auch
    # auf der Platte steht - kein Auseinanderlaufen von Datei und Aufruf.
    einstellungen = load_config("einstellungen")
    api_konfig = load_config("api_key")

    print(f"  versand_modus            : {einstellungen['versand_modus']}")
    print(f"  test_empfaenger_override : {einstellungen['test_empfaenger_override']}")
    print(f"  Empfaenger im Datensatz  : {TESTKUNDE['email_empfaenger']} "
          f"(wird durch den Override ersetzt)")
    print("  ... sende ...")

    ergebnis = {"adresse": adresse, "betreff": None, "status": None, "fehler": None}
    mitschnitt = GraphMitschnitt()
    try:
        with mitschnitt:
            sende_rechnung(
                kunde=TESTKUNDE,
                rechnungsnummer=TEST_RECHNUNGSNUMMER,
                pdf_pfad=pdf_pfad,
                leistungszeitraum_von=von,
                leistungszeitraum_bis=bis,
                einstellungen=einstellungen,
                api_konfig=api_konfig,
            )
    except MailSendError as exc:
        ergebnis["fehler"] = str(exc)
        print()
        print("  ERGEBNIS: Graph hat die Anfrage ABGELEHNT.")
        print(f"  {exc}")
    else:
        print()
        print("  ERGEBNIS: Graph hat die Anfrage angenommen.")
        print(f"  {mail_sender.beschreibe_versand(TESTKUNDE, einstellungen)}")

    for aufruf in mitschnitt.aufrufe:
        berichte_graph_antwort(aufruf)
        nachricht = (aufruf["nutzlast"] or {}).get("message", aufruf["nutzlast"]) or {}
        ergebnis["betreff"] = nachricht.get("subject")
        ergebnis["status"] = f"{aufruf['status']} {aufruf['grund']}"

    return ergebnis


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--empfaenger",
        required=True,
        metavar="ADRESSEN",
        help="Kommagetrennte Liste der Testempfaenger. An jede Adresse geht "
             "eine eigene Mail mit derselben PDF.",
    )
    argumente = parser.parse_args()

    try:
        adressen = zerlege_empfaenger(argumente.empfaenger)
    except ValueError as exc:
        print(f"FEHLER: {exc}")
        return 1

    heute = date.today()
    von, bis = bestimme_leistungszeitraum(heute)

    # --- Schritt 1: Einstellungen sichern -----------------------------------
    vorher = load_config("einstellungen")
    print("=" * 78)
    print("SCHRITT 1  Einstellungen sichern")
    print("=" * 78)
    print("  Sicherung des aktuellen Inhalts von config/einstellungen.json")
    print("  (nur hier im Terminal, es wird keine Sicherungsdatei angelegt):")
    print(f"    {json.dumps(vorher, ensure_ascii=False)}")
    print(f"  Testempfaenger ({len(adressen)}):")
    for nummer, adresse in enumerate(adressen, start=1):
        print(f"    {nummer}. {adresse}")

    tempordner = Path(tempfile.mkdtemp(prefix="probe_mailversand_"))
    ergebnisse = []
    abbruch = None

    try:
        # --- Schritt 2: fiktive Testrechnung erzeugen (nur einmal) ----------
        print()
        print("=" * 78)
        print("SCHRITT 2  Fiktive Testrechnung erzeugen (einmal fuer alle Sendungen)")
        print("=" * 78)
        netto = float(TESTKUNDE["anzahl"]) * float(TESTKUNDE["einzelpreis"])
        brutto = netto * (1 + float(TESTKUNDE["steuersatz"]) / 100)
        print(f"  Kunde             : {TESTKUNDE['firma']} (frei erfunden)")
        print(f"  Rechnungsnummer   : {TEST_RECHNUNGSNUMMER} (fest, ohne Zaehler)")
        print(f"  Rechnungsdatum    : {heute:%d.%m.%Y}")
        print(f"  Leistungszeitraum : {von:%d.%m.%Y} – {bis:%d.%m.%Y}")
        print(f"  Netto/Brutto      : {netto:.2f} / {brutto:.2f} EUR")
        print(f"  Beschreibung      : {len(TESTKUNDE['beschreibung'])} Zeichen")

        pdf_pfad = erzeuge_test_pdf(tempordner, heute, von, bis)
        print(f"  PDF               : {pdf_pfad.name} "
              f"({pdf_pfad.stat().st_size} Bytes)")
        print(f"  temporaer unter   : {pdf_pfad.parent}")

        # --- Schritt 3: an jede Adresse einzeln versenden -------------------
        for nummer, adresse in enumerate(adressen, start=1):
            print()
            print("=" * 78)
            print(f"SCHRITT 3.{nummer}  Versand an {adresse} (echter Versand!)")
            print("=" * 78)
            ergebnisse.append(
                sende_an_adresse(adresse, pdf_pfad, heute, von, bis, vorher)
            )

    except Exception as exc:
        abbruch = exc
        print()
        print(f"  ABBRUCH vor bzw. waehrend des Versands - "
              f"{type(exc).__name__}: {exc}")

    finally:
        # --- Schritt 4: aufraeumen ------------------------------------------
        print()
        print("=" * 78)
        print("SCHRITT 4  Aufraeumen")
        print("=" * 78)

        zurueck = dict(vorher)
        zurueck["versand_modus"] = "entwurf"
        zurueck["test_empfaenger_override"] = vorher.get("test_empfaenger_override")
        schreibe_einstellungen(zurueck)
        kontrolle = load_config("einstellungen")
        print(f"  config/einstellungen.json    : "
              f"{json.dumps(kontrolle, ensure_ascii=False)}")
        print(f"  versand_modus == 'entwurf'   : "
              f"{kontrolle['versand_modus'] == 'entwurf'}")
        print(f"  override == Wert vor Lauf    : "
              f"{kontrolle.get('test_empfaenger_override') == vorher.get('test_empfaenger_override')}")

        shutil.rmtree(tempordner, ignore_errors=True)
        print(f"  temporaerer Ordner geloescht : {not tempordner.exists()}")

        for datei in ("rechnungsnummern_zaehler.json", "monatssperre.json"):
            pfad = PROJEKT_DIR / "logs" / datei
            print(f"  logs/{datei:<32}: "
                  f"{'existiert' if pfad.exists() else 'nicht vorhanden'} "
                  f"(unberuehrt)")

    # --- Zusammenfassung ----------------------------------------------------
    print()
    print("=" * 78)
    print("ZUSAMMENFASSUNG")
    print("=" * 78)
    for ergebnis in ergebnisse:
        zustand = "ABGELEHNT" if ergebnis["fehler"] else "angenommen"
        print(f"  {ergebnis['adresse']}")
        print(f"    Status  : {ergebnis['status'] or '-'}  ({zustand})")
        print(f"    Betreff : {ergebnis['betreff']!r}")
        if ergebnis["fehler"]:
            print(f"    Fehler  : {ergebnis['fehler'].splitlines()[0]}")

    fehlgeschlagen = [e for e in ergebnisse if e["fehler"]]
    nicht_versucht = len(adressen) - len(ergebnisse)
    if nicht_versucht:
        print(f"  {nicht_versucht} Adresse(n) wurden wegen des Abbruchs nicht "
              f"mehr angesprochen.")

    print()
    if abbruch is None and not fehlgeschlagen:
        print(f"Probe erfolgreich: {len(ergebnisse)} von {len(adressen)} Mails "
              f"von Graph angenommen.")
    else:
        print("Probe unvollstaendig - siehe Meldungen oben.")
    print("Hinweis: Dieses Skript ist ein Test-Werkzeug und NICHT Teil des")
    print("automatischen Laufs (automatischer_lauf.py / src/main.py).")
    print("=" * 78)

    return 1 if (abbruch or fehlgeschlagen or nicht_versucht) else 0


if __name__ == "__main__":
    sys.exit(main())
