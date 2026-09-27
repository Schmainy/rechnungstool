"""Versand der fertigen Rechnungen per E-Mail ueber Microsoft Graph.

Authentifizierung erfolgt im Client-Credentials-Flow (App-Berechtigung, kein
angemeldeter Benutzer) mit den Zugangsdaten aus config/api_key.json. Das
Postfach, aus dem gesendet wird, steht dort als "sender_email".

Benoetigte Graph-Anwendungsberechtigung: Mail.Send (fuer den Direktversand)
bzw. Mail.ReadWrite (fuer das Anlegen von Entwuerfen). Beide brauchen eine
Administratorzustimmung im Azure-Portal.

Zwei Betriebsarten, gesteuert ueber config/einstellungen.json:

    "versand_modus": "entwurf"  Nachricht wird nur als Entwurf im Postfach
                                abgelegt (POST /users/{sender}/messages).
                                Es geht nichts an Kunden raus.
    "versand_modus": "direkt"   Nachricht wird versendet
                                (POST /users/{sender}/sendMail).

    "test_empfaenger_override": "adresse@example.com"
                                Leitet ALLE Kundenmails an diese eine Adresse
                                um und setzt "[TEST] " vor den Betreff. Fuer
                                Live-Tests, bevor scharf geschaltet wird.
                                null = Versand an die echten Kunden.

Die Fehlerberichtsmail an das eigene Postfach (sende_fehlerbericht) ist eine
interne Systembenachrichtigung und ignoriert beide Einstellungen bewusst:
sie wird immer direkt versendet und nie umgeleitet.
"""

import base64
import json
from datetime import date
from pathlib import Path
from urllib.parse import quote

import msal
import requests

PROJEKT_DIR = Path(__file__).resolve().parent.parent
MAIL_VORLAGE = PROJEKT_DIR / "templates" / "email_vorlage.txt"

GRAPH_BASIS = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
AUTHORITY_MUSTER = "https://login.microsoftonline.com/{tenant_id}"

HTTP_TIMEOUT = 60

VERSAND_MODI = ("entwurf", "direkt")
API_PFLICHTFELDER = ("tenant_id", "client_id", "client_secret", "sender_email")
# Platzhalter aus config/api_key.example.json - so belegte Felder gelten als
# noch nicht ausgefuellt.
PLATZHALTER_NICHT_AUSGEFUELLT = (
    "HIER_EINTRAGEN",
    "your-tenant-id",
    "your-client-id",
    "your-client-secret",
    "invoices@yourdomain.com",
)

# Graph nimmt Anhaenge nur bis rund 3 MB direkt in der Nachricht entgegen;
# groessere Dateien brauchen eine Upload-Session. Rechnungs-PDFs liegen weit
# darunter - die Pruefung ersetzt nur eine unverstaendliche Graph-Fehlermeldung.
MAX_ANHANG_BYTES = 3 * 1024 * 1024

TEST_PRAEFIX = "[TEST] "

# Eine ConfidentialClientApplication je Zugangsdatensatz und Prozess. msal
# haelt darin seinen Token-Cache; ein neues Objekt pro Mail wuerde fuer jede
# Rechnung ein neues Token anfordern.
_APP_CACHE = {}


class MailSendError(Exception):
    """Fehler beim Aufbau oder Versand einer Rechnungsmail."""


# ---------------------------------------------------------------- Konfiguration


def _pruefe_api_konfig(api_konfig: dict) -> dict:
    """Prueft config/api_key.json auf vollstaendig ausgefuellte Pflichtfelder.

    Returns:
        Dictionary mit den vier bereinigten Pflichtwerten.

    Raises:
        MailSendError: Wenn ein Feld fehlt, leer ist oder noch auf dem
            Platzhalter der Beispieldatei steht.
    """
    if not isinstance(api_konfig, dict):
        raise MailSendError("config/api_key.json enthaelt kein JSON-Objekt.")

    fehlend = []
    for feld in API_PFLICHTFELDER:
        wert = api_konfig.get(feld)
        if not isinstance(wert, str) or not wert.strip():
            fehlend.append(feld)
        elif wert.strip() in PLATZHALTER_NICHT_AUSGEFUELLT:
            fehlend.append(feld)

    if fehlend:
        raise MailSendError(
            "config/api_key.json ist unvollstaendig - noch nicht ausgefuellt: "
            + ", ".join(fehlend)
            + ".\nBitte die Werte der Azure-App-Registrierung eintragen "
            "(Vorlage: config/api_key.example.json)."
        )

    return {feld: api_konfig[feld].strip() for feld in API_PFLICHTFELDER}


def _lies_einstellungen(einstellungen: dict):
    """Liefert (versand_modus, test_empfaenger_override) aus einstellungen.json.

    Raises:
        MailSendError: Bei unbekanntem Versandmodus oder ungueltiger
            Override-Adresse.
    """
    if not isinstance(einstellungen, dict):
        raise MailSendError("config/einstellungen.json enthaelt kein JSON-Objekt.")

    modus = einstellungen.get("versand_modus")
    if modus not in VERSAND_MODI:
        raise MailSendError(
            f"Unbekannter 'versand_modus' in config/einstellungen.json: {modus!r}.\n"
            f"Erlaubt sind {' oder '.join(repr(m) for m in VERSAND_MODI)} "
            f"('entwurf' legt die Mail nur im Postfach ab, 'direkt' versendet sie)."
        )

    override = einstellungen.get("test_empfaenger_override")
    if override is not None:
        if not isinstance(override, str) or not override.strip():
            raise MailSendError(
                "'test_empfaenger_override' in config/einstellungen.json muss "
                "entweder null sein oder eine E-Mail-Adresse enthalten, ist aber "
                f"{override!r}."
            )
        override = override.strip()

    return modus, override


def pruefe_versandkonfiguration(einstellungen: dict, api_konfig: dict) -> None:
    """Prueft Einstellungen, Zugangsdaten und Mailvorlage vor dem Lauf.

    Wird von main.py aufgerufen, bevor die erste PDF entsteht, damit ein
    Tippfehler in einstellungen.json nicht erst nach der Rechnungserzeugung
    auffaellt. Es wird bewusst kein Token angefordert - das passiert erst beim
    ersten tatsaechlichen Versand.

    Raises:
        MailSendError: Bei ungueltigen Einstellungen, unvollstaendigen
            Zugangsdaten oder fehlender/fehlerhafter Mailvorlage.
    """
    _pruefe_api_konfig(api_konfig)
    _lies_einstellungen(einstellungen)
    lade_mailvorlage()


# -------------------------------------------------------------------- Vorlage


def lade_mailvorlage():
    """Liest templates/email_vorlage.txt und trennt Betreff von Nachrichtentext.

    Aufbau der Datei: eine Zeile "Betreff: ..." (optional von "---"-Zeilen
    eingerahmt), danach durch "---" abgetrennt der Nachrichtentext. Eine
    abschliessende "---"-Zeile gehoert zur Rahmung und wird verworfen.

    Returns:
        (betreff_vorlage, text_vorlage) - beide noch mit Platzhaltern.

    Raises:
        MailSendError: Wenn die Datei fehlt, keine Betreffzeile enthaelt oder
            der Nachrichtentext leer ist.
    """
    if not MAIL_VORLAGE.is_file():
        raise MailSendError(
            f"Mailvorlage nicht gefunden: {MAIL_VORLAGE}\n"
            f"Die Datei muss eine Zeile 'Betreff: ...' und - durch '---' "
            f"abgetrennt - den Nachrichtentext enthalten."
        )

    zeilen = MAIL_VORLAGE.read_text(encoding="utf-8").splitlines()

    betreff = None
    rest = []
    for nummer, zeile in enumerate(zeilen):
        if betreff is None:
            if zeile.strip() in ("", "---"):
                continue  # Rahmung oberhalb der Betreffzeile
            if not zeile.strip().lower().startswith("betreff:"):
                raise MailSendError(
                    f"{MAIL_VORLAGE.name}: erwartet wurde zuerst eine Zeile "
                    f"'Betreff: ...', gefunden wurde in Zeile {nummer + 1}: "
                    f"{zeile.strip()!r}"
                )
            betreff = zeile.split(":", 1)[1].strip()
            continue
        rest.append(zeile)

    if betreff is None:
        raise MailSendError(f"{MAIL_VORLAGE.name} enthaelt keine Zeile 'Betreff: ...'.")
    if not betreff:
        raise MailSendError(f"{MAIL_VORLAGE.name}: die Betreffzeile ist leer.")

    # Trennlinie unter dem Betreff und Rahmung am Dateiende entfernen.
    if rest and rest[0].strip() == "---":
        rest = rest[1:]
    while rest and not rest[-1].strip():
        rest.pop()
    if rest and rest[-1].strip() == "---":
        rest.pop()

    text = "\n".join(rest).strip("\n")
    if not text.strip():
        raise MailSendError(
            f"{MAIL_VORLAGE.name} enthaelt keinen Nachrichtentext unterhalb "
            f"der Betreffzeile."
        )

    return betreff, text


def _fuelle_vorlage(vorlage: str, werte: dict, was: str) -> str:
    """Setzt die Platzhalter einer Vorlage per .format() ein.

    Raises:
        MailSendError: Wenn die Vorlage einen unbekannten Platzhalter oder
            eine unvollstaendige geschweifte Klammer enthaelt.
    """
    try:
        return vorlage.format(**werte)
    except KeyError as exc:
        raise MailSendError(
            f"{MAIL_VORLAGE.name}: unbekannter Platzhalter im {was}: "
            f"{{{exc.args[0]}}}. Verfuegbar sind "
            + ", ".join("{" + name + "}" for name in sorted(werte))
            + "."
        ) from exc
    except (IndexError, ValueError) as exc:
        raise MailSendError(
            f"{MAIL_VORLAGE.name}: {was} enthaelt eine fehlerhafte geschweifte "
            f"Klammer ({exc}). Klammern, die als Text erscheinen sollen, muessen "
            f"verdoppelt werden."
        ) from exc


# --------------------------------------------------------------- Graph-Zugriff


def _hole_app(zugang: dict) -> msal.ConfidentialClientApplication:
    """Liefert die msal-App zu diesen Zugangsdaten, je Prozess nur einmal gebaut."""
    schluessel = (zugang["tenant_id"], zugang["client_id"])
    app = _APP_CACHE.get(schluessel)
    if app is None:
        try:
            app = msal.ConfidentialClientApplication(
                client_id=zugang["client_id"],
                client_credential=zugang["client_secret"],
                authority=AUTHORITY_MUSTER.format(tenant_id=zugang["tenant_id"]),
            )
        except ValueError as exc:
            # msal prueft die Authority-URL bereits beim Aufbau.
            raise MailSendError(
                f"Microsoft-Anmeldung konnte nicht vorbereitet werden - bitte "
                f"'tenant_id' in config/api_key.json pruefen: {exc}"
            ) from exc
        _APP_CACHE[schluessel] = app
    return app


def _hole_token(zugang: dict) -> str:
    """Besorgt ein Graph-Zugriffstoken im Client-Credentials-Flow.

    msal beantwortet den Aufruf aus seinem eigenen Cache, solange das Token
    gueltig ist; pro Lauf entsteht daher in der Regel nur eine Tokenanfrage.

    Raises:
        MailSendError: Wenn kein Token ausgestellt wird.
    """
    app = _hole_app(zugang)
    try:
        ergebnis = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    except requests.RequestException as exc:
        raise MailSendError(f"Microsoft-Anmeldedienst nicht erreichbar: {exc}") from exc

    token = (ergebnis or {}).get("access_token")
    if not token:
        fehler = (ergebnis or {}).get("error", "unbekannter Fehler")
        beschreibung = (ergebnis or {}).get("error_description", "")
        raise MailSendError(
            f"Microsoft Graph hat kein Zugriffstoken ausgestellt ({fehler}).\n"
            f"{beschreibung.strip()}\n"
            f"Bitte tenant_id, client_id und client_secret in config/api_key.json "
            f"pruefen - ein abgelaufenes Client-Secret ist die haeufigste Ursache."
        )
    return token


def _graph_fehlertext(antwort: requests.Response) -> str:
    """Formt aus einer Graph-Fehlerantwort eine lesbare Meldung."""
    try:
        fehler = (antwort.json() or {}).get("error", {})
    except (ValueError, json.JSONDecodeError):
        fehler = {}

    if fehler:
        text = f"{fehler.get('code', 'Fehler')}: {fehler.get('message', '')}".strip()
    else:
        text = (antwort.text or "").strip()[:500] or "keine Fehlerbeschreibung"

    if antwort.status_code in (401, 403):
        text += (
            "\nHinweis: Die Azure-App braucht die Anwendungsberechtigung "
            "Mail.Send (Direktversand) bzw. Mail.ReadWrite (Entwuerfe) inklusive "
            "Administratorzustimmung, und sender_email muss ein existierendes "
            "Postfach dieses Tenants sein."
        )
    return text


def _sende_an_graph(zugang: dict, pfad: str, nutzlast: dict, was: str) -> dict:
    """Fuehrt einen POST gegen Microsoft Graph aus und wertet die Antwort aus.

    Args:
        zugang: Geprueftes Zugangsdaten-Dictionary aus _pruefe_api_konfig().
        pfad: Graph-Pfad ab der Version, z. B. "/users/x@y.de/sendMail".
        nutzlast: JSON-Rumpf der Anfrage.
        was: Beschreibung fuer die Fehlermeldung, z. B. "Fehlerbericht senden".

    Returns:
        Die JSON-Antwort als dict; leeres dict, wenn Graph keinen Rumpf liefert
        (sendMail antwortet mit 202 ohne Inhalt).

    Raises:
        MailSendError: Bei Netzwerkfehlern oder HTTP-Status ausserhalb 2xx.
    """
    token = _hole_token(zugang)

    try:
        antwort = requests.post(
            f"{GRAPH_BASIS}{pfad}",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json=nutzlast,
            timeout=HTTP_TIMEOUT,
        )
    except requests.RequestException as exc:
        raise MailSendError(f"Microsoft Graph nicht erreichbar ({was}): {exc}") from exc

    if not 200 <= antwort.status_code < 300:
        raise MailSendError(
            f"Microsoft Graph hat den Vorgang abgelehnt ({was}, "
            f"HTTP {antwort.status_code}).\n{_graph_fehlertext(antwort)}"
        )

    try:
        return antwort.json() or {}
    except (ValueError, json.JSONDecodeError):
        return {}


# ------------------------------------------------------------------ Nachricht


def _baue_anhang(pdf_pfad: Path) -> dict:
    """Liest die Rechnungs-PDF und verpackt sie base64-codiert als Graph-Anhang.

    Raises:
        MailSendError: Wenn die Datei fehlt, leer, unlesbar oder zu gross ist.
    """
    pdf_pfad = Path(pdf_pfad)
    if not pdf_pfad.is_file():
        raise MailSendError(f"Rechnungs-PDF nicht gefunden: {pdf_pfad}")

    try:
        rohdaten = pdf_pfad.read_bytes()
    except OSError as exc:
        raise MailSendError(
            f"Rechnungs-PDF konnte nicht gelesen werden ({pdf_pfad}): {exc}"
        ) from exc

    if not rohdaten:
        raise MailSendError(f"Rechnungs-PDF ist leer: {pdf_pfad}")

    if len(rohdaten) > MAX_ANHANG_BYTES:
        raise MailSendError(
            f"Rechnungs-PDF ist mit {len(rohdaten) / 1024 / 1024:.1f} MB zu gross "
            f"fuer einen direkten Mailanhang (Grenze rund "
            f"{MAX_ANHANG_BYTES / 1024 / 1024:.0f} MB): {pdf_pfad.name}"
        )

    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": pdf_pfad.name,
        "contentType": "application/pdf",
        "contentBytes": base64.b64encode(rohdaten).decode("ascii"),
    }


def _baue_nachricht(betreff: str, text: str, empfaenger: str, anhaenge=None) -> dict:
    """Baut den Graph-Nachrichtenrumpf (Betreff, Text, Empfaenger, Anhaenge)."""
    nachricht = {
        "subject": betreff,
        "body": {"contentType": "Text", "content": text},
        "toRecipients": [{"emailAddress": {"address": empfaenger}}],
    }
    if anhaenge:
        nachricht["attachments"] = list(anhaenge)
    return nachricht


def _ermittle_empfaenger(kunde: dict, override):
    """Liefert (empfaenger, echter_empfaenger) - bei Umleitung sind sie verschieden.

    Raises:
        MailSendError: Wenn der Kundendatensatz keine Empfaengeradresse hat.
    """
    echter_empfaenger = (kunde.get("email_empfaenger") or "").strip()
    if not echter_empfaenger:
        kunde_id = kunde.get("kunde_id", "<ohne kunde_id>")
        raise MailSendError(
            f"Pflichtfeld 'email_empfaenger' fehlt oder ist leer im "
            f"Kundendatensatz {kunde_id!r} in config/kunden.json."
        )
    return (override or echter_empfaenger), echter_empfaenger


# -------------------------------------------------------------------- Versand


def sende_rechnung(
    kunde: dict,
    rechnungsnummer: str,
    pdf_pfad: Path,
    leistungszeitraum_von: date,
    leistungszeitraum_bis: date,
    einstellungen: dict,
    api_konfig: dict,
) -> None:
    """Versendet eine Rechnung mit PDF-Anhang - oder legt sie als Entwurf ab.

    Welches von beidem geschieht, entscheidet "versand_modus" in
    config/einstellungen.json. Ist "test_empfaenger_override" gesetzt, geht die
    Mail an diese Adresse statt an den Kunden und traegt "[TEST] " im Betreff.

    Args:
        kunde: Kundendatensatz aus config/kunden.json; "email_empfaenger" wird
            benoetigt (auch bei aktiver Umleitung, fuer die Protokollierung).
        rechnungsnummer: Vergebene Nummer, z. B. "RE202609001".
        pdf_pfad: Pfad zur erzeugten Rechnungs-PDF; wird angehaengt.
        leistungszeitraum_von: Erster Tag des Leistungszeitraums.
        leistungszeitraum_bis: Letzter Tag des Leistungszeitraums.
        einstellungen: Inhalt von config/einstellungen.json.
        api_konfig: Inhalt von config/api_key.json.

    Raises:
        MailSendError: Bei ungueltiger Konfiguration, fehlender PDF, fehlendem
            Token oder abgelehnter Graph-Anfrage.
    """
    zugang = _pruefe_api_konfig(api_konfig)
    modus, override = _lies_einstellungen(einstellungen)
    empfaenger, _echter_empfaenger = _ermittle_empfaenger(kunde, override)

    betreff_vorlage, text_vorlage = lade_mailvorlage()
    werte = {
        "rechnungsnummer": rechnungsnummer,
        "leistungszeitraum_von": f"{leistungszeitraum_von:%d.%m.%Y}",
        "leistungszeitraum_bis": f"{leistungszeitraum_bis:%d.%m.%Y}",
    }
    betreff = _fuelle_vorlage(betreff_vorlage, werte, "Betreff")
    text = _fuelle_vorlage(text_vorlage, werte, "Nachrichtentext")

    if override:
        # Umgeleitete Mails im Postfach sofort als Test erkennbar machen.
        betreff = f"{TEST_PRAEFIX}{betreff}"

    nachricht = _baue_nachricht(betreff, text, empfaenger, [_baue_anhang(pdf_pfad)])
    postfach = quote(zugang["sender_email"])

    if modus == "entwurf":
        _sende_an_graph(
            zugang,
            f"/users/{postfach}/messages",
            nachricht,
            f"Entwurf fuer {empfaenger} anlegen",
        )
    else:
        _sende_an_graph(
            zugang,
            f"/users/{postfach}/sendMail",
            {"message": nachricht, "saveToSentItems": True},
            f"Mail an {empfaenger} senden",
        )


def beschreibe_versand(kunde: dict, einstellungen: dict) -> str:
    """Beschreibt in einem Satz, was sende_rechnung() mit dieser Mail getan hat.

    Dient der Protokollierung in main.py: bei aktiver Testumleitung wird
    festgehalten, dass die Mail umgeleitet wurde und wer eigentlich Empfaenger
    gewesen waere.
    """
    modus, override = _lies_einstellungen(einstellungen)
    empfaenger, echter_empfaenger = _ermittle_empfaenger(kunde, override)

    tat = "als Entwurf abgelegt" if modus == "entwurf" else "versendet"
    if override:
        return (
            f"{tat} an {empfaenger} - UMGELEITETE TESTMAIL, eigentlicher "
            f"Empfaenger waere {echter_empfaenger} gewesen"
        )
    return f"{tat} an {empfaenger}"


def sende_fehlerbericht(
    fehler,
    api_konfig: dict,
    rechnungsdatum: date,
    log_datei: Path = None,
) -> str:
    """Meldet fehlgeschlagene Kunden per Mail an das eigene Postfach.

    Interne Systembenachrichtigung: sie wird bewusst immer direkt versendet und
    nie ueber test_empfaenger_override umgeleitet, unabhaengig davon, was in
    config/einstellungen.json steht. Absender und Empfaenger sind beide
    sender_email aus config/api_key.json.

    Args:
        fehler: Liste von (kundenname, phase, meldung); phase ist "PDF" oder
            "Mail".
        api_konfig: Inhalt von config/api_key.json.
        rechnungsdatum: Datum des Laufs, erscheint im Betreff.
        log_datei: Optionaler Pfad zum Laufprotokoll, wird im Text genannt.

    Returns:
        Die Adresse, an die der Bericht gesendet wurde.

    Raises:
        MailSendError: Wenn der Bericht nicht zugestellt werden kann. Der
            Aufrufer soll das protokollieren, aber den Lauf nicht abbrechen.
    """
    zugang = _pruefe_api_konfig(api_konfig)
    empfaenger = zugang["sender_email"]

    betreff = f"Rechnungslauf {rechnungsdatum:%d.%m.%Y}: {len(fehler)} Fehler"

    zeilen = [
        f"Der Rechnungslauf vom {rechnungsdatum:%d.%m.%Y} ist nicht vollstaendig "
        f"durchgelaufen.",
        "",
        f"Fehlgeschlagen ({len(fehler)}):",
        "",
    ]
    for name, phase, meldung in fehler:
        zeilen.append(f"- {name} [{phase}]")
        for zeile in str(meldung).splitlines():
            zeilen.append(f"    {zeile}")
        zeilen.append("")

    if log_datei is not None:
        zeilen.append(f"Vollstaendiges Protokoll: {log_datei}")
        zeilen.append("")
    zeilen.append("Diese Nachricht wurde automatisch vom Rechnungstool erzeugt.")

    nachricht = _baue_nachricht(betreff, "\n".join(zeilen), empfaenger)
    _sende_an_graph(
        zugang,
        f"/users/{quote(empfaenger)}/sendMail",
        {"message": nachricht, "saveToSentItems": True},
        "Fehlerbericht senden",
    )
    return empfaenger
