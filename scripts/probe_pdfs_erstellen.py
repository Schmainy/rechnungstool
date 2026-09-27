"""Probelauf der reinen PDF-Erzeugung mit drei erfundenen Testkunden.

Zweck: pruefen, ob unterschiedliche Textlaengen und Betragsgroessen sauber ins
Layout der Vorlage passen - besonders eine sehr lange Leistungsbeschreibung.

Bewusst unberuehrt bleiben:

* logs/rechnungsnummern_zaehler.json - die Rechnungsnummern werden hier aus dem
  Rechnungsmonat und der Position in der Kundenliste gebildet und tragen das
  Praefix "TEST" statt "RE"; rechnungsnummer.py wird nicht importiert.
* config/kunden.json - die Testkunden stehen unten im Skript.
* output/rechnungen/ - die PDFs entstehen in einem temporaeren Ordner und
  landen anschliessend flach im gewaehlten Probeordner im Projektroot.
* src/mail_sender.py und src/main.py - werden nicht importiert. Die Berechnung
  des Leistungszeitraums ist deshalb hier nachgebildet (identisch zu
  main.bestimme_leistungszeitraum); ein Import von main wuerde mail_sender
  mitziehen.

Aufrufe aus dem Projektroot:

    .\\venv\\Scripts\\python.exe scripts\\probe_pdfs_erstellen.py
    .\\venv\\Scripts\\python.exe scripts\\probe_pdfs_erstellen.py --ordner probeordner_2 --monate-voraus 1

Mit --monate-voraus laesst sich ein kuenftiger Abrechnungsmonat simulieren:
Rechnungsdatum = heute + N Monate, Leistungszeitraum = der volle Kalendermonat
davor, Rechnungsnummern entsprechend TEST{JJJJMM}{NNN}.
"""

import argparse
import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path

from dateutil.relativedelta import relativedelta

PROJEKT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJEKT_DIR))

from src import pdf_generator  # noqa: E402  (erst nach sys.path-Anpassung moeglich)
from src.pdf_generator import erstelle_rechnung  # noqa: E402

# Praefix der Testnummern - bewusst nicht "RE", damit eine Probe nie mit
# einer echten Rechnung verwechselt werden kann.
TEST_PRAEFIX = "TEST"

TESTKUNDEN = [
    {
        "kunde_id": "probe_nord",
        "firma": "Musterfirma Nord GmbH",
        "strasse": "Industrieweg 5",
        "plz_ort": "22525 Hamburg",
        "beschreibung": "Monatliche Wartung Website",
        "anzahl": 1,
        "einzelpreis": 450.00,
        "steuersatz": 19.0,
        "zahlungsziel_tage": 14,
    },
    {
        "kunde_id": "probe_baeckerei",
        "firma": "Bäckerei Sonnenschein e.K.",
        "strasse": "Marktplatz 12",
        "plz_ort": "04109 Leipzig",
        "beschreibung": (
            "Social-Media-Betreuung: Content-Erstellung, Postings und "
            "Community-Management für Instagram und Facebook"
        ),
        "anzahl": 1,
        "einzelpreis": 980.50,
        "steuersatz": 19.0,
        "zahlungsziel_tage": 30,
    },
    {
        "kunde_id": "probe_consulting",
        "firma": "Technologie & Innovation Consulting AG",
        "strasse": "Zeppelinstraße 88",
        "plz_ort": "70174 Stuttgart",
        "beschreibung": (
            "Umfassende digitale Transformationsberatung: Analyse "
            "bestehender IT-Infrastruktur, Entwicklung einer mehrjährigen "
            "Digitalisierungsstrategie, Schulung der Belegschaft sowie "
            "begleitende Prozessoptimierung über mehrere Abteilungen hinweg"
        ),
        "anzahl": 1,
        "einzelpreis": 4750.00,
        "steuersatz": 19.0,
        "zahlungsziel_tage": 14,
    },
]


def bestimme_leistungszeitraum(heute: date):
    """Liefert (erster_tag, letzter_tag) des zuletzt abgeschlossenen Monats.

    Nachbildung von main.bestimme_leistungszeitraum - siehe Modul-Docstring.
    """
    von = (heute - relativedelta(months=1)).replace(day=1)
    bis = von + relativedelta(day=31)  # klemmt auf den letzten Tag des Monats
    return von, bis


def testrechnungsnummer(rechnungsdatum: date, position: int) -> str:
    """Bildet eine feste Testnummer, z. B. TEST202610001 - ohne Zaehlerdatei."""
    return f"{TEST_PRAEFIX}{rechnungsdatum:%Y%m}{position:03d}"


def erzeuge_probe_pdf(rechnungsnummer, kunde, zielordner, heute, von, bis) -> Path:
    """Erzeugt eine Probe-Rechnung und legt sie flach im Zielordner ab.

    erstelle_rechnung() schreibt immer nach AUSGABE_DIR/{Jahr}/{Monat}. Damit
    der echte Ausgabebaum unberuehrt bleibt, zeigt AUSGABE_DIR waehrend des
    Aufrufs auf einen temporaeren Ordner; die fertige PDF wird von dort in den
    Probeordner verschoben.
    """
    original_ausgabe_dir = pdf_generator.AUSGABE_DIR
    tempordner = Path(tempfile.mkdtemp(prefix="probe_rechnungen_"))
    try:
        pdf_generator.AUSGABE_DIR = tempordner
        erzeugt = erstelle_rechnung(
            kunde=kunde,
            firmendaten={},  # von erstelle_rechnung derzeit nicht ausgewertet
            rechnungsnummer=rechnungsnummer,
            rechnungsdatum=heute,
            zeitraum_von=von,
            zeitraum_bis=bis,
        )
        ziel = zielordner / erzeugt.name
        shutil.move(str(erzeugt), str(ziel))
        return ziel
    finally:
        pdf_generator.AUSGABE_DIR = original_ausgabe_dir
        shutil.rmtree(tempordner, ignore_errors=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--ordner",
        default="probeordner_1",
        help="Zielordner im Projektroot (Standard: probeordner_1)",
    )
    parser.add_argument(
        "--monate-voraus",
        type=int,
        default=0,
        metavar="N",
        help="Rechnungsdatum um N Monate in die Zukunft verschieben "
             "(Standard: 0 = heute)",
    )
    argumente = parser.parse_args()

    heute = date.today() + relativedelta(months=argumente.monate_voraus)
    von, bis = bestimme_leistungszeitraum(heute)
    zielordner = PROJEKT_DIR / argumente.ordner
    zielordner.mkdir(parents=True, exist_ok=True)

    print(f"Rechnungsdatum:    {heute:%d.%m.%Y}"
          + (f"  (heute + {argumente.monate_voraus} Monat(e))"
             if argumente.monate_voraus else "  (heute)"))
    print(f"Leistungszeitraum: {von:%d.%m.%Y} – {bis:%d.%m.%Y}")
    print(f"Zielordner:        {zielordner}")
    print("-" * 78)

    fehler = 0
    for position, kunde in enumerate(TESTKUNDEN, start=1):
        rechnungsnummer = testrechnungsnummer(heute, position)
        netto = float(kunde["anzahl"]) * float(kunde["einzelpreis"])
        brutto = netto * (1 + float(kunde["steuersatz"]) / 100)
        try:
            pfad = erzeuge_probe_pdf(rechnungsnummer, kunde, zielordner, heute, von, bis)
        except Exception as exc:
            fehler += 1
            print(f"FEHLER  {kunde['firma']}: {exc}")
            continue

        print(f"OK      {rechnungsnummer}  {kunde['firma']}")
        print(f"        Beschreibung:  {len(kunde['beschreibung'])} Zeichen")
        print(f"        Netto/Brutto:  {netto:.2f} / {brutto:.2f} EUR")
        print(f"        Zahlungsziel:  {kunde['zahlungsziel_tage']} Tage")
        print(f"        Datei:         {pfad.name} ({pfad.stat().st_size} Bytes)")

    print("-" * 78)
    print(f"{len(TESTKUNDEN) - fehler} von {len(TESTKUNDEN)} Probe-PDFs erzeugt.")
    return 1 if fehler else 0


if __name__ == "__main__":
    sys.exit(main())
