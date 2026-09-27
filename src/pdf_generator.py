"""PDF-Erzeugung der Rechnungen.

Ablauf: Excel-Vorlage kopieren -> Kopie mit openpyxl befuellen ->
per LibreOffice (headless) nach PDF konvertieren -> PDF archivieren.


Aufbau der Vorlage templates/Musterrechnung.xlsx (Blatt "RG Muster")
-------------------------------------------------------------------
Zielzellen der Vorlage (Beispielinhalte wie in Musterrechnung_EXAMPLE.xlsx):

    Zelle  Inhalt in der Vorlage        Art              Zellformat
    F7     "info@example.com"           statisch (Text)  Zellstil "Link" + mailto:
    B15    "Musterfirma GmbH"           statisch (Text)  General
    B16    "Beispielweg 2"              statisch (Text)  General
    B17    "10117 Berlin"               statisch (Text)  General
    F23    31.01.2026                   statisch (Datum) [$-407]d/ mmmm yyyy;@
    F24    "RE202601500"                statisch (Text)  mmm-yy (wirkt bei Text nicht)
    F25    "Musterfirma GmbH"           statisch (Text)  mmm-yy (wirkt bei Text nicht)
    B25    "01.01.2026 - 31.01.2026"    statisch (Text)  General
    B28    Leistungsbeschreibung        statisch (Text)  0.00 (wirkt bei Text nicht)
    C28    1                            statisch (Zahl)  0.00
    D28    1000                         statisch (Zahl)  EUR-Waehrungsformat
    E28    =D28*C28                     FORMEL           EUR-Buchhaltungsformat
    E29    =E28                         FORMEL           EUR-Buchhaltungsformat
    E30    "19.00%"                     statisch (Text)  0% (wirkt bei Text nicht)
    E31    =E29*0.19                    FORMEL           EUR-Buchhaltungsformat
    E32    =E29+E31                     FORMEL           EUR-Buchhaltungsformat
    B34    Zahlungshinweis (14 Tage)    statisch (Text)  verbunden B34:F35

Folgerungen fuer dieses Modul:

* E28, E29, E31 und E32 sind in der Vorlage Formeln. openpyxl schreibt keine
  Formelergebnisse; eine nur befuellte Kopie enthielte an diesen Stellen
  unberechnete Formeln, und der Steuersatz waere ausserdem fest auf 0.19
  verdrahtet (E31). Deshalb werden Zwischensumme, Umsatzsteuer und Summe hier
  in Python gerechnet und als reine Zahlenwerte geschrieben. Die PDF-Ausgabe
  ist damit unabhaengig davon, ob das konvertierende Programm Formeln neu
  berechnet.
* E30 ist in der Vorlage Text ("19.00%"), nicht die Zahl 0.19 - das Zellformat
  0% greift daher nicht. Der Steuersatz wird in derselben Schreibweise
  geschrieben, damit die Darstellung unveraendert bleibt.
* F7 traegt in der Vorlage den benannten Zellstil "Link" und ein
  mailto:-Hyperlinkobjekt. LibreOffice rendert das als blauen, unterstrichenen
  Text mit klickbarer Verknuepfung - im gedruckten Original ist die Adresse
  aber schwarz wie die uebrigen Kopfzeilen. _entschaerfe_hyperlink_optik()
  entfernt darum Hyperlink, Linkfarbe und Unterstreichung aus der Kopie.
* Zeile 28 hat eine feste Hoehe (customHeight). Laeuft die Beschreibung in
  B28 ueber diese Hoehe hinaus, schneidet LibreOffice sie beim Export still
  ab - und weil B28 vertikal zentriert ist, oben *und* unten. Vor dem Export
  wird darum gemessen, wie viele Zeilen der Text tatsaechlich braucht; passt
  er nicht, bricht _pruefe_beschreibungshoehe() mit BeschreibungZuLangError
  ab, statt eine beschnittene Rechnung zu erzeugen.
* B34 enthaelt in der Vorlage einen festen Zahlungshinweis mit "14 Tagen".
  Das Zahlungsziel steht je Kunde in kunden.json ("zahlungsziel_tage"), der
  Text wird deshalb hier neu erzeugt und geschrieben. B34 ist der Ankerpunkt
  des verbundenen Bereichs B34:F35; das Setzen von .value laesst Verbund,
  Schriftart (Poppins 10) und Zeilenumbruch unveraendert.
* F24/F25/B28/C28 tragen Zahlenformate, die auf ihren tatsaechlichen
  Textinhalt nicht wirken; sie werden unveraendert uebernommen.
* Es wird ausschliesslich .value gesetzt. Zellformate, Farben, Rahmen,
  Spaltenbreiten und das eingebettete Logo bleiben dadurch unberuehrt.
  Fuer das Logo muss Pillow installiert sein: ohne Pillow liest openpyxl
  eingebettete Bilder nicht ein und verwirft sie beim Speichern.
"""

import os
import shutil
import subprocess
import tempfile
from copy import copy
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl
from openpyxl.styles import Color
from PIL import ImageFont

PROJEKT_DIR = Path(__file__).resolve().parent.parent
VORLAGE = PROJEKT_DIR / "templates" / "Musterrechnung.xlsx"
AUSGABE_DIR = PROJEKT_DIR / "output" / "rechnungen"

# Textfarbe, mit der Hyperlinkzellen ueberschrieben werden - dasselbe Schwarz
# wie die uebrigen Kopfzeilen der Vorlage (F5, F6, F8).
SCHRIFTFARBE_STANDARD = "FF000000"

# Benannter Zellstil, den Excel Hyperlinkzellen zuweist (deutsche wie englische
# Oberflaeche); er bringt Blau und Unterstreichung mit.
LINK_ZELLSTILE = ("Link", "Hyperlink", "Followed Hyperlink", "Besuchter Link")

# --- Masse fuer die Laengenpruefung der Leistungsbeschreibung --------------
#
# Alle Werte in Excel-Punkten, also unskaliert. LibreOffice verkleinert die
# Seite beim PDF-Export auf rund 78 %, was Schriftgroesse, Zeilenhoehe und
# Spaltenbreite gleichermassen betrifft - fuer den Vergleich "benoetigte
# gegen verfuegbare Hoehe" kuerzt sich der Faktor heraus.
BESCHREIBUNG_ZEILE = 28
BESCHREIBUNG_ZELLE = "B28"

# Pillow rundet Textbreiten auf ganze Pixel der angeforderten Schriftgroesse.
# Bei 11 pt waere das viel zu grob, darum wird mit dem Zehnfachen gemessen und
# das Ergebnis wieder geteilt - das liegt auf 0.5 % an die Breiten heran, die
# LibreOffice tatsaechlich setzt.
MESS_UEBERABTASTUNG = 10

# Nutzbare Textbreite in B28. Aus der Excel-Spaltenbreite laesst sie sich nicht
# zuverlaessig herleiten (die Standardformel liefert 240.7 pt und damit weniger
# als die breiteste Zeile, die LibreOffice tatsaechlich setzt). Der Wert wurde
# stattdessen aus den Umbruechen erzeugter PDFs eingegrenzt: jede gesetzte
# Zeile ist schmaler, jede Zeile samt naechstem Wort breiter. Mit den
# Pillow-Messwerten ergibt das das Intervall 244.9 - 255.4 pt; genommen ist
# dessen Mitte, die die Umbrueche aller bekannten Faelle korrekt reproduziert.
BESCHREIBUNG_BREITE_PT = 250.15

# Zeilenabstand als Vielfaches der Schriftgroesse, gemessen an Poppins 11 pt
# (12.0 pt im PDF bei 78 % Skalierung entsprechen 15.35 pt unskaliert). Der
# Wert stammt aus dem fertigen PDF, nicht aus einer Schriftmessung, und ist
# darum unabhaengig von der verwendeten Messbibliothek. Die Font-Metriken von
# Pillow (ascent + descent = 1.455) beschreiben etwas anderes: den maximalen
# Platzbedarf der Schrift, nicht den Zeilenabstand, den LibreOffice setzt.
ZEILENABSTAND_FAKTOR = 1.395

# Luft, die oben und unten in der Zelle frei bleiben soll. Der als sauber
# bewertete Referenzfall hat rund 3.6 pt je Seite.
BESCHREIBUNG_MARGE_PT = 3.0

# Bekannte Installationsorte von LibreOffice unter Windows.
SOFFICE_STANDARDPFADE = (
    Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
)

INSTALL_HINWEIS = (
    "LibreOffice wird fuer die Umwandlung der Excel-Rechnung in ein PDF benoetigt.\n"
    "Installation:\n"
    "    winget install --id TheDocumentFoundation.LibreOffice\n"
    "oder Download unter https://www.libreoffice.org/download/\n"
    "Danach liegt soffice.exe ueblicherweise unter\n"
    r"    C:\Program Files\LibreOffice\program\soffice.exe"
)

# Zeitlimit fuer einen Konvertierungslauf. LibreOffice legt beim ersten
# Aufruf ein neues Benutzerprofil an und braucht dann deutlich laenger.
KONVERTIERUNG_TIMEOUT = 180


class PDFGeneratorError(Exception):
    """Fehler bei der Erzeugung einer Rechnung."""


class LibreOfficeNichtGefunden(PDFGeneratorError):
    """LibreOffice ist nicht installiert oder nicht auffindbar."""


class BeschreibungZuLangError(PDFGeneratorError):
    """Die Leistungsbeschreibung passt nicht in die Positionszelle der Vorlage.

    Erbt von PDFGeneratorError und wird dadurch in main.py wie jeder andere
    Fehler der PDF-Erstellung behandelt: der Kunde wird uebersprungen,
    protokolliert und in den Sammelbericht aufgenommen.
    """


def pruefe_vorlage() -> Path:
    """Stellt sicher, dass die Excel-Vorlage vorhanden ist.

    Raises:
        PDFGeneratorError: Wenn templates/Musterrechnung.xlsx fehlt.
    """
    if not VORLAGE.is_file():
        raise PDFGeneratorError(
            f"Rechnungsvorlage nicht gefunden: {VORLAGE}\n"
            f"Bitte die Musterrechnung als 'Musterrechnung.xlsx' im Ordner "
            f"{VORLAGE.parent} ablegen - z. B. Musterrechnung_EXAMPLE.xlsx kopieren "
            f"und anpassen (siehe templates/README_TEMPLATE.txt)."
        )
    return VORLAGE


def finde_soffice() -> Path:
    """Sucht die LibreOffice-Programmdatei.

    Reihenfolge: PATH, danach die bekannten Standardinstallationspfade.

    Raises:
        LibreOfficeNichtGefunden: Wenn soffice.exe nirgends gefunden wird.
    """
    im_pfad = shutil.which("soffice") or shutil.which("soffice.exe")
    if im_pfad:
        return Path(im_pfad)

    for pfad in SOFFICE_STANDARDPFADE:
        if pfad.is_file():
            return pfad

    raise LibreOfficeNichtGefunden(
        "LibreOffice (soffice.exe) wurde nicht gefunden - weder im PATH noch unter "
        + " oder ".join(str(p) for p in SOFFICE_STANDARDPFADE)
        + ".\n\n"
        + INSTALL_HINWEIS
    )


def pruefe_voraussetzungen() -> None:
    """Prueft Vorlage und LibreOffice, bevor der Rechnungslauf startet."""
    pruefe_vorlage()
    finde_soffice()


def _euro(betrag) -> float:
    """Rundet einen Geldbetrag kaufmaennisch auf zwei Nachkommastellen."""
    gerundet = Decimal(str(betrag)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return float(gerundet)


def _zeitraum_text(von: date, bis: date) -> str:
    """Formatiert den Leistungszeitraum als 'TT.MM.JJJJ \u2013 TT.MM.JJJJ'."""
    return f"{von:%d.%m.%Y} \u2013 {bis:%d.%m.%Y}"


def _zahlungsziel_text(zahlungsziel_tage: int) -> str:
    """Baut den Zahlungshinweis fuer Zelle B34 aus dem Zahlungsziel des Kunden."""
    return (
        "Bitte verwenden Sie als Verwendungszweck Ihre Rechnungsnummer. "
        f"Bitte überweisen Sie den Betrag innerhalb von {zahlungsziel_tage} "
        "Tagen nach Rechnungserhalt."
    )


def _dateiname(rechnungsdatum: date, rechnungsnummer: str) -> str:
    """Dateiname im Schema des Muster-PDFs, z. B. 26.09.02_RechnungRE202609001."""
    return f"{rechnungsdatum:%y.%m.%d}_Rechnung{rechnungsnummer}"


def _befuelle_arbeitsmappe(
    xlsx_pfad: Path,
    kunde: dict,
    rechnungsnummer: str,
    rechnungsdatum: date,
    zeitraum_von: date,
    zeitraum_bis: date,
) -> None:
    """Schreibt alle Rechnungsdaten in die Kopie der Vorlage.

    Es wird nur .value gesetzt, damit saemtliche Formatierungen der Vorlage
    erhalten bleiben. Die Betraege sind bereits fertig gerechnete Zahlen.
    """
    kunde_id = kunde.get("kunde_id", "<ohne kunde_id>")

    for feld in ("firma", "strasse", "plz_ort", "beschreibung"):
        if not kunde.get(feld):
            raise PDFGeneratorError(
                f"Pflichtfeld {feld!r} fehlt oder ist leer im Kundendatensatz "
                f"{kunde_id!r} in config/kunden.json."
            )

    try:
        anzahl = float(kunde["anzahl"])
        einzelpreis = float(kunde["einzelpreis"])
        steuersatz = float(kunde["steuersatz"])
        zahlungsziel_tage = int(kunde["zahlungsziel_tage"])
    except KeyError as exc:
        raise PDFGeneratorError(
            f"Pflichtfeld {exc.args[0]!r} fehlt im Kundendatensatz {kunde_id!r} "
            f"in config/kunden.json."
        ) from exc
    except (TypeError, ValueError) as exc:
        raise PDFGeneratorError(
            f"'anzahl', 'einzelpreis', 'steuersatz' und 'zahlungsziel_tage' "
            f"muessen Zahlen sein (Kunde {kunde_id!r}): {exc}"
        ) from exc

    if zahlungsziel_tage < 1:
        raise PDFGeneratorError(
            f"'zahlungsziel_tage' muss mindestens 1 sein (Kunde {kunde_id!r}), "
            f"ist aber {zahlungsziel_tage}."
        )

    # Betraege selbst rechnen (siehe Modul-Docstring): keine Formeln schreiben.
    zwischensumme = _euro(anzahl * einzelpreis)
    umsatzsteuer = _euro(zwischensumme * steuersatz / 100)
    summe = _euro(zwischensumme + umsatzsteuer)

    wb = openpyxl.load_workbook(xlsx_pfad)
    ws = wb.active

    # Empfaengeranschrift
    ws["B15"] = kunde["firma"]
    ws["B16"] = kunde["strasse"]
    ws["B17"] = kunde["plz_ort"]

    # Rechnungskopf
    ws["F23"] = rechnungsdatum
    ws["F24"] = rechnungsnummer
    ws["F25"] = kunde["firma"]
    ws["B25"] = _zeitraum_text(zeitraum_von, zeitraum_bis)

    # Position - vor dem Schreiben pruefen, ob der Text in die Zelle passt.
    _pruefe_beschreibungshoehe(
        ws, kunde["beschreibung"], kunde_id, kunde.get("firma", kunde_id)
    )
    ws["B28"] = kunde["beschreibung"]
    ws["C28"] = anzahl
    ws["D28"] = einzelpreis
    ws["E28"] = zwischensumme

    # Summenblock - als Zahlen statt der Formeln aus der Vorlage
    ws["E29"] = zwischensumme
    ws["E30"] = f"{steuersatz:.2f}%"  # Vorlage haelt den Steuersatz als Text
    ws["E31"] = umsatzsteuer
    ws["E32"] = summe

    # Zahlungshinweis - ersetzt den festen "14 Tage"-Text der Vorlage.
    ws["B34"] = _zahlungsziel_text(zahlungsziel_tage)

    # Mailadresse im Kopf soll schwarz und ohne Unterstreichung erscheinen.
    _entschaerfe_hyperlink_optik(ws)

    wb.save(xlsx_pfad)


def _schriftdatei(schriftname: str):
    """Sucht die TrueType-Datei einer Schrift in den Font-Ordnern von Windows.

    Durchsucht zuerst die Schriften des Benutzers (dort liegt Poppins, wenn es
    ohne Administratorrechte installiert wurde), danach die Systemschriften.

    Returns:
        Pfad zur .ttf-Datei oder None, wenn die Schrift nicht gefunden wird.
    """
    ordner = []
    lokal = os.environ.get("LOCALAPPDATA")
    if lokal:
        ordner.append(Path(lokal) / "Microsoft" / "Windows" / "Fonts")
    ordner.append(Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts")

    stamm = (schriftname or "").replace(" ", "")
    if not stamm:
        return None

    for verzeichnis in ordner:
        for dateiname in (f"{stamm}-Regular.ttf", f"{stamm}.ttf"):
            pfad = verzeichnis / dateiname
            if pfad.is_file():
                return pfad
    return None


def _textbreite(font, text: str) -> float:
    """Breite eines Textes in Punkten, gemessen mit der ueberabgetasteten Schrift."""
    return font.getlength(text) / MESS_UEBERABTASTUNG


def _umbreche_text(text: str, font, breite_pt: float):
    """Bricht Text so um, wie es die Zelle mit Zeilenumbruch tut.

    Beruecksichtigt vorhandene Zeilenumbrueche im Text und trennt einzelne
    Woerter, die fuer sich genommen breiter als die Zelle sind.

    Args:
        text: Die Leistungsbeschreibung.
        font: Pillow-Schrift, angelegt mit dem MESS_UEBERABTASTUNG-fachen der
            tatsaechlichen Schriftgroesse.
        breite_pt: Nutzbare Zellbreite in Punkten.

    Returns:
        Liste der gesetzten Zeilen.
    """
    zeilen = []
    for absatz in text.split("\n"):
        aktuell = ""
        for wort in absatz.split():
            probe = f"{aktuell} {wort}".strip()
            if aktuell and _textbreite(font, probe) > breite_pt:
                zeilen.append(aktuell)
                aktuell = wort
            else:
                aktuell = probe

            # Ein einzelnes ueberlanges Wort bricht innerhalb des Wortes um.
            while _textbreite(font, aktuell) > breite_pt:
                rest = aktuell
                passend = ""
                for zeichen in rest:
                    if _textbreite(font, passend + zeichen) > breite_pt:
                        break
                    passend += zeichen
                if not passend:  # Sicherheitsnetz gegen eine Endlosschleife
                    break
                zeilen.append(passend)
                aktuell = rest[len(passend):]
        zeilen.append(aktuell)
    return zeilen


def _pruefe_beschreibungshoehe(ws, beschreibung: str, kunde_id, firma) -> None:
    """Stellt sicher, dass die Leistungsbeschreibung in die Positionszelle passt.

    Gemessen wird mit Pillow anhand der Schrift, die die Vorlage fuer B28
    vorgibt. Uebersprungen wird die Pruefung, wenn Zeile 28 keine feste Hoehe
    hat (dann waechst die Zeile mit) oder wenn die Schriftdatei nicht gefunden
    wird - ohne sie waere jede Messung geraten.

    Raises:
        BeschreibungZuLangError: Wenn der Text mehr Zeilen braucht, als die
            feste Zeilenhoehe abzueglich der Sicherheitsmarge hergibt.
    """
    zeilenhoehe = ws.row_dimensions[BESCHREIBUNG_ZEILE].height
    if not zeilenhoehe:
        return

    schrift = ws[BESCHREIBUNG_ZELLE].font
    schriftdatei = _schriftdatei(schrift.name)
    if schriftdatei is None:
        return

    schriftgroesse = float(schrift.size or 11.0)
    try:
        font = ImageFont.truetype(
            str(schriftdatei), size=int(schriftgroesse * MESS_UEBERABTASTUNG)
        )
    except OSError:
        # Schriftdatei vorhanden, aber nicht lesbar - lieber nicht messen als
        # den Rechnungslauf an der Pruefung scheitern lassen.
        return

    zeilen = _umbreche_text(beschreibung, font, BESCHREIBUNG_BREITE_PT)
    zeilenabstand = schriftgroesse * ZEILENABSTAND_FAKTOR
    benoetigt = len(zeilen) * zeilenabstand
    verfuegbar = zeilenhoehe - 2 * BESCHREIBUNG_MARGE_PT

    if benoetigt <= verfuegbar:
        return

    moegliche_zeilen = max(1, int(verfuegbar // zeilenabstand))
    zeichen_je_zeile = max(1, round(len(beschreibung) / len(zeilen)))
    empfohlene_laenge = moegliche_zeilen * zeichen_je_zeile

    raise BeschreibungZuLangError(
        f"Die Leistungsbeschreibung fuer {firma!r} (Kunde {kunde_id!r}) passt nicht "
        f"in die Positionszelle der Rechnungsvorlage.\n"
        f"Sie belegt {len(zeilen)} Zeilen ({len(beschreibung)} Zeichen); bei der "
        f"festen Hoehe von Zeile {BESCHREIBUNG_ZEILE} ({zeilenhoehe:.0f} pt) sind "
        f"hoechstens {moegliche_zeilen} Zeilen darstellbar - das sind rund "
        f"{empfohlene_laenge} Zeichen.\n"
        f"Laengerer Text wuerde im PDF oben und unten abgeschnitten, ohne dass "
        f"LibreOffice das meldet. Bitte die Beschreibung in config/kunden.json "
        f"kuerzen oder die Hoehe von Zeile {BESCHREIBUNG_ZEILE} in "
        f"templates/Musterrechnung.xlsx vergroessern."
    )


def _entschaerfe_hyperlink_optik(ws) -> None:
    """Nimmt allen Hyperlinkzellen Verknuepfung, Linkfarbe und Unterstreichung.

    Betrifft in der aktuellen Vorlage nur F7 (Mailadresse im Kopf), wird aber
    bewusst ueber das ganze Blatt angewandt: sollte in einer spaeteren Fassung
    der Vorlage eine weitere Adresse oder URL verlinkt sein, faellt sie damit
    nicht optisch aus dem Rahmen.

    Erfasst werden Zellen mit Hyperlinkobjekt *und* Zellen, die lediglich einen
    Link-Zellstil tragen - der genuegt fuer die blaue Unterstreichung bereits.
    Geaendert werden nur Schriftfarbe und Unterstreichung; Schriftart, Groesse,
    Fettung, Ausrichtung, Rahmen und Zahlenformat der Vorlage bleiben erhalten.
    """
    for zeile in ws.iter_rows():
        for zelle in zeile:
            if zelle.hyperlink is None and zelle.style not in LINK_ZELLSTILE:
                continue

            # Klickbare Verknuepfung entfernen - sonst zeichnet LibreOffice
            # weiterhin eine Link-Annotation ins PDF.
            zelle.hyperlink = None

            # copy(): zelle.font ist ein gemeinsam genutzter StyleProxy, der
            # sich nicht direkt aendern laesst.
            schrift = copy(zelle.font)
            schrift.color = Color(rgb=SCHRIFTFARBE_STANDARD)
            schrift.underline = "none"
            zelle.font = schrift

            # Der benannte Zellstil "Link" bleibt bewusst stehen. Ihn auf
            # "Standard" zu setzen wuerde zwar ebenfalls die Linkfarbe
            # entfernen, dabei aber die Ausrichtung der Zelle mitloeschen -
            # F7 ist rechtsbuendig, linksbuendig laeuft die Adresse ueber den
            # Seitenrand hinaus und wird abgeschnitten. Die direkt auf der
            # Zelle gesetzte Schrift ueberstimmt den Zellstil ohnehin.


def _konvertiere_zu_pdf(xlsx_pfad: Path, zielordner: Path) -> Path:
    """Wandelt die Arbeitsmappe per LibreOffice in ein PDF im Zielordner um."""
    soffice = finde_soffice()
    zielordner.mkdir(parents=True, exist_ok=True)
    erwartetes_pdf = zielordner / f"{xlsx_pfad.stem}.pdf"

    befehl = [
        str(soffice),
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(zielordner),
        str(xlsx_pfad),
    ]

    try:
        ergebnis = subprocess.run(
            befehl,
            capture_output=True,
            text=True,
            timeout=KONVERTIERUNG_TIMEOUT,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise PDFGeneratorError(
            f"LibreOffice hat die Konvertierung von {xlsx_pfad.name} nicht innerhalb "
            f"von {KONVERTIERUNG_TIMEOUT} Sekunden abgeschlossen. Laeuft noch eine "
            f"LibreOffice-Instanz mit offenem Dialog?"
        ) from exc

    # LibreOffice meldet Fehler nicht zuverlaessig ueber den Rueckgabewert,
    # daher wird zusaetzlich das Vorhandensein der PDF-Datei geprueft.
    if not erwartetes_pdf.is_file():
        raise PDFGeneratorError(
            f"LibreOffice hat kein PDF erzeugt (Rueckgabewert {ergebnis.returncode}).\n"
            f"Befehl: {' '.join(befehl)}\n"
            f"Ausgabe: {(ergebnis.stdout or '').strip()}\n"
            f"Fehler: {(ergebnis.stderr or '').strip()}"
        )

    return erwartetes_pdf


def erstelle_rechnung(
    kunde: dict,
    firmendaten: dict,
    rechnungsnummer: str,
    rechnungsdatum: date,
    zeitraum_von: date,
    zeitraum_bis: date,
) -> Path:
    """Erzeugt die Rechnung fuer einen Kunden als PDF.

    Args:
        kunde: Kundendatensatz aus config/kunden.json.
        firmendaten: Eigene Stammdaten aus config/firmendaten.json. Werden
            derzeit nicht in die Vorlage geschrieben, da diese die eigenen
            Firmen-, Bank- und Steuerdaten bereits fest enthaelt.
        rechnungsnummer: Bereits vergebene Nummer, z. B. "RE202609001".
        rechnungsdatum: Datum im Rechnungskopf.
        zeitraum_von: Erster Tag des Leistungszeitraums.
        zeitraum_bis: Letzter Tag des Leistungszeitraums.

    Returns:
        Pfad zur erzeugten PDF-Datei unter output/rechnungen/{Jahr}/{Monat}/.

    Raises:
        PDFGeneratorError: Bei fehlender Vorlage, unvollstaendigen Kundendaten
            oder fehlgeschlagener Konvertierung.
        LibreOfficeNichtGefunden: Wenn soffice.exe nicht auffindbar ist.
    """
    vorlage = pruefe_vorlage()
    finde_soffice()  # frueh pruefen, bevor Dateien angelegt werden

    zielordner = AUSGABE_DIR / f"{rechnungsdatum:%Y}" / f"{rechnungsdatum:%m}"
    basisname = _dateiname(rechnungsdatum, rechnungsnummer)

    arbeitsordner = Path(tempfile.mkdtemp(prefix="rechnung_"))
    arbeitsdatei = arbeitsordner / f"{basisname}.xlsx"
    try:
        # copy2 statt copy: die Vorlage selbst bleibt unangetastet.
        shutil.copy2(vorlage, arbeitsdatei)
        _befuelle_arbeitsmappe(
            arbeitsdatei,
            kunde,
            rechnungsnummer,
            rechnungsdatum,
            zeitraum_von,
            zeitraum_bis,
        )
        pdf_pfad = _konvertiere_zu_pdf(arbeitsdatei, zielordner)
    finally:
        # Temporaere xlsx-Kopie und Arbeitsordner in jedem Fall aufraeumen.
        shutil.rmtree(arbeitsordner, ignore_errors=True)

    return pdf_pfad
