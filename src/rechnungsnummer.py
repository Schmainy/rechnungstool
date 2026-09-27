"""Vergabe und Verwaltung fortlaufender Rechnungsnummern.

Format: RE{JJJJMM}{laufende_nummer:03d}, z. B. RE202609500

Die laufende Nummer wird gemeinsam ueber alle Kunden hochgezaehlt und in
logs/rechnungsnummern_zaehler.json persistiert. Struktur der Datei:

    {"2026-08": 503, "2026-09": 502}

Pro neuem Monatsschluessel beginnt die Zaehlung bei 500 (RE202609500,
RE202609501, ...). Der Abstand zum Anfang des Nummernkreises haelt die
unteren Nummern fuer von Hand geschriebene Rechnungen frei, damit sich
manuelle und automatische Rechnungsnummern nicht ins Gehege kommen.
Bereits vorhandene Monatsschluessel behalten ihren Stand und zaehlen von
dort aus weiter.

Jeder Aufruf von naechste_rechnungsnummer() zaehlt um eins hoch und
schreibt den Stand sofort zurueck, damit auch mehrere Aufrufe innerhalb
eines Laufs (und abgebrochene Laeufe) keine Nummer doppelt vergeben.
"""

import json
import os
import tempfile
from datetime import date
from pathlib import Path

LOG_DIR = Path(__file__).resolve().parent.parent / "logs"
ZAEHLER_DATEI = LOG_DIR / "rechnungsnummern_zaehler.json"

PRAEFIX = "RE"
# Erste Nummer eines neuen Monats; 001-499 bleiben manuellen Rechnungen vorbehalten.
START_LAUFENDE_NUMMER = 500
MAX_LAUFENDE_NUMMER = 999


class RechnungsnummerError(Exception):
    """Fehler bei der Vergabe einer Rechnungsnummer."""


def _lade_zaehler():
    """Liest den Zaehlerstand. Fehlende Datei ergibt einen leeren Stand."""
    if not ZAEHLER_DATEI.is_file():
        return {}
    try:
        with ZAEHLER_DATEI.open(encoding="utf-8") as f:
            daten = json.load(f)
    except json.JSONDecodeError as exc:
        raise RechnungsnummerError(
            f"Zaehlerdatei {ZAEHLER_DATEI} ist beschaedigt "
            f"(Zeile {exc.lineno}, Spalte {exc.colno}): {exc.msg}\n"
            f"Bitte den letzten gueltigen Stand von Hand wiederherstellen, damit "
            f"keine Rechnungsnummer doppelt vergeben wird."
        ) from exc
    if not isinstance(daten, dict):
        raise RechnungsnummerError(
            f"Zaehlerdatei {ZAEHLER_DATEI} enthaelt kein JSON-Objekt der Form "
            f'{{"JJJJ-MM": laufende_nummer}}.'
        )
    return daten


def _speichere_zaehler(daten):
    """Schreibt den Zaehlerstand atomar (erst temporaer, dann ersetzen)."""
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    fd, tmp_pfad = tempfile.mkstemp(dir=str(LOG_DIR), prefix=".zaehler_", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(daten, f, indent=2, ensure_ascii=False, sort_keys=True)
            f.write("\n")
        os.replace(tmp_pfad, ZAEHLER_DATEI)
    except BaseException:
        Path(tmp_pfad).unlink(missing_ok=True)
        raise


def naechste_rechnungsnummer(datum: date) -> str:
    """Vergibt die naechste freie Rechnungsnummer fuer den Monat von `datum`.

    Im ersten Aufruf eines Monats wird START_LAUFENDE_NUMMER (500) vergeben,
    danach wird um eins hochgezaehlt. Ein Monatsschluessel, der bereits in der
    Zaehlerdatei steht, behaelt seinen Stand und zaehlt von dort weiter.

    Der Zaehlerstand wird sofort persistiert, der Aufruf ist damit auch bei
    mehrfacher Verwendung innerhalb eines Laufs eindeutig.

    Args:
        datum: Rechnungsdatum; Jahr und Monat bestimmen den Zaehlerschluessel.

    Returns:
        Rechnungsnummer, z. B. "RE202609500".

    Raises:
        RechnungsnummerError: Wenn die Zaehlerdatei unlesbar ist oder der
            Monatszaehler den dreistelligen Nummernkreis ueberschreitet.
    """
    schluessel = f"{datum:%Y-%m}"
    daten = _lade_zaehler()

    if schluessel in daten:
        bisher = daten[schluessel]
        if not isinstance(bisher, int) or isinstance(bisher, bool) or bisher < 0:
            raise RechnungsnummerError(
                f"Ungueltiger Zaehlerstand fuer {schluessel} in {ZAEHLER_DATEI}: {bisher!r}"
            )
        laufende_nummer = bisher + 1
    else:
        # Neuer Monat: Zaehlung startet oberhalb des manuellen Nummernkreises.
        laufende_nummer = START_LAUFENDE_NUMMER

    if laufende_nummer > MAX_LAUFENDE_NUMMER:
        raise RechnungsnummerError(
            f"Nummernkreis fuer {schluessel} erschoepft: die letzte vergebene "
            f"Nummer ist {MAX_LAUFENDE_NUMMER}. Im Format {PRAEFIX}JJJJMMNNN "
            f"stehen ab {START_LAUFENDE_NUMMER} nur "
            f"{MAX_LAUFENDE_NUMMER - START_LAUFENDE_NUMMER + 1} Rechnungen pro "
            f"Monat zur Verfuegung."
        )

    daten[schluessel] = laufende_nummer
    _speichere_zaehler(daten)

    return f"{PRAEFIX}{datum:%Y%m}{laufende_nummer:03d}"
