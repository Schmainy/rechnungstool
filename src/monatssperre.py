"""Monatssperre gegen einen doppelten automatischen Rechnungslauf.

Der Windows-Task-Scheduler startet automatischer_lauf.py taeglich. Damit in
einem Monat trotzdem nur ein einziger Lauf stattfindet, merkt sich diese
Sperre den zuletzt automatisch abgearbeiteten Monat in
logs/monatssperre.json:

    {"letzter_automatischer_lauf": "2026-09"}

Ist die Datei unlesbar oder beschaedigt, wird bewusst ein Fehler ausgeloest
statt sie zu ignorieren: ein abgebrochener Lauf ist deutlich weniger
schmerzhaft als Rechnungen, die ein zweites Mal beim Kunden landen.
"""

import json
import os
import re
import tempfile
from pathlib import Path

LOG_DIR = Path(__file__).resolve().parent.parent / "logs"
SPERR_DATEI = LOG_DIR / "monatssperre.json"

SCHLUESSEL = "letzter_automatischer_lauf"
_MONAT_MUSTER = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")


class MonatssperreError(Exception):
    """Die Monatssperre konnte nicht sicher gelesen oder geschrieben werden."""


def _pruefe_monat(monat):
    """Stellt sicher, dass `monat` die Form 'JJJJ-MM' hat."""
    if not isinstance(monat, str) or not _MONAT_MUSTER.match(monat):
        raise MonatssperreError(
            f"Ungueltiger Monat {monat!r}: erwartet wird das Format 'JJJJ-MM', "
            f"z. B. '2026-09'."
        )
    return monat


def _lade_sperre():
    """Liest die Sperrdatei. Eine fehlende Datei ergibt einen leeren Stand."""
    if not SPERR_DATEI.is_file():
        return {}
    try:
        with SPERR_DATEI.open(encoding="utf-8") as f:
            daten = json.load(f)
    except json.JSONDecodeError as exc:
        raise MonatssperreError(
            f"Monatssperre {SPERR_DATEI} ist beschaedigt "
            f"(Zeile {exc.lineno}, Spalte {exc.colno}): {exc.msg}\n"
            f"Der Lauf wird abgebrochen, damit die Rechnungen dieses Monats "
            f"nicht versehentlich ein zweites Mal verschickt werden. Bitte die "
            f"Datei von Hand auf {{\"{SCHLUESSEL}\": \"JJJJ-MM\"}} bringen oder "
            f"loeschen, wenn dieser Monat noch nicht gelaufen ist."
        ) from exc
    except OSError as exc:
        raise MonatssperreError(
            f"Monatssperre {SPERR_DATEI} ist nicht lesbar: {exc}\n"
            f"Der Lauf wird abgebrochen, damit die Rechnungen dieses Monats "
            f"nicht versehentlich ein zweites Mal verschickt werden."
        ) from exc

    if not isinstance(daten, dict):
        raise MonatssperreError(
            f"Monatssperre {SPERR_DATEI} enthaelt kein JSON-Objekt der Form "
            f'{{"{SCHLUESSEL}": "JJJJ-MM"}}, sondern {type(daten).__name__}.'
        )

    stand = daten.get(SCHLUESSEL)
    if stand is not None and not (
        isinstance(stand, str) and _MONAT_MUSTER.match(stand)
    ):
        raise MonatssperreError(
            f"Monatssperre {SPERR_DATEI} enthaelt unter \"{SCHLUESSEL}\" den "
            f"unbrauchbaren Wert {stand!r}; erwartet wird 'JJJJ-MM'.\n"
            f"Der Lauf wird abgebrochen, damit die Rechnungen dieses Monats "
            f"nicht versehentlich ein zweites Mal verschickt werden."
        )
    return daten


def ist_bereits_gelaufen(monat: str) -> bool:
    """Sagt, ob fuer `monat` bereits ein automatischer Lauf vermerkt ist.

    Args:
        monat: Monat im Format "JJJJ-MM", z. B. "2026-09".

    Returns:
        True, wenn dieser Monat schon abgearbeitet wurde.

    Raises:
        MonatssperreError: Wenn `monat` kein gueltiger Monatsschluessel ist
            oder die Sperrdatei nicht sicher gelesen werden kann.
    """
    _pruefe_monat(monat)
    return _lade_sperre().get(SCHLUESSEL) == monat


def setze_gelaufen(monat: str) -> None:
    """Vermerkt `monat` als abgearbeitet - atomar wie der Nummernzaehler.

    Erst wird eine temporaere Datei im selben Ordner geschrieben, dann per
    os.replace an die richtige Stelle geschoben. Ein Absturz mitten im
    Schreiben laesst damit nie eine halbe Datei zurueck.

    Args:
        monat: Monat im Format "JJJJ-MM", z. B. "2026-09".

    Raises:
        MonatssperreError: Wenn `monat` ungueltig ist oder die Sperre nicht
            geschrieben werden kann.
    """
    _pruefe_monat(monat)
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        fd, tmp_pfad = tempfile.mkstemp(
            dir=str(LOG_DIR), prefix=".monatssperre_", suffix=".tmp"
        )
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump({SCHLUESSEL: monat}, f, indent=2, ensure_ascii=False)
                f.write("\n")
            os.replace(tmp_pfad, SPERR_DATEI)
        except BaseException:
            Path(tmp_pfad).unlink(missing_ok=True)
            raise
    except OSError as exc:
        raise MonatssperreError(
            f"Monatssperre {SPERR_DATEI} konnte nicht geschrieben werden: {exc}\n"
            f"Ohne gesetzte Sperre koennte der Lauf morgen erneut starten - der "
            f"Lauf wird deshalb abgebrochen."
        ) from exc
