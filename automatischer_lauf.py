"""Wrapper fuer den unbeaufsichtigten Rechnungslauf (Windows Task Scheduler).

Der geplante Task startet dieses Skript taeglich um 06:30 Uhr. Zwei Wachen
sorgen dafuer, dass daraus trotzdem genau ein Lauf pro Monat wird:

1. zeitplan.ist_planmaessiger_lauftag() - gerechnet wird nur am ersten
   Werktag des Monats. An jedem anderen Tag endet das Skript sofort mit
   Exit-Code 0.
2. monatssperre - ist der Monat bereits vermerkt, endet das Skript ebenfalls
   mit Exit-Code 0.

Die Sperre wird gesetzt, *bevor* der eigentliche Lauf beginnt. Stuerzt der
Lauf mitten in der Abarbeitung ab, startet er am naechsten Tag also nicht von
selbst noch einmal - der Fehler geht ohnehin per Sammelmail an das eigene
Postfach, und nachgeholt wird von Hand.

Aufrufvarianten:

    python automatischer_lauf.py            geplanter Lauf mit beiden Wachen
    python automatischer_lauf.py --force    bewusster Nachlauf ohne Wachen
    python src/main.py                      direkter Einzellauf ohne Wachen

Exit-Codes: 0 = Lauf erfolgreich oder planmaessig uebersprungen, 1 = Fehler.
"""

import argparse
import sys
from datetime import date
from pathlib import Path

PROJEKT_DIR = Path(__file__).resolve().parent
if str(PROJEKT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJEKT_DIR))

from src import main as main_modul  # noqa: E402
from src import monatssperre, zeitplan  # noqa: E402
from src.config_loader import ConfigError, load_config  # noqa: E402
from src.monatssperre import MonatssperreError  # noqa: E402

log = main_modul.log

# Phasenkennzeichen fuer den Sammelbericht, passend zu main.PHASE_PDF/PHASE_MAIL.
PHASE_SPERRE = "Monatssperre"


def _melde_abbruch(meldung: str, heute: date, log_datei: Path) -> None:
    """Schickt den Sammelbericht, wenn der Lauf vor dem ersten Kunden scheitert.

    Der Bericht laeuft ueber dieselbe Benachrichtigung wie Fehler waehrend des
    Laufs. Scheitert auch das (z. B. weil die Konfiguration nicht lesbar ist),
    bleibt es beim Protokolleintrag - mehr kann an dieser Stelle nicht getan
    werden.
    """
    try:
        api_konfig = load_config("api_key")
    except ConfigError as exc:
        log.error("Fehlerbericht nicht moeglich, api_key.json nicht lesbar: %s", exc)
        return
    main_modul._melde_fehler_per_mail(
        fehler=[("automatischer Lauf", PHASE_SPERRE, meldung)],
        api_konfig=api_konfig,
        heute=heute,
        log_datei=log_datei,
    )


def fuehre_automatischen_lauf_aus(heute: date = None, force: bool = False) -> int:
    """Fuehrt den geplanten Lauf aus, sofern Tag und Monatssperre das zulassen.

    Args:
        heute: Laufdatum; standardmaessig der heutige Tag.
        force: Ueberspringt Tagespruefung und Monatssperre, setzt die Sperre
            aber nach dem Lauf trotzdem neu.

    Returns:
        0 bei Erfolg oder planmaessigem Ueberspringen, sonst 1.
    """
    heute = heute or date.today()
    log_datei = main_modul.richte_logging_ein(heute)
    monat = f"{heute:%Y-%m}"

    if force:
        log.info(
            "Erzwungener Lauf (--force): Lauftag und Monatssperre werden "
            "uebersprungen, die Sperre fuer %s wird danach neu gesetzt.",
            monat,
        )
        try:
            return main_modul.fuehre_lauf_aus(heute)
        finally:
            # Auch nach einem Abbruch soll der Monat als erledigt gelten -
            # sonst wuerde ein spaeterer geplanter Lauf ungefragt nachlegen.
            try:
                monatssperre.setze_gelaufen(monat)
                log.info("Monatssperre fuer %s gesetzt.", monat)
            except MonatssperreError as exc:
                log.error("Monatssperre konnte nicht gesetzt werden: %s", exc)

    if not zeitplan.ist_planmaessiger_lauftag(heute):
        log.info(
            "Kein planmaessiger Lauftag, uebersprungen "
            "(heute %s, faellig am %s).",
            f"{heute:%d.%m.%Y}",
            f"{zeitplan.erster_werktag_des_monats(heute):%d.%m.%Y}",
        )
        return 0

    try:
        bereits_gelaufen = monatssperre.ist_bereits_gelaufen(monat)
    except MonatssperreError as exc:
        log.error(
            "Monatssperre nicht auswertbar, der Lauf wird zur Sicherheit "
            "abgebrochen:\n%s",
            exc,
        )
        _melde_abbruch(str(exc), heute, log_datei)
        return 1

    if bereits_gelaufen:
        log.info("Fuer diesen Monat bereits gelaufen, uebersprungen (%s).", monat)
        return 0

    # Sperre vor dem Lauf setzen: ein Absturz mittendrin soll morgen keinen
    # zweiten automatischen Versuch ausloesen.
    try:
        monatssperre.setze_gelaufen(monat)
    except MonatssperreError as exc:
        log.error(
            "Monatssperre konnte nicht gesetzt werden, der Lauf wird zur "
            "Sicherheit abgebrochen:\n%s",
            exc,
        )
        _melde_abbruch(str(exc), heute, log_datei)
        return 1

    log.info("Monatssperre fuer %s gesetzt, Lauf startet.", monat)
    return main_modul.fuehre_lauf_aus(heute)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="automatischer_lauf.py",
        description=(
            "Geplanter Rechnungslauf: laeuft nur am ersten Werktag des Monats "
            "und nur einmal pro Monat."
        ),
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help=(
            "Lauftag und Monatssperre ignorieren (bewusster Nachlauf); "
            "die Monatssperre wird danach trotzdem neu gesetzt."
        ),
    )
    args = parser.parse_args(argv)
    return fuehre_automatischen_lauf_aus(force=args.force)


if __name__ == "__main__":
    sys.exit(main())
