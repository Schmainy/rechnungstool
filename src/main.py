"""Hauptorchestrierung des Rechnungslaufs.

Ablauf eines Laufs:

1. Konfiguration (kunden.json, firmendaten.json, einstellungen.json,
   api_key.json) laden.
2. Voraussetzungen pruefen: Excel-Vorlage, LibreOffice sowie Versandmodus,
   Zugangsdaten und Mailvorlage - alles vor der ersten erzeugten Datei.
3. Leistungszeitraum bestimmen - immer der zuletzt abgeschlossene volle
   Kalendermonat; Rechnungsdatum ist der heutige Tag.
4. Fuer jeden Kunden mit "aktiv": true eine Rechnungsnummer vergeben, die PDF
   erzeugen und die Rechnung per Mail verschicken bzw. als Entwurf ablegen.
5. Erfolg und Fehler je Kunde mit Zeitstempel nach logs/lauf_{datum}.log
   protokollieren und am Ende zusammenfassen.
6. Sind Fehler aufgetreten, geht ein Sammelbericht an das eigene Postfach.

Ein Fehler bei einem Kunden beendet den Lauf nicht - die uebrigen Kunden
werden weiter abgearbeitet. PDF-Erstellung und Versand werden dabei getrennt
betrachtet: eine erzeugte Rechnung bleibt erhalten und archiviert, auch wenn
anschliessend der Versand scheitert.
"""

import logging
import sys
from datetime import date
from pathlib import Path

from dateutil.relativedelta import relativedelta

if __package__ in (None, ""):
    # Direkter Aufruf: python src/main.py
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from src.config_loader import ConfigError, load_config
    from src.mail_sender import (
        beschreibe_versand,
        pruefe_versandkonfiguration,
        sende_fehlerbericht,
        sende_rechnung,
    )
    from src.pdf_generator import erstelle_rechnung, pruefe_voraussetzungen
    from src.rechnungsnummer import naechste_rechnungsnummer
else:
    from .config_loader import ConfigError, load_config
    from .mail_sender import (
        beschreibe_versand,
        pruefe_versandkonfiguration,
        sende_fehlerbericht,
        sende_rechnung,
    )
    from .pdf_generator import erstelle_rechnung, pruefe_voraussetzungen
    from .rechnungsnummer import naechste_rechnungsnummer

PROJEKT_DIR = Path(__file__).resolve().parent.parent
LOG_DIR = PROJEKT_DIR / "logs"

# Phasenkennzeichen fuer die Fehlerliste und den Sammelbericht.
PHASE_PDF = "PDF"
PHASE_MAIL = "Mail"

log = logging.getLogger("rechnungslauf")


def richte_logging_ein(heute: date) -> Path:
    """Schreibt das Laufprotokoll nach logs/lauf_{JJJJ-MM-TT}.log und auf die Konsole."""
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_datei = LOG_DIR / f"lauf_{heute:%Y-%m-%d}.log"

    log.setLevel(logging.INFO)
    log.handlers.clear()

    datei_handler = logging.FileHandler(log_datei, encoding="utf-8")
    datei_handler.setFormatter(
        logging.Formatter("%(asctime)s  %(levelname)-7s  %(message)s", "%Y-%m-%d %H:%M:%S")
    )
    log.addHandler(datei_handler)

    konsole_handler = logging.StreamHandler(sys.stdout)
    konsole_handler.setFormatter(logging.Formatter("%(levelname)-7s  %(message)s"))
    log.addHandler(konsole_handler)

    return log_datei


def bestimme_leistungszeitraum(heute: date):
    """Liefert (erster_tag, letzter_tag) des zuletzt abgeschlossenen Monats."""
    von = (heute - relativedelta(months=1)).replace(day=1)
    bis = von + relativedelta(day=31)  # klemmt auf den letzten Tag des Monats
    return von, bis


def _versandmodus_text(einstellungen: dict) -> str:
    """Beschreibt den eingestellten Versandmodus fuer den Protokollkopf."""
    modus = einstellungen.get("versand_modus")
    erklaerung = {
        "entwurf": "Mails werden nur als Entwurf im Postfach abgelegt",
        "direkt": "Mails werden tatsaechlich versendet",
    }.get(modus, "unbekannt")
    return f"{modus} ({erklaerung})"


def _melde_fehler_per_mail(fehler, api_konfig, heute, log_datei) -> None:
    """Verschickt den Sammelbericht ueber alle Fehler des Laufs.

    Scheitert auch das, wird der Fehler nur protokolliert: der Lauf selbst ist
    an dieser Stelle bereits abgeschlossen und soll nicht daran haengen, dass
    die Benachrichtigung nicht zugestellt werden konnte.
    """
    try:
        empfaenger = sende_fehlerbericht(
            fehler=fehler,
            api_konfig=api_konfig,
            rechnungsdatum=heute,
            log_datei=log_datei,
        )
    except Exception as exc:
        log.error(
            "Fehlerbericht konnte nicht verschickt werden (der Lauf selbst ist "
            "davon unberuehrt): %s",
            exc,
        )
    else:
        log.info("Fehlerbericht an %s verschickt.", empfaenger)


def fuehre_lauf_aus(heute: date = None) -> int:
    """Fuehrt einen kompletten Rechnungslauf aus.

    Returns:
        0, wenn alle aktiven Kunden erfolgreich abgerechnet wurden, sonst 1.
    """
    heute = heute or date.today()
    log_datei = richte_logging_ein(heute)

    try:
        kunden_config = load_config("kunden")
        firmendaten = load_config("firmendaten")
        einstellungen = load_config("einstellungen")
        api_konfig = load_config("api_key")
    except ConfigError as exc:
        log.error("Konfiguration konnte nicht geladen werden:\n%s", exc)
        return 1

    try:
        pruefe_voraussetzungen()
        # Versandkonfiguration frueh pruefen, damit ein Tippfehler in
        # einstellungen.json nicht erst nach der Rechnungserzeugung auffaellt.
        pruefe_versandkonfiguration(einstellungen, api_konfig)
    except Exception as exc:
        log.error("Voraussetzungen nicht erfuellt:\n%s", exc)
        return 1

    kunden = kunden_config.get("kunden", [])
    aktive_kunden = [k for k in kunden if k.get("aktiv")]

    zeitraum_von, zeitraum_bis = bestimme_leistungszeitraum(heute)

    log.info("=" * 70)
    log.info("Rechnungslauf gestartet")
    log.info("Rechnungsdatum:     %s", f"{heute:%d.%m.%Y}")
    log.info(
        "Leistungszeitraum:  %s – %s",
        f"{zeitraum_von:%d.%m.%Y}",
        f"{zeitraum_bis:%d.%m.%Y}",
    )
    log.info("Aktive Kunden:      %d von %d", len(aktive_kunden), len(kunden))
    log.info("Versandmodus:       %s", _versandmodus_text(einstellungen))
    log.info("Protokoll:          %s", log_datei)
    log.info("=" * 70)

    if einstellungen.get("test_empfaenger_override"):
        log.warning(
            "Testumleitung aktiv: ALLE Kundenmails gehen an %s, nicht an die Kunden.",
            einstellungen["test_empfaenger_override"],
        )

    if not aktive_kunden:
        log.warning("Kein Kunde in config/kunden.json ist auf \"aktiv\": true gesetzt.")
        return 0

    erfolge = []  # PDF erzeugt UND verschickt bzw. als Entwurf abgelegt
    nur_pdf = []  # PDF erzeugt, Versand fehlgeschlagen
    fehler = []  # (name, phase, meldung) fuer Protokoll und Sammelbericht

    for kunde in aktive_kunden:
        name = kunde.get("firma") or kunde.get("kunde_id") or "<unbenannter Kunde>"
        try:
            rechnungsnummer = naechste_rechnungsnummer(heute)
            pdf_pfad = erstelle_rechnung(
                kunde=kunde,
                firmendaten=firmendaten,
                rechnungsnummer=rechnungsnummer,
                rechnungsdatum=heute,
                zeitraum_von=zeitraum_von,
                zeitraum_bis=zeitraum_bis,
            )
        except Exception as exc:
            # Ein fehlerhafter Kunde darf den Lauf nicht abbrechen.
            fehler.append((name, PHASE_PDF, str(exc)))
            log.error("FEHLER  %s: PDF-Erstellung fehlgeschlagen: %s", name, exc)
            continue

        log.info("PDF     %s | %s | %s", name, rechnungsnummer, pdf_pfad)

        try:
            sende_rechnung(
                kunde=kunde,
                rechnungsnummer=rechnungsnummer,
                pdf_pfad=pdf_pfad,
                leistungszeitraum_von=zeitraum_von,
                leistungszeitraum_bis=zeitraum_bis,
                einstellungen=einstellungen,
                api_konfig=api_konfig,
            )
        except Exception as exc:
            # Die PDF ist erzeugt und archiviert - nur der Versand fehlt.
            nur_pdf.append((name, rechnungsnummer, pdf_pfad))
            fehler.append((name, PHASE_MAIL, str(exc)))
            log.error(
                "FEHLER  %s: PDF erstellt, aber Versand fehlgeschlagen: %s", name, exc
            )
            continue

        erfolge.append((name, rechnungsnummer, pdf_pfad))
        log.info(
            "OK      %s | %s | %s",
            name,
            rechnungsnummer,
            beschreibe_versand(kunde, einstellungen),
        )

    log.info("-" * 70)
    log.info(
        "Zusammenfassung: %d vollstaendig, %d nur als PDF, %d Fehler",
        len(erfolge),
        len(nur_pdf),
        len(fehler),
    )
    for name, rechnungsnummer, pdf_pfad in erfolge:
        log.info("  erledigt       %s (%s) -> %s", name, rechnungsnummer, pdf_pfad.name)
    for name, rechnungsnummer, pdf_pfad in nur_pdf:
        log.info("  ohne Versand   %s (%s) -> %s", name, rechnungsnummer, pdf_pfad.name)
    for name, phase, meldung in fehler:
        log.info("  fehlgeschlagen %s [%s]: %s", name, phase, meldung)

    if fehler:
        _melde_fehler_per_mail(fehler, api_konfig, heute, log_datei)

    log.info("=" * 70)

    return 1 if fehler else 0


def main() -> int:
    return fuehre_lauf_aus()


if __name__ == "__main__":
    sys.exit(main())
