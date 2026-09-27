"""Tests fuer den Wrapper des automatischen Laufs.

main.fuehre_lauf_aus() wird durchgehend durch einen Stub ersetzt: es werden
weder PDFs erzeugt noch Mails verschickt, protokolliert wird nur, OB der Lauf
gestartet worden waere. Die Monatssperre liegt in einem temporaeren Ordner,
damit logs/monatssperre.json des echten Projekts unberuehrt bleibt.
"""

import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

from . import _pfad  # noqa: F401

import automatischer_lauf
from src import monatssperre

# 01.09.2026 ist ein Dienstag und damit der planmaessige Lauftag.
LAUFTAG = date(2026, 9, 1)
KEIN_LAUFTAG = date(2026, 9, 2)
MONAT = "2026-09"


class AutomatischerLaufTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        log_dir = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

        # Sperre und Protokoll in den temporaeren Ordner umlenken.
        patches = [
            mock.patch.object(monatssperre, "LOG_DIR", log_dir),
            mock.patch.object(monatssperre, "SPERR_DATEI", log_dir / "monatssperre.json"),
            mock.patch.object(
                automatischer_lauf.main_modul,
                "richte_logging_ein",
                lambda heute: log_dir / f"lauf_{heute:%Y-%m-%d}.log",
            ),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)

        # Der eigentliche Rechnungslauf wird nie ausgefuehrt, nur mitgezaehlt.
        self.lauf = mock.Mock(return_value=0, name="fuehre_lauf_aus")
        p = mock.patch.object(
            automatischer_lauf.main_modul, "fuehre_lauf_aus", self.lauf
        )
        p.start()
        self.addCleanup(p.stop)

        # Der Fehlerbericht wuerde eine Mail verschicken - ebenfalls stubben.
        self.bericht = mock.Mock(name="_melde_abbruch")
        p = mock.patch.object(automatischer_lauf, "_melde_abbruch", self.bericht)
        p.start()
        self.addCleanup(p.stop)

    # -- Fall 1: falscher Tag ------------------------------------------------

    def test_kein_planmaessiger_lauftag_ueberspringt(self):
        ergebnis = automatischer_lauf.fuehre_automatischen_lauf_aus(heute=KEIN_LAUFTAG)
        self.assertEqual(ergebnis, 0)
        self.lauf.assert_not_called()
        self.assertFalse(monatssperre.SPERR_DATEI.exists())

    # -- Fall 2: richtiger Tag, Sperre offen ---------------------------------

    def test_lauftag_ohne_sperre_startet_lauf_und_setzt_sperre(self):
        ergebnis = automatischer_lauf.fuehre_automatischen_lauf_aus(heute=LAUFTAG)
        self.assertEqual(ergebnis, 0)
        self.lauf.assert_called_once_with(LAUFTAG)
        self.assertTrue(monatssperre.ist_bereits_gelaufen(MONAT))

    def test_sperre_wird_vor_dem_lauf_gesetzt(self):
        # Der Stub prueft mitten im "Lauf", ob die Sperre schon steht.
        stand = {}

        def pruefe(heute):
            stand["gesperrt"] = monatssperre.ist_bereits_gelaufen(MONAT)
            return 0

        self.lauf.side_effect = pruefe
        automatischer_lauf.fuehre_automatischen_lauf_aus(heute=LAUFTAG)
        self.assertTrue(stand["gesperrt"])

    def test_rueckgabewert_des_laufs_wird_durchgereicht(self):
        self.lauf.return_value = 1
        ergebnis = automatischer_lauf.fuehre_automatischen_lauf_aus(heute=LAUFTAG)
        self.assertEqual(ergebnis, 1)

    # -- Fall 3: Sperre bereits gesetzt --------------------------------------

    def test_bereits_gelaufen_ueberspringt(self):
        monatssperre.setze_gelaufen(MONAT)
        ergebnis = automatischer_lauf.fuehre_automatischen_lauf_aus(heute=LAUFTAG)
        self.assertEqual(ergebnis, 0)
        self.lauf.assert_not_called()

    def test_zweiter_aufruf_am_selben_tag_laeuft_nicht_erneut(self):
        automatischer_lauf.fuehre_automatischen_lauf_aus(heute=LAUFTAG)
        automatischer_lauf.fuehre_automatischen_lauf_aus(heute=LAUFTAG)
        self.assertEqual(self.lauf.call_count, 1)

    # -- Fall 4: --force -----------------------------------------------------

    def test_force_ignoriert_falschen_tag(self):
        ergebnis = automatischer_lauf.fuehre_automatischen_lauf_aus(
            heute=KEIN_LAUFTAG, force=True
        )
        self.assertEqual(ergebnis, 0)
        self.lauf.assert_called_once_with(KEIN_LAUFTAG)
        self.assertTrue(monatssperre.ist_bereits_gelaufen(MONAT))

    def test_force_ignoriert_gesetzte_sperre(self):
        monatssperre.setze_gelaufen(MONAT)
        ergebnis = automatischer_lauf.fuehre_automatischen_lauf_aus(
            heute=LAUFTAG, force=True
        )
        self.assertEqual(ergebnis, 0)
        self.lauf.assert_called_once_with(LAUFTAG)
        self.assertTrue(monatssperre.ist_bereits_gelaufen(MONAT))

    def test_force_setzt_sperre_auch_nach_abbruch(self):
        self.lauf.side_effect = RuntimeError("Lauf abgestuerzt")
        with self.assertRaises(RuntimeError):
            automatischer_lauf.fuehre_automatischen_lauf_aus(heute=LAUFTAG, force=True)
        self.assertTrue(monatssperre.ist_bereits_gelaufen(MONAT))

    # -- Defekte Sperre ------------------------------------------------------

    def test_kaputte_sperre_bricht_ab_und_meldet(self):
        monatssperre.SPERR_DATEI.parent.mkdir(parents=True, exist_ok=True)
        monatssperre.SPERR_DATEI.write_text("{kaputt", encoding="utf-8")
        ergebnis = automatischer_lauf.fuehre_automatischen_lauf_aus(heute=LAUFTAG)
        self.assertEqual(ergebnis, 1)
        self.lauf.assert_not_called()
        self.bericht.assert_called_once()

    # -- Kommandozeile -------------------------------------------------------

    def test_argumente(self):
        with mock.patch.object(
            automatischer_lauf, "fuehre_automatischen_lauf_aus", return_value=0
        ) as lauf:
            automatischer_lauf.main([])
            lauf.assert_called_once_with(force=False)
        with mock.patch.object(
            automatischer_lauf, "fuehre_automatischen_lauf_aus", return_value=0
        ) as lauf:
            automatischer_lauf.main(["--force"])
            lauf.assert_called_once_with(force=True)


if __name__ == "__main__":
    unittest.main()
