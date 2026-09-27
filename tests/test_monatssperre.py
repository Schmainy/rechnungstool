"""Tests fuer die Monatssperre."""

import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from . import _pfad  # noqa: F401

from src import monatssperre
from src.monatssperre import MonatssperreError, ist_bereits_gelaufen, setze_gelaufen


class MonatssperreTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        log_dir = Path(self._tmp.name)
        patches = [
            mock.patch.object(monatssperre, "LOG_DIR", log_dir),
            mock.patch.object(monatssperre, "SPERR_DATEI", log_dir / "monatssperre.json"),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        self.addCleanup(self._tmp.cleanup)
        self.sperr_datei = monatssperre.SPERR_DATEI

    def test_ohne_datei_ist_nichts_gelaufen(self):
        self.assertFalse(self.sperr_datei.exists())
        self.assertFalse(ist_bereits_gelaufen("2026-09"))

    def test_setzen_und_lesen(self):
        setze_gelaufen("2026-09")
        self.assertTrue(ist_bereits_gelaufen("2026-09"))

    def test_anderer_monat_ist_nicht_gesperrt(self):
        setze_gelaufen("2026-09")
        self.assertFalse(ist_bereits_gelaufen("2026-10"))
        self.assertFalse(ist_bereits_gelaufen("2026-08"))

    def test_dateiformat(self):
        setze_gelaufen("2026-09")
        inhalt = json.loads(self.sperr_datei.read_text(encoding="utf-8"))
        self.assertEqual(inhalt, {"letzter_automatischer_lauf": "2026-09"})

    def test_erneutes_setzen_ueberschreibt(self):
        setze_gelaufen("2026-09")
        setze_gelaufen("2026-10")
        self.assertFalse(ist_bereits_gelaufen("2026-09"))
        self.assertTrue(ist_bereits_gelaufen("2026-10"))

    def test_keine_temporaeren_reste(self):
        setze_gelaufen("2026-09")
        uebrig = sorted(p.name for p in Path(self._tmp.name).iterdir())
        self.assertEqual(uebrig, ["monatssperre.json"])

    def test_kaputtes_json_meldet_fehler(self):
        self.sperr_datei.write_text("{nicht wirklich json", encoding="utf-8")
        with self.assertRaises(MonatssperreError) as ctx:
            ist_bereits_gelaufen("2026-09")
        self.assertIn("beschaedigt", str(ctx.exception))

    def test_falscher_dateityp_meldet_fehler(self):
        self.sperr_datei.write_text('["2026-09"]', encoding="utf-8")
        with self.assertRaises(MonatssperreError):
            ist_bereits_gelaufen("2026-09")

    def test_unbrauchbarer_wert_meldet_fehler(self):
        self.sperr_datei.write_text(
            '{"letzter_automatischer_lauf": "September"}', encoding="utf-8"
        )
        with self.assertRaises(MonatssperreError):
            ist_bereits_gelaufen("2026-09")

    def test_ungueltiger_monat_meldet_fehler(self):
        for monat in ("2026-13", "26-09", "2026/09", "", None, 202609):
            with self.subTest(monat=monat):
                with self.assertRaises(MonatssperreError):
                    ist_bereits_gelaufen(monat)
                with self.assertRaises(MonatssperreError):
                    setze_gelaufen(monat)


if __name__ == "__main__":
    unittest.main()
