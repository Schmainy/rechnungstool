"""Tests fuer die Vergabe der Rechnungsnummern."""

import json
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

from . import _pfad  # noqa: F401

from src import rechnungsnummer
from src.rechnungsnummer import RechnungsnummerError, naechste_rechnungsnummer


class RechnungsnummerTest(unittest.TestCase):
    """Jeder Test bekommt einen eigenen leeren logs-Ordner."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        log_dir = Path(self._tmp.name)
        patches = [
            mock.patch.object(rechnungsnummer, "LOG_DIR", log_dir),
            mock.patch.object(
                rechnungsnummer, "ZAEHLER_DATEI", log_dir / "rechnungsnummern_zaehler.json"
            ),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        self.addCleanup(self._tmp.cleanup)
        self.zaehler_datei = rechnungsnummer.ZAEHLER_DATEI

    def _schreibe_zaehler(self, daten):
        self.zaehler_datei.write_text(json.dumps(daten), encoding="utf-8")

    def test_erste_nummer_eines_monats_startet_bei_500(self):
        self.assertEqual(
            naechste_rechnungsnummer(date(2026, 9, 1)), "RE202609500"
        )

    def test_drei_aufrufe_hintereinander(self):
        nummern = [naechste_rechnungsnummer(date(2026, 9, 1)) for _ in range(3)]
        self.assertEqual(
            nummern, ["RE202609500", "RE202609501", "RE202609502"]
        )

    def test_neuer_monat_startet_wieder_bei_500(self):
        naechste_rechnungsnummer(date(2026, 9, 1))
        naechste_rechnungsnummer(date(2026, 9, 1))
        self.assertEqual(
            naechste_rechnungsnummer(date(2026, 10, 1)), "RE202610500"
        )

    def test_bestehender_monatsschluessel_zaehlt_von_dort_weiter(self):
        # Alte Staende (z. B. aus der Zeit vor der Umstellung) bleiben gueltig.
        self._schreibe_zaehler({"2026-08": 3})
        self.assertEqual(
            naechste_rechnungsnummer(date(2026, 8, 15)), "RE202608004"
        )

    def test_zaehlerstand_wird_sofort_geschrieben(self):
        naechste_rechnungsnummer(date(2026, 9, 1))
        stand = json.loads(self.zaehler_datei.read_text(encoding="utf-8"))
        self.assertEqual(stand, {"2026-09": 500})

    def test_beschaedigte_zaehlerdatei_meldet_fehler(self):
        self.zaehler_datei.write_text("{kaputt", encoding="utf-8")
        with self.assertRaises(RechnungsnummerError):
            naechste_rechnungsnummer(date(2026, 9, 1))

    def test_ungueltiger_zaehlerstand_meldet_fehler(self):
        self._schreibe_zaehler({"2026-09": "viele"})
        with self.assertRaises(RechnungsnummerError):
            naechste_rechnungsnummer(date(2026, 9, 1))

    def test_erschoepfter_nummernkreis_meldet_fehler(self):
        self._schreibe_zaehler({"2026-09": rechnungsnummer.MAX_LAUFENDE_NUMMER})
        with self.assertRaises(RechnungsnummerError):
            naechste_rechnungsnummer(date(2026, 9, 1))


if __name__ == "__main__":
    unittest.main()
