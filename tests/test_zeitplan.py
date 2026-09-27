"""Tests fuer die Bestimmung des Lauftags.

Abgedeckt sind alle sieben Wochentage, auf die ein Monatserster fallen kann.
Die Monate stammen aus echten Kalenderjahren, damit die Wochentage stimmen.
"""

import unittest
from datetime import date, timedelta

from . import _pfad  # noqa: F401

from src.zeitplan import erster_werktag_des_monats, ist_planmaessiger_lauftag

# (Monatserster, Wochentagsname, weekday()-Index, erwarteter erster Werktag)
FAELLE = [
    (date(2026, 6, 1), "Montag", 0, date(2026, 6, 1)),
    (date(2026, 9, 1), "Dienstag", 1, date(2026, 9, 1)),
    (date(2026, 7, 1), "Mittwoch", 2, date(2026, 7, 1)),
    (date(2026, 10, 1), "Donnerstag", 3, date(2026, 10, 1)),
    (date(2026, 5, 1), "Freitag", 4, date(2026, 5, 1)),
    (date(2026, 8, 1), "Samstag", 5, date(2026, 8, 3)),
    (date(2026, 11, 1), "Sonntag", 6, date(2026, 11, 2)),
]


class ErsterWerktagTest(unittest.TestCase):
    def test_alle_wochentage_des_ersten(self):
        for erster, wochentag, index, erwartet in FAELLE:
            with self.subTest(wochentag=wochentag, monat=f"{erster:%Y-%m}"):
                # Absicherung, dass der Testfall wirklich diesen Wochentag trifft.
                self.assertEqual(erster.weekday(), index)
                self.assertEqual(erster_werktag_des_monats(erster), erwartet)
                self.assertLess(erwartet.weekday(), 5)

    def test_samstag_verschiebt_auf_montag_den_dritten(self):
        # 01.08.2026 ist ein Samstag.
        self.assertEqual(date(2026, 8, 1).weekday(), 5)
        ergebnis = erster_werktag_des_monats(date(2026, 8, 1))
        self.assertEqual(ergebnis, date(2026, 8, 3))
        self.assertEqual(ergebnis.weekday(), 0)

    def test_sonntag_verschiebt_auf_montag_den_zweiten(self):
        # 01.11.2026 ist ein Sonntag.
        self.assertEqual(date(2026, 11, 1).weekday(), 6)
        ergebnis = erster_werktag_des_monats(date(2026, 11, 1))
        self.assertEqual(ergebnis, date(2026, 11, 2))
        self.assertEqual(ergebnis.weekday(), 0)

    def test_beliebiger_tag_im_monat_liefert_denselben_werktag(self):
        for tag in (1, 3, 17, 31):
            with self.subTest(tag=tag):
                self.assertEqual(
                    erster_werktag_des_monats(date(2026, 8, tag)), date(2026, 8, 3)
                )

    def test_jahreswechsel(self):
        # 01.01.2028 ist ein Samstag -> Montag, der 3. Januar.
        self.assertEqual(erster_werktag_des_monats(date(2028, 1, 20)), date(2028, 1, 3))


class PlanmaessigerLauftagTest(unittest.TestCase):
    def test_genau_der_erste_werktag_ist_lauftag(self):
        for _, wochentag, _index, lauftag in FAELLE:
            with self.subTest(wochentag=wochentag):
                self.assertTrue(ist_planmaessiger_lauftag(lauftag))

    def test_tag_davor_und_danach_ist_kein_lauftag(self):
        for _, wochentag, _index, lauftag in FAELLE:
            with self.subTest(wochentag=wochentag):
                self.assertFalse(ist_planmaessiger_lauftag(lauftag + timedelta(days=1)))
                if lauftag.day > 1:
                    # Wochenendtage vor dem Lauftag zaehlen ebenfalls nicht.
                    self.assertFalse(
                        ist_planmaessiger_lauftag(lauftag - timedelta(days=1))
                    )

    def test_genau_ein_lauftag_pro_monat(self):
        tag = date(2026, 1, 1)
        gezaehlt = {}
        while tag < date(2028, 1, 1):
            if ist_planmaessiger_lauftag(tag):
                schluessel = f"{tag:%Y-%m}"
                gezaehlt[schluessel] = gezaehlt.get(schluessel, 0) + 1
            tag += timedelta(days=1)
        self.assertEqual(len(gezaehlt), 24)
        self.assertEqual(set(gezaehlt.values()), {1})


if __name__ == "__main__":
    unittest.main()
