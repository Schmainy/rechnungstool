"""Bestimmt, an welchem Tag der automatische Rechnungslauf faellig ist.

Gerechnet wird immer am ersten Werktag eines Monats. Werktag heisst hier
Montag bis Freitag - Feiertage werden bewusst nicht geprueft: sie sind je
Bundesland verschieden, und ein Lauf an einem Feiertag schadet nicht, weil
die Rechnungen ohnehin unbeaufsichtigt erzeugt und verschickt werden.

Faellt der Monatserste auf einen Samstag oder Sonntag, verschiebt sich der
Lauf auf den darauffolgenden Montag (Sa -> der 3., So -> der 2.).
"""

from datetime import date, timedelta

# date.weekday(): Montag = 0 ... Samstag = 5, Sonntag = 6
SAMSTAG = 5


def erster_werktag_des_monats(datum: date) -> date:
    """Liefert den ersten Werktag (Mo-Fr) des Kalendermonats von `datum`.

    Args:
        datum: Ein beliebiger Tag; nur Jahr und Monat werden ausgewertet.

    Returns:
        Der erste Montag bis Freitag dieses Monats. Faellt der Monatserste auf
        ein Wochenende, ist das der darauffolgende Montag.
    """
    erster = datum.replace(day=1)
    if erster.weekday() >= SAMSTAG:
        # Samstag -> zwei Tage weiter, Sonntag -> einen Tag weiter.
        erster += timedelta(days=7 - erster.weekday())
    return erster


def ist_planmaessiger_lauftag(datum: date) -> bool:
    """Sagt, ob `datum` genau der erste Werktag seines Monats ist."""
    return datum == erster_werktag_des_monats(datum)
