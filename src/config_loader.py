"""Laden der JSON-Konfigurationsdateien aus dem config/-Verzeichnis.

Die produktiven Konfigurationsdateien (api_key.json, einstellungen.json,
firmendaten.json, kunden.json) sind bewusst nicht im Repository enthalten.
Zu jeder Datei existiert eine Vorlage <name>.example.json.
"""

import json
from pathlib import Path

CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"


class ConfigError(Exception):
    """Fehler beim Laden oder Auswerten einer Konfigurationsdatei."""


def load_config(name):
    """Lädt config/<name>.json und gibt den Inhalt als dict zurück.

    Args:
        name: Name der Konfiguration, mit oder ohne ".json"-Endung,
              z. B. "kunden" oder "kunden.json".

    Raises:
        ConfigError: Wenn die Datei fehlt oder kein gültiges JSON enthält.
    """
    dateiname = name if name.endswith(".json") else f"{name}.json"
    pfad = CONFIG_DIR / dateiname
    vorlage = CONFIG_DIR / f"{Path(dateiname).stem}.example.json"

    if not pfad.is_file():
        hinweis = (
            f"Konfigurationsdatei nicht gefunden: {pfad}\n"
            f"Bitte die Vorlage kopieren und ausfüllen:\n"
            f"    copy \"{vorlage}\" \"{pfad}\""
        )
        if not vorlage.is_file():
            hinweis = (
                f"Konfigurationsdatei nicht gefunden: {pfad}\n"
                f"Auch die Vorlage {vorlage.name} fehlt im config-Verzeichnis."
            )
        raise ConfigError(hinweis)

    try:
        with pfad.open(encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as exc:
        raise ConfigError(
            f"Ungültiges JSON in {pfad} (Zeile {exc.lineno}, Spalte {exc.colno}): {exc.msg}\n"
            f"Vergleiche den Aufbau mit der Vorlage {vorlage.name}."
        ) from exc
