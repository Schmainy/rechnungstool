# Rechnungsautomatisierung

Automatische monatliche Rechnungserstellung und E-Mail-Versand über die Microsoft Graph API: Eine Excel-Vorlage wird je Kunde befüllt, per LibreOffice in ein PDF umgewandelt und als Anhang an den Kunden verschickt – gesteuert vom Windows-Task-Scheduler.

## Features

- **PDF-Generierung aus Excel-Vorlage** – Kundendaten, Rechnungsnummer, Leistungszeitraum und Beträge werden in eine Kopie der Vorlage geschrieben; Formate, Schriften und Logo bleiben erhalten. Beträge werden in Python berechnet, nicht per Excel-Formel.
- **Versand über Microsoft Graph** (Client-Credentials-Flow) – wahlweise als Entwurf im Postfach oder direkt; optionale Testumleitung aller Mails an eine Adresse.
- **Monatliche Automatisierung** – gerechnet wird am ersten Werktag des Monats für den zuletzt abgeschlossenen Kalendermonat.
- **Locking-Mechanismus** – eine Monatssperre verhindert doppelte Läufe im selben Monat; `--force` erlaubt bewusste Nachläufe.
- **Fortlaufende Rechnungsnummern** – Format `RE{JJJJMM}{NNN}`, Start je Monat bei 500 (001–499 bleiben für manuelle Rechnungen frei), atomar gespeichert.
- **Fehler-Benachrichtigung per E-Mail** – schlägt die PDF-Erstellung oder der Versand für einen Kunden fehl, geht am Ende ein Sammelbericht an das eigene Postfach.
- **Längenprüfung** – zu lange Leistungsbeschreibungen werden erkannt, bevor sie im PDF abgeschnitten würden.

## Tech-Stack

Python 3 · [msal](https://github.com/AzureAD/microsoft-authentication-library-for-python) · requests · openpyxl · pillow · python-dateutil · LibreOffice (headless, für die PDF-Konvertierung)

## Voraussetzungen

- **Azure App Registration** mit der Anwendungsberechtigung `Mail.Send` (für direkten Versand) und – für den Entwurfsmodus – `Mail.ReadWrite`, jeweils mit Administratorzustimmung. Die Absenderadresse muss ein Postfach dieses Tenants sein.
- **LibreOffice** für die Umwandlung Excel → PDF:
  ```powershell
  winget install --id TheDocumentFoundation.LibreOffice
  ```
- **Windows Task Scheduler** für den automatischen Betrieb.
- Optional die Schrift der Vorlage (Poppins) für die Längenprüfung der Leistungsbeschreibung.

## Installation

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Konfiguration

Im Ordner `config/` liegen vier Vorlagen. Jede davon ohne `.example` kopieren und mit den eigenen Daten füllen:

```powershell
copy config\api_key.example.json       config\api_key.json
copy config\einstellungen.example.json config\einstellungen.json
copy config\firmendaten.example.json   config\firmendaten.json
copy config\kunden.example.json        config\kunden.json
```

| Datei | Inhalt |
| --- | --- |
| `api_key.json` | Azure-Zugangsdaten: `tenant_id`, `client_id`, `client_secret`, `sender_email` |
| `einstellungen.json` | `versand_modus` (`"entwurf"` oder `"direkt"`) und `test_empfaenger_override` (`null` oder eine Testadresse) |
| `firmendaten.json` | Eigene Firmen-, Bank- und Steuerdaten |
| `kunden.json` | Kundenstamm mit Anschrift, Empfängeradresse, Leistungsbeschreibung, Menge, Preis, Steuersatz, Zahlungsziel und `aktiv` |

Außerdem die Rechnungsvorlage anlegen: `templates/Musterrechnung_EXAMPLE.xlsx` nach `templates/Musterrechnung.xlsx` kopieren und Kopf- und Fußzeile (Firma, Anschrift, Bankverbindung, Steuernummer, ggf. Logo) anpassen. Welche Zellen das Tool befüllt, steht in `src/pdf_generator.py`.

Betreff und Text der Rechnungsmail stehen in `templates/email_vorlage.txt` (Platzhalter `{rechnungsnummer}`, `{leistungszeitraum_von}`, `{leistungszeitraum_bis}`).

Alle echten Konfigurationsdateien, die Vorlage, Logs und erzeugte PDFs sind per `.gitignore` von der Versionierung ausgeschlossen.

**Empfohlene Reihenfolge vor dem Scharfschalten:** erst `versand_modus: "entwurf"`, dann `"direkt"` mit gesetztem `test_empfaenger_override`, zuletzt `"direkt"` mit `null`.

## Nutzung

Geplanter Lauf (prüft Lauftag und Monatssperre):

```powershell
python automatischer_lauf.py
```

Nachlauf erzwingen (ignoriert Lauftag und Sperre):

```powershell
python automatischer_lauf.py --force
```

Einzellauf ohne Prüfungen, z. B. zum Testen:

```powershell
python src\main.py
```

Erzeugte Rechnungen landen unter `output/rechnungen/{Jahr}/{Monat}/`, Protokolle unter `logs/`.

### Windows Task Scheduler

Beispiel: täglich um 06:30 Uhr – das Skript entscheidet selbst, ob heute gerechnet wird.

```powershell
$projekt = "<Projektpfad>"
$aktion = New-ScheduledTaskAction -Execute "$projekt\venv\Scripts\python.exe" `
    -Argument "`"$projekt\automatischer_lauf.py`"" -WorkingDirectory $projekt
$trigger = New-ScheduledTaskTrigger -Daily -At 06:30
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName "Rechnungsautomatisierung" `
    -Action $aktion -Trigger $trigger -Principal $principal
```

### Hilfsskripte

| Skript | Zweck |
| --- | --- |
| `scripts/probe_pdfs_erstellen.py` | Erzeugt Probe-PDFs mit erfundenen Testkunden – ohne Versand, ohne Zähler |
| `scripts/probe_mailversand.py --empfaenger "a@example.com"` | Live-Test: erzeugt eine Testrechnung und **versendet sie wirklich** an die angegebenen Adressen |

## Tests

Die Tests brauchen nur die Standardbibliothek; sie erzeugen weder PDFs noch Mails.

```powershell
python -m unittest discover -s tests -t .
```

## Projektstruktur

```
├── automatischer_lauf.py   Einstieg für den Task-Scheduler (Lauftag + Monatssperre + Lauf)
├── config/                 *.example.json-Vorlagen (echte Dateien nur lokal)
├── templates/              Rechnungsvorlage (Beispiel), Mailvorlage
├── scripts/                Probe-Skripte für PDF-Erzeugung und Mailversand
├── tests/                  Unit-Tests
└── src/
    ├── config_loader.py    JSON-Konfiguration laden
    ├── rechnungsnummer.py  fortlaufende Rechnungsnummern
    ├── zeitplan.py         erster Werktag des Monats
    ├── monatssperre.py     höchstens ein Lauf pro Monat
    ├── pdf_generator.py    Vorlage befüllen und nach PDF wandeln
    ├── mail_sender.py      Versand via Microsoft Graph
    └── main.py             Orchestrierung des Rechnungslaufs
```
