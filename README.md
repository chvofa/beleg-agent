# Beleg-Agent

Automatische Verarbeitung von Rechnungen und Belegen mit Claude Vision API.

Der Agent ueberwacht einen Inbox-Ordner, extrahiert Rechnungsdaten (Rechnungssteller, Betrag, Datum, Waehrung, Zahlungsart) via KI und legt Belege strukturiert in Monatsordnern ab. Zusaetzlich werden Kreditkarten- und Bank-Transaktionen automatisch mit den erfassten Belegen abgeglichen.

## Features

- **Automatische Belegerfassung** – PDFs und Bilder in die Inbox legen, der Agent erledigt den Rest
- **Claude Vision API** – Extrahiert strukturierte Daten aus Rechnungen/Belegen
- **Confidence-basierte Ablage** – Hohe Sicherheit: automatisch ablegen, niedrige: zur Pruefung markieren
- **Excel-Protokoll** – Alle Belege werden in einer zentralen Excel-Datei erfasst
- **KK-Abgleich** – Kreditkarten-CSV mit Belegen abgleichen (Semikolon-Format)
- **Bank-Abgleich** – Bankkonto-CSV mit Belegen abgleichen (Semikolon-Format)
- **Dauerauftraege** – Wiederkehrende Rechnungen separat erfassen
- **System Tray** – Laeuft als Windows-Hintergrund-Dienst mit Tray-Icon
- **Toast-Benachrichtigungen** – Status-Updates und Erinnerungen
- **Duplikat-Erkennung** – Verhindert doppelte Erfassung

## Voraussetzungen

- Windows 10/11
- Python 3.11+
- [Anthropic API Key](https://console.anthropic.com/)

## Installation

### 1. Repository klonen

```bash
git clone https://github.com/DEIN_USER/beleg-agent.git
cd beleg-agent
```

### 2. Abhaengigkeiten installieren

```bash
pip install -r requirements.txt
```

### 3. API-Key setzen

Den Anthropic API Key als Windows-Umgebungsvariable setzen:

```powershell
[Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", "sk-ant-...", "User")
```

Oder ueber: Systemsteuerung > System > Erweiterte Systemeinstellungen > Umgebungsvariablen

### 4. Lokale Konfiguration erstellen

```bash
copy config_local.example.py config_local.py
```

Dann `config_local.py` bearbeiten und den Pfad zu deinem Belege-Ordner anpassen:

```python
ABLAGE_STAMMPFAD = r"C:\Users\DEIN_USER\Pfad\zu\Belege"
```

Der Ordner sollte folgende Unterordner enthalten (werden automatisch erstellt):
- `_Inbox` – Hier Belege ablegen zur Verarbeitung
- `_Abgleich` – Hier KK/Bank-CSVs ablegen fuer den Abgleich
- `_Dauerauftraege` – Hier PDFs von Dauerauftraegen ablegen

## Starten

### Option A: System Tray (empfohlen)

Doppelklick auf `start_beleg_agent.vbs` – startet den Agent unsichtbar mit Tray-Icon.

### Option B: Terminal

```bash
python beleg_agent.py
```

### Autostart einrichten

1. `Win + R` > `shell:startup`
2. Verkuepfung zu `start_beleg_agent.vbs` in den Autostart-Ordner legen

## Verwendung

### Belege verarbeiten

1. PDF oder Bild in den `_Inbox`-Ordner legen
2. Der Agent erkennt die Datei automatisch
3. Je nach Confidence:
   - **>= 85%** – Automatisch abgelegt in `[Jahr]/[Monat]/[Name] - [Waehrung] [Betrag].pdf`
   - **60-85%** – Toast-Hinweis, Datei wird als `[PRUEFEN]_...` markiert
   - **< 60%** – Datei wird als `[PRUEFEN]_...` markiert

### KK-Abgleich

1. KK-Auszug als CSV (Semikolon-getrennt) in `_Abgleich` legen
2. Im Tray-Menue: Abgleich > KK-Abgleich starten
3. Oder: `python abgleich.py`

### Bank-Abgleich

1. Bankauszug als CSV (Semikolon-getrennt) in `_Abgleich` legen
2. Im Tray-Menue: Abgleich > Bank-Abgleich starten
3. Oder: `python abgleich_bank.py`

### Dauerauftraege

1. Dauerauftrags-PDFs in `_Dauerauftraege` legen
2. Im Tray-Menue: Abgleich > Dauerauftraege erfassen
3. Oder: `python dauerauftraege.py`

## Dateistruktur

```
beleg-agent/
  beleg_agent.py          # Hauptagent (Watchdog + Claude Vision + Ablage)
  tray_agent.py           # System Tray Launcher
  config.py               # Allgemeine Konfiguration (Schwellenwerte, Spalten, etc.)
  config_local.py         # Lokale Konfiguration mit persoenlichen Pfaden (nicht im Git)
  config_local.example.py # Template fuer config_local.py
  abgleich.py             # Kreditkarten-Abgleich
  abgleich_bank.py        # Bank-Abgleich
  dauerauftraege.py       # Dauerauftraege erfassen
  status.py               # Status-Checker
  start_beleg_agent.bat   # Startskript
  start_beleg_agent.vbs   # Unsichtbarer Start (kein Konsolenfenster)
  requirements.txt        # Python-Abhaengigkeiten
```

## Konfiguration

Alle Einstellungen in `config.py`:

| Einstellung | Standard | Beschreibung |
|-------------|----------|--------------|
| `ANTHROPIC_MODEL` | `claude-sonnet-4-20250514` | Claude-Modell fuer Vision API |
| `CONFIDENCE_AUTO` | `0.85` | Ab diesem Score: automatische Ablage |
| `CONFIDENCE_RUECKFRAGE` | `0.60` | Ab diesem Score: zur Pruefung markieren |
| `RUECKFRAGE_TIMEOUT_SEKUNDEN` | `60` | Timeout fuer Terminal-Rueckfragen |
| `ERLAUBTE_ENDUNGEN` | `.pdf .jpg .jpeg .png` | Unterstuetzte Dateiformate |

## Lizenz

MIT
