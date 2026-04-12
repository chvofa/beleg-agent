# Beleg-Agent

Automatische Verarbeitung von Rechnungen und Belegen mit Claude Vision API.

Der Agent ueberwacht einen Inbox-Ordner, extrahiert Rechnungsdaten via KI und legt Belege strukturiert ab. Kreditkarten- und Bank-Transaktionen werden automatisch mit den erfassten Belegen abgeglichen. Unterstuetzt mehrere Banken (UBS, Raiffeisen, PostFinance) und Fremdwaehrungstransaktionen.

Laeuft auf **Windows** und **macOS**. Das Web-Interface ist ueber den Browser bedienbar – Belege koennen auch direkt per Adobe Scan, OneDrive oder Drag & Drop in die Inbox gelegt werden.

## So funktioniert es

```
                         _Inbox/
                           |
                 PDF / Bild ablegen
            (manuell, Scan, OneDrive Sync)
                           |
                    Claude Vision API
                    analysiert den Beleg
                           |
              +------------+------------+
              |            |            |
         Confidence    Confidence    Confidence
          >= 85%       60-85%         < 60%
              |            |            |
        Automatisch   [PRUEFEN]_    [PRUEFEN]_
         abgelegt      markiert      markiert
              |
        Duplikat erkannt?
              |
        [DUPLIKAT]_
         markiert
              |
    2026/03_Maerz/
    Firma - CHF 99.90.pdf
              |
    Excel-Protokoll aktualisiert
```

## Schnellstart

### Neuinstallation

```bash
git clone https://github.com/chvofa/beleg-agent.git
cd beleg-agent
python setup_beleg_agent.py          # Windows
python3 setup_beleg_agent.py         # macOS (erstellt automatisch .venv)
```

Das Setup fragt interaktiv alles Noetige ab:
1. Installiert Abhaengigkeiten (Windows: global, macOS: in `.venv`)
2. Fragt den Anthropic API Key ab und speichert ihn (Windows: Umgebungsvariable, macOS: ~/.zshrc)
3. Fragt den Belege-Ordner ab und erstellt die Ordnerstruktur
4. Fragt die Bank ab (UBS, Raiffeisen oder PostFinance) und optional Kreditkartennummern
5. Richtet optional den Autostart ein (Windows: Startup-Ordner, macOS: LaunchAgent)
6. Bietet an, den Agent direkt zu starten

### Update (bestehende Installation)

```bash
cd beleg-agent
git pull
pip install -r requirements.txt         # Windows
.venv/bin/pip install -r requirements.txt  # macOS
```

Das Update ist sicher: `config_local.py` (deine lokalen Einstellungen), das Excel-Protokoll und die Belege-Ordner bleiben erhalten. Nur neue Abhaengigkeiten (z.B. `flask`) werden nachinstalliert. Danach den Agent neu starten.

> **Hinweis:** Falls der Agent per Autostart laeuft, einmal stoppen und neu starten, damit die neuen Dateien geladen werden.

## Ordnerstruktur

Nach dem Setup wird folgende Struktur im Belege-Ordner angelegt:

```
Belege/                          <-- dein gewaehlter Ordner
  _Inbox/                        <-- Belege hier ablegen
  _Abgleich/                     <-- KK/Bank-CSVs hier ablegen
    archiv/                      <-- verarbeitete CSVs werden hierhin verschoben
  _Dauerauftraege/               <-- wiederkehrende Rechnungen hier ablegen
  Belege_Protokoll.xlsx          <-- zentrale Uebersicht aller Belege
  2026/
    01_Januar/
      Firma A - CHF 150.00.pdf
      Firma B - EUR 49.99 KK EUR.pdf
    02_Februar/
      Gutschrift Firma C - CHF 200.00.pdf
    ...
```

## Starten

### Web-Interface + Tray-Icon (empfohlen)

| | Befehl |
|---|---|
| **Windows** | Doppelklick auf `start_beleg_agent.vbs` |
| **macOS** | `./start_beleg_agent.sh` oder `.venv/bin/python3 tray_agent.py` |

Was passiert:
- Das **Web-Interface** oeffnet sich automatisch unter `http://localhost:5001`
- Ein **Tray-Icon** erscheint in der Taskleiste/Menueleiste (zeigt Status + KPIs)
- Der **Watchdog** ueberwacht die Inbox und verarbeitet neue Belege automatisch

Das Tray-Icon zeigt:
- **Gruen** = Agent laeuft
- **Gelb** = Agent reagiert nicht mehr
- **Rot** = Agent gestoppt
- Rechtsklick/Klick: Web-Interface oeffnen, KPI-Ueberblick

### Nur Agent (ohne UI)

```bash
python beleg_agent.py                # Windows
.venv/bin/python3 beleg_agent.py     # macOS
```

### Autostart

| | Einrichtung |
|---|---|
| **Windows** | `Win + R` > `shell:startup` > Verknuepfung zu `start_beleg_agent.vbs` |
| **macOS** | Wird vom Setup automatisch eingerichtet (`~/Library/LaunchAgents/com.meocon.beleg-agent.plist`) |

Oder beim Setup mit "Ja" auf die Autostart-Frage antworten.

## Web-Interface

Das Web-Interface laeuft unter `http://localhost:5001` und bietet alle Funktionen in einer Oberflaeche:

| Seite | Funktion |
|-------|----------|
| **Dashboard** | Agent-Status, KPI-Kacheln (anklickbar), letzte Belege, Zeitstempel |
| **Upload** | Drag & Drop fuer PDF/JPG/PNG direkt in die Inbox |
| **Protokoll** | Alle Belege als Tabelle mit Filtern, Sortierung, Volltextsuche, Excel/CSV-Export. Klick auf das Dokument-Icon oeffnet den Beleg direkt. |
| **Abgleich** | KK/Bank-Abgleich und Dauerauftraege starten, CSV hochladen, Live-Output |
| **Pruefung** | [PRUEFEN]-Dateien korrigieren und freigeben oder ablehnen |
| **Logs** | Live-Log-Viewer |
| **Einstellungen** | Pfade, Bank-Profil, Schwellenwerte, API Key aendern |

Das Web-Interface ist nur lokal erreichbar (127.0.0.1) und mit CSRF-Schutz abgesichert.

> **Tipp:** Du brauchst die Inbox nicht ueber das Web-Interface zu befuellen. Belege koennen auch direkt per Adobe Scan, OneDrive Sync oder Finder/Explorer in den `_Inbox`-Ordner gelegt werden – der Watchdog erkennt sie automatisch.

## Verwendung

### 1. Belege verarbeiten

**Was tun:** PDF oder Bild (JPG/PNG) in den `_Inbox`-Ordner legen – oder im Web-Interface per Drag & Drop hochladen.

**Was passiert:**
- Der Agent erkennt die neue Datei automatisch (Watchdog)
- Claude Vision API analysiert den Beleg und extrahiert:
  - Rechnungssteller, Datum, Betrag, Waehrung
  - Zahlungsart (KK CHF, KK EUR, Ueberweisung)
  - Typ (Rechnung oder Gutschrift)
  - PayPal ja/nein
- Je nach Confidence-Score:
  - **>= 85%** – Automatisch umbenannt und in Monatsordner abgelegt
  - **60-85%** – Als `[PRUEFEN]_...` markiert, im Web-Interface korrigierbar
  - **< 60%** – Als `[PRUEFEN]_...` markiert
- Ein Eintrag wird im Excel-Protokoll erstellt
- Bei bekannten Rechnungsstellern wird die Zahlungsart automatisch aus der Historie uebernommen
- Duplikate werden erkannt und als `[DUPLIKAT]_...` markiert

**Dateinamens-Schema:**
```
[Gutschrift ][Rechnungssteller] - [Waehrung] [Betrag][ KK CHF/EUR].pdf
```

Beispiele:
```
Adobe - CHF 674.55 KK CHF.pdf
Amazon - EUR 49.99 KK EUR.pdf
Gutschrift Versicherung - CHF 200.00.pdf
AWS - USD 150.00 KK CHF.pdf
```

### 2. KK-Abgleich (Kreditkarten)

**Was tun:** Kreditkarten-Auszug als CSV in `_Abgleich` legen.

**Starten:** Im Web-Interface unter "Abgleich" oder via Terminal: `python abgleich.py`

**Was passiert:**
- CSV wird eingelesen (Format wird automatisch anhand des Bank-Profils erkannt)
- Waehrung erkannt (KK CHF oder KK EUR)
- Jede Transaktion wird mit dem Excel-Protokoll abgeglichen (Betrag + Name + Datum)
- Passende Belege werden als "Abgeglichen = Ja" markiert
- Fehlende Zahlungsart wird ergaenzt
- PayPal-Transaktionen werden erkannt
- **Fremdwaehrungen:** Bei USD-Rechnungen auf KK CHF/EUR wird der tatsaechlich belastete Betrag (`Betrag_Belastet`) und die Abrechnungswaehrung (`Waehrung_Belastet`) im Protokoll erfasst
- Transaktionen ohne passenden Beleg werden aufgelistet
- CSV wird ins `_Abgleich/archiv/` verschoben

### 3. Bank-Abgleich

**Was tun:** Bankauszug als CSV in `_Abgleich` legen.

**Starten:** Im Web-Interface unter "Abgleich" oder via Terminal: `python abgleich_bank.py`

**Was passiert:**
- CSV wird eingelesen (Format wird automatisch anhand des Bank-Profils erkannt)
- Belastungen und Gutschriften werden unterschieden
- Matching wie beim KK-Abgleich, aber mit mehr Datums-Toleranz (bis 45 Tage)
- Bankgebuehren und KK-Rechnungen werden automatisch uebersprungen
- CSV wird ins `_Abgleich/archiv/` verschoben

### 4. Dauerauftraege

**Was tun:** PDFs von wiederkehrenden Rechnungen (Miete, Leasing, Abos) in `_Dauerauftraege` legen.

**Starten:** Im Web-Interface unter "Abgleich" oder via Terminal: `python dauerauftraege.py`

**Was passiert:**
- Jedes PDF wird via Claude Vision analysiert
- Datei wird umbenannt: `Dauerauftrag [Name] - [Waehrung] [Betrag].pdf`
- Eintrag im Excel als Typ "Dauerauftrag" mit "Abgeglichen = Ja"
- Bereits erfasste Dateien werden uebersprungen

### 5. Excel-Protokoll

Die Datei `Belege_Protokoll.xlsx` ist die zentrale Uebersicht. Spaltenreihenfolge: Kerndaten, Finanzen, Abgleich, Bemerkungen, Metadaten.

| Spalte | Beschreibung |
|--------|--------------|
| Datum_Rechnung | Rechnungsdatum (YYYY-MM-DD) |
| Rechnungssteller | Name des Lieferanten |
| Typ | Rechnung / Gutschrift / Dauerauftrag |
| Betrag | Rechnungsbetrag |
| Waehrung | CHF, EUR, USD, etc. |
| Zahlungsart | KK CHF / KK EUR / Ueberweisung / eBill / leer |
| PayPal | Ja / Nein |
| Waehrung_Belastet | KK-Abrechnungswaehrung bei Fremdwaehrung |
| Betrag_Belastet | Tatsaechlich belasteter Betrag in KK-Waehrung |
| Abgeglichen | Ja / Nein (wird durch KK/Bank-Abgleich gesetzt) |
| Bemerkungen | Trinkgeld, Hinweise etc. |
| Originaldateiname | Urspruenglicher Dateiname |
| Ablagepfad | Wo die Datei jetzt liegt |
| Confidence_Score | Wie sicher die KI-Erkennung war |
| Verarbeitungsdatum | Wann der Beleg verarbeitet wurde |

Bestehende Excel-Dateien mit alter Spaltenreihenfolge werden beim Start automatisch migriert (Backup wird erstellt).

### 6. Erinnerungen

Der Agent prueft alle 6 Stunden automatisch:
- Wie lange kein Beleg abgelegt wurde (Warnung nach 14/30 Tagen)
- Wie viele Belege noch nicht abgeglichen sind
- Wann der letzte KK/Bank-Abgleich war
- Ob [PRUEFEN]- oder [DUPLIKAT]-Dateien in der Inbox liegen
- Am Monatsanfang: Erinnerung an eBill/Monatsberichte

## Bank-Profile

Der Agent unterstuetzt verschiedene Banken mit unterschiedlichen CSV-Formaten:

| Bank | Profil-Name | KK-Auszuege | Bankauszuege |
|------|-------------|-------------|--------------|
| UBS | `ubs` | Ja | Ja |
| Raiffeisen | `raiffeisen` | Ja | Ja |
| PostFinance | `postfinance` | Ja | Ja |

Das Bank-Profil wird in `config_local.py` oder im Web-Interface unter Einstellungen konfiguriert.

Optional koennen Kreditkartennummern hinterlegt werden, um die Zahlungsart automatisch zuzuordnen:

```python
BEKANNTE_KARTEN = {
    "1234": "KK CHF",
    "5678": "KK EUR",
}
```

## Konfiguration

Alle Einstellungen in `config.py` (Defaults) und `config_local.py` (deine Anpassungen):

| Einstellung | Standard | Beschreibung |
|-------------|----------|--------------|
| `ANTHROPIC_MODEL` | `claude-sonnet-4-20250514` | Claude-Modell fuer Vision API |
| `CONFIDENCE_AUTO` | `0.85` | Ab diesem Score: automatische Ablage |
| `CONFIDENCE_RUECKFRAGE` | `0.60` | Ab diesem Score: zur Pruefung markieren |
| `ERLAUBTE_ENDUNGEN` | `.pdf .jpg .jpeg .png` | Unterstuetzte Dateiformate |
| `BANK_PROFIL` | `ubs` | Bank-Profil fuer CSV-Import |

Diese Werte koennen auch bequem ueber das Web-Interface (Einstellungen) geaendert werden.

## Manuelle Installation

<details>
<summary>Falls du das Setup manuell durchfuehren willst</summary>

### 1. Abhaengigkeiten installieren

**Windows:**
```bash
pip install -r requirements.txt
```

**macOS:**
```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

### 2. API-Key setzen

**Windows (PowerShell):**
```powershell
[Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", "sk-ant-...", "User")
```

**macOS (Terminal):**
```bash
echo 'export ANTHROPIC_API_KEY="sk-ant-..."' >> ~/.zshrc
source ~/.zshrc
```

Oder ueber das Web-Interface: Setup-Wizard beim ersten Start, danach unter Einstellungen aenderbar.

### 3. Lokale Konfiguration erstellen

```bash
cp config_local.example.py config_local.py
```

Dann `config_local.py` bearbeiten und den Pfad zu deinem Belege-Ordner anpassen:

```python
# Windows:
ABLAGE_STAMMPFAD = r"C:\Users\DEIN_USER\Pfad\zu\Belege"

# macOS:
ABLAGE_STAMMPFAD = "/Users/DEIN_USER/Pfad/zu/Belege"
```

Unterordner (`_Inbox`, `_Abgleich`, `_Dauerauftraege`) werden automatisch erstellt.

</details>

## Plattform-Unterschiede

| | Windows | macOS |
|---|---|---|
| **Python** | Globale Installation | Virtuelle Umgebung (`.venv`) |
| **API Key** | Umgebungsvariable (User) | `~/.zshrc` + Fallback in App |
| **Autostart** | `shell:startup` Verknuepfung | LaunchAgent Plist |
| **Start-Befehl** | `start_beleg_agent.vbs` | `./start_beleg_agent.sh` |
| **Tray-Icon** | Taskleiste unten rechts | Menueleiste oben rechts |
| **Dock/Taskbar** | Kein extra Fenster (pythonw) | Kein Dock-Icon (LSUIElement) |

## Dateistruktur (Code)

```
beleg-agent/
  beleg_agent.py          # Hauptagent (Watchdog + Claude Vision + Ablage)
  web_app.py              # Web-Interface (Flask)
  tray_agent.py           # System Tray + Web-Server Launcher
  platform_utils.py       # Plattform-Abstraktion (Windows/macOS)
  config.py               # Allgemeine Konfiguration
  config_local.py         # Lokale Pfade und Bank-Profil (nicht im Git)
  config_local.example.py # Template fuer config_local.py
  bank_profile.py         # Bank-Profile (UBS, Raiffeisen, PostFinance)
  abgleich.py             # Kreditkarten-Abgleich
  abgleich_bank.py        # Bank-Abgleich
  dauerauftraege.py       # Dauerauftraege erfassen
  setup_beleg_agent.py    # Interaktives Setup
  templates/              # Jinja2-Templates (Dashboard, Upload, etc.)
  static/                 # CSS + JS
  start_beleg_agent.bat   # Startskript (Windows)
  start_beleg_agent.vbs   # Unsichtbarer Start (Windows)
  start_beleg_agent.sh    # Startskript (macOS)
  requirements.txt        # Python-Abhaengigkeiten
```

## Lizenz

MIT
