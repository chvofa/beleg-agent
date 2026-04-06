# Beleg-Agent

Automatische Verarbeitung von Rechnungen und Belegen mit Claude Vision API.

Der Agent ueberwacht einen Inbox-Ordner, extrahiert Rechnungsdaten via KI und legt Belege strukturiert ab. Kreditkarten- und Bank-Transaktionen werden automatisch mit den erfassten Belegen abgeglichen. Unterstuetzt mehrere Banken (UBS, Raiffeisen, PostFinance) und Fremdwaehrungstransaktionen.

## So funktioniert es

```
                         _Inbox/
                           |
                     PDF oder Bild ablegen
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

```bash
git clone https://github.com/chvofa/beleg-agent.git
cd beleg-agent
python setup_beleg_agent.py
```

Das Setup fragt interaktiv alles Noetige ab:
1. Installiert Abhaengigkeiten
2. Fragt den Anthropic API Key ab und speichert ihn als Windows-Umgebungsvariable
3. Fragt den Belege-Ordner ab und erstellt die Ordnerstruktur
4. Fragt die Bank ab (UBS, Raiffeisen oder PostFinance) und optional Kreditkartennummern
5. Richtet optional den Windows-Autostart ein
6. Bietet an, den Agent direkt zu starten

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

### Option A: System Tray (empfohlen)

Doppelklick auf `start_beleg_agent.vbs` - startet den Agent unsichtbar im Hintergrund mit Tray-Icon.

Das Tray-Icon zeigt den Status:
- Gruen = Agent laeuft
- Gelb = Agent reagiert nicht mehr
- Rot = Agent gestoppt

Rechtsklick auf das Tray-Icon oeffnet das Menue mit allen Funktionen.

### Option B: Terminal

```bash
python beleg_agent.py
```

### Autostart bei Windows-Start

1. `Win + R` > `shell:startup`
2. Verknuepfung zu `start_beleg_agent.vbs` in den Autostart-Ordner legen

Oder beim Setup mit "Ja" auf die Autostart-Frage antworten.

## Verwendung

### 1. Belege verarbeiten

**Was tun:** PDF oder Bild (JPG/PNG) in den `_Inbox`-Ordner legen.

**Was passiert:**
- Der Agent erkennt die neue Datei automatisch (Watchdog)
- Claude Vision API analysiert den Beleg und extrahiert:
  - Rechnungssteller, Datum, Betrag, Waehrung
  - Zahlungsart (KK CHF, KK EUR, Ueberweisung)
  - Typ (Rechnung oder Gutschrift)
  - PayPal ja/nein
- Je nach Confidence-Score:
  - **>= 85%** - Datei wird automatisch umbenannt und in den richtigen Monatsordner verschoben
  - **60-85%** - Toast-Benachrichtigung, Datei wird als `[PRUEFEN]_...` markiert
  - **< 60%** - Datei wird als `[PRUEFEN]_...` markiert
- Ein Eintrag wird im Excel-Protokoll erstellt
- Bei bekannten Rechnungsstellern wird die Zahlungsart automatisch aus der Historie uebernommen
- Duplikate werden erkannt und als `[DUPLIKAT]_...` markiert (koennen geloescht werden)

**Dateinamens-Schema:**
```
[Gutschrift ][Rechnungssteller] - [Waehrung] [Betrag][ KK CHF/EUR].pdf
```

Beispiele:
```
Adobe - CHF 674.55 KK CHF.pdf
Amazon - EUR 49.99 KK EUR.pdf
Gutschrift Versicherung - CHF 200.00.pdf
Hosting - USD 29.00.pdf
AWS - USD 150.00 KK CHF.pdf
```

**[PRUEFEN]-Dateien:** Belege mit niedrigem Confidence bleiben in der `_Inbox` mit dem Prefix `[PRUEFEN]_`. Die Toast-Nachricht zeigt den Grund an (z.B. Confidence-Wert).

**[DUPLIKAT]-Dateien:** Wenn ein Beleg bereits im Protokoll existiert, wird die Datei als `[DUPLIKAT]_...` markiert. Diese koennen bedenkenlos geloescht werden.

### 2. KK-Abgleich (Kreditkarten)

**Was tun:** Kreditkarten-Auszug als CSV in `_Abgleich` legen.

**Starten:**
- Tray-Menue > Abgleich > KK-Abgleich starten
- Oder: `python abgleich.py`

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

**Starten:**
- Tray-Menue > Abgleich > Bank-Abgleich starten
- Oder: `python abgleich_bank.py`

**Was passiert:**
- CSV wird eingelesen (Format wird automatisch anhand des Bank-Profils erkannt)
- Belastungen und Gutschriften werden unterschieden
- Matching wie beim KK-Abgleich, aber mit mehr Datums-Toleranz (bis 45 Tage)
- Bankgebuehren und KK-Rechnungen werden automatisch uebersprungen
- CSV wird ins `_Abgleich/archiv/` verschoben

### 4. Dauerauftraege

**Was tun:** PDFs von wiederkehrenden Rechnungen (Miete, Leasing, Abos) in `_Dauerauftraege` legen.

**Starten:**
- Tray-Menue > Abgleich > Dauerauftraege erfassen
- Oder: `python dauerauftraege.py`

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
| Waehrung_Belastet | KK-Abrechnungswaehrung bei Fremdwaehrung (z.B. CHF bei USD-Rechnung auf KK CHF) |
| Betrag_Belastet | Tatsaechlich belasteter Betrag in KK-Waehrung (inkl. Wechselkurs) |
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

## Tray-Menue

Rechtsklick auf das Tray-Icon:

```
Status anzeigen          <-- zeigt aktuellen Agent-Status
---
Abgleich >
  KK-Abgleich starten
  Bank-Abgleich starten
  Dauerauftraege erfassen
  ---
  Abgleich-Ordner oeffnen
  Dauerauftraege-Ordner oeffnen
---
Inbox oeffnen
Excel oeffnen
Log oeffnen
---
Agent >
  Starten
  Stoppen
  Neustarten
Hilfe                    <-- oeffnet README auf GitHub
Beenden
```

## Bank-Profile

Der Agent unterstuetzt verschiedene Banken mit unterschiedlichen CSV-Formaten:

| Bank | Profil-Name | KK-Auszuege | Bankauszuege |
|------|-------------|-------------|--------------|
| UBS | `ubs` | Ja | Ja |
| Raiffeisen | `raiffeisen` | Ja | Ja |
| PostFinance | `postfinance` | Ja | Ja |

Das Bank-Profil wird in `config_local.py` konfiguriert:

```python
BANK_PROFIL = "ubs"  # oder "raiffeisen" / "postfinance"
```

Optional koennen Kreditkartennummern hinterlegt werden, um die Zahlungsart automatisch zuzuordnen:

```python
BEKANNTE_KARTEN = {
    "1234": "KK CHF",
    "5678": "KK EUR",
}
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
| `BANK_PROFIL` | `ubs` | Bank-Profil fuer CSV-Import (siehe oben) |

## Manuelle Installation

<details>
<summary>Falls du das Setup manuell durchfuehren willst</summary>

### 1. Abhaengigkeiten installieren

```bash
pip install -r requirements.txt
```

### 2. API-Key setzen

```powershell
[Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", "sk-ant-...", "User")
```

### 3. Lokale Konfiguration erstellen

```bash
copy config_local.example.py config_local.py
```

Dann `config_local.py` bearbeiten und den Pfad zu deinem Belege-Ordner anpassen:

```python
ABLAGE_STAMMPFAD = r"C:\Users\DEIN_USER\Pfad\zu\Belege"
```

Unterordner (`_Inbox`, `_Abgleich`, `_Dauerauftraege`) werden automatisch erstellt.

</details>

## Dateistruktur (Code)

```
beleg-agent/
  beleg_agent.py          # Hauptagent (Watchdog + Claude Vision + Ablage)
  tray_agent.py           # System Tray Launcher
  config.py               # Allgemeine Konfiguration
  config_local.py         # Lokale Pfade und Bank-Profil (nicht im Git)
  config_local.example.py # Template fuer config_local.py
  bank_profile.py         # Bank-Profile (UBS, Raiffeisen, PostFinance)
  abgleich.py             # Kreditkarten-Abgleich
  abgleich_bank.py        # Bank-Abgleich
  dauerauftraege.py       # Dauerauftraege erfassen
  status.py               # Status-Checker
  setup_beleg_agent.py    # Interaktives Setup
  start_beleg_agent.bat   # Startskript
  start_beleg_agent.vbs   # Unsichtbarer Start
  requirements.txt        # Python-Abhaengigkeiten
```

## Lizenz

MIT
