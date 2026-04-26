---
name: Fabios Setup fuer Beleg-Agent
description: Bank/Karten, Datenablage und Lauf-Umgebung — damit ich in neuen Gespraechen nicht blind loslaufe
type: user
originSessionId: 73c7fa4b-622e-4279-bd84-e0946fb1d309
---
**Wer:** Fabio Oro, Meocon GmbH. Nutzt den Beleg-Agent produktiv fuer die eigene Buchhaltung und baut ihn gleichzeitig als Produkt fuer andere Kunden weiter.

**Bank/Karten:**
- BANK_PROFIL = "ubs" (UBS Bankkonto + UBS Kreditkarte)
- Primaere Waehrungen: CHF und EUR
- UBS Bank-CSV hat Transaktions-Nr (wird seit Commit b381e5b als Referenz genutzt)
- UBS KK-CSV hat KEINE Transaktions-Nr — fuer KK kein Ref-basiertes Matching moeglich

**Datenablage:**
- Excel-Protokoll liegt in OneDrive: `~/Library/CloudStorage/OneDrive-FreigegebeneBibliotheken–MeoconGmbH/Meocon Admin - Documents/02 - Meocon/02 - Financials/01-Belege/Belege_Protokoll.xlsx`
- OneDrive gibt gelegentlich PermissionError bei openpyxl — typischerweise Sync-Lock, loest sich von allein
- `_Abgleich/` und `_Abgleich/archiv/` sind Unterordner vom gleichen OneDrive-Pfad

**Lauf-Umgebung:**
- macOS, Python 3.14 aus Homebrew (wird regelmaessig durch brew upgrade ersetzt → alter stdlib-Pfad ist weg, Tray muss dann neu gestartet werden)
- `.venv` im Repo-Root, `.venv/bin/python` startet `tray_agent.py`
- Tray laeuft als macOS-Menubar-App (LSUIElement=1), Web-UI auf http://127.0.0.1:5001

**Parallel auf Windows:** Excel-Protokoll wird auch von einer Windows-Instanz beschrieben. Die Ablagepfade im Protokoll sind deshalb oft Windows-Pfade `C:\Users\FabioOro\...` — bei macOS-Zugriff muessen die per `_resolve_ablagepfad` umgemappt werden.
