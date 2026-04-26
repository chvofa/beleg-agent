---
name: Tray nach Code-Aenderung neu starten
description: Python-Modul-Cache — beim Bearbeiten von abgleich*.py / offene_posten.py / beleg_agent.py muss der tray_agent.py-Prozess neu gestartet werden, sonst laufen alte Versionen
type: feedback
originSessionId: 73c7fa4b-622e-4279-bd84-e0946fb1d309
---
Nach einer Aenderung an Python-Modulen (besonders `abgleich.py`, `abgleich_bank.py`, `offene_posten.py`, `beleg_agent.py`, `bank_profile.py`) muss der Tray-Agent neu gestartet werden, bevor man die Aenderung testet.

**Why:** Der tray_agent.py-Prozess hat die Module nach dem ersten Aufruf im `sys.modules`-Cache. Flask/pystray reloaden nichts. Ich bin in dieser Session mehrfach darauf reingefallen — Aenderung gemacht, Bank-Abgleich getriggert, altes Verhalten bekommen, Minuten mit Debugging verloren. Einmal hat sogar der Homebrew-Python-Upgrade das stdlib unter dem laufenden Prozess weggezogen (ModuleNotFoundError: 'csv'), was nur ein Neustart behoben hat.

**How to apply:**
- Vor jedem Test nach Code-Aenderung: `pgrep -f tray_agent.py | grep -v zsh | xargs -I{} kill {}` und dann `.venv/bin/python tray_agent.py` im Hintergrund neu starten.
- Bei "komischem" Verhalten (z.B. Cleanup entfernt weniger als erwartet) immer zuerst den Tray neu starten, bevor man den Code weiter debuggt.
- API-Trigger via `curl -s -X POST http://127.0.0.1:5001/api/reconciliation/{bank,kk}`.
