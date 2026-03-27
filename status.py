#!/usr/bin/env python3
"""Prüft ob der Beleg-Agent läuft."""
import os, sys
from datetime import datetime

STATUS_DATEI = os.path.join(os.path.dirname(os.path.abspath(__file__)), "beleg-agent.status")

if not os.path.exists(STATUS_DATEI):
    print("Beleg-Agent: NICHT GESTARTET (keine Status-Datei)")
    sys.exit(1)

with open(STATUS_DATEI, "r") as f:
    inhalt = f.read()

print(inhalt)

# Prüfe ob Status aktuell ist (< 60s alt)
mtime = os.path.getmtime(STATUS_DATEI)
alter = datetime.now().timestamp() - mtime
if alter > 60:
    print(f"WARNUNG: Status ist {alter:.0f}s alt - Agent laeuft vermutlich NICHT!")
else:
    print(f"Status ist {alter:.0f}s alt - Agent laeuft.")
