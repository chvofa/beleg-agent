#!/usr/bin/env python3
"""
Beleg-Agent – System Tray Launcher
Startet den Beleg-Agent mit System Tray Icon und dunklem Custom-Menü.
"""

import os
import sys
import threading
import subprocess
import time
import ctypes
from datetime import datetime

from PIL import Image, ImageDraw, ImageFont
import pystray

# ── DPI Awareness (gegen Pixelation auf HiDPI-Displays) ───────────────────
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-Monitor DPI Aware
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# ── Pfade ──────────────────────────────────────────────────────────────────
AGENT_DIR = os.path.dirname(os.path.abspath(__file__))
STATUS_DATEI = os.path.join(AGENT_DIR, "beleg-agent.status")
LOG_DATEI = os.path.join(AGENT_DIR, "beleg-agent.log")
PYTHON_EXE = sys.executable
AGENT_SCRIPT = os.path.join(AGENT_DIR, "beleg_agent.py")
ABGLEICH_SCRIPT = os.path.join(AGENT_DIR, "abgleich.py")
ABGLEICH_BANK_SCRIPT = os.path.join(AGENT_DIR, "abgleich_bank.py")
DAUERAUFTRAEGE_SCRIPT = os.path.join(AGENT_DIR, "dauerauftraege.py")

import config

INBOX_PFAD = config.INBOX_PFAD
ABGLEICH_PFAD = config.ABGLEICH_PFAD
DAUERAUFTRAEGE_PFAD = os.path.join(config.ABLAGE_STAMMPFAD, "_Dauerauftraege")
EXCEL_PFAD = config.EXCEL_PROTOKOLL


# ── Icon erstellen ─────────────────────────────────────────────────────────

def erstelle_icon(farbe="green", size=256):
    """Erstellt ein hochaufgelöstes Icon."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    farben = {
        "green": (46, 160, 67),
        "red": (207, 34, 46),
        "yellow": (227, 149, 31),
        "gray": (110, 118, 129),
    }
    fill = farben.get(farbe, farben["gray"])

    # Abgerundeter Kreis
    margin = size // 16
    draw.ellipse(
        [margin, margin, size - margin, size - margin],
        fill=fill,
        outline=(255, 255, 255, 200),
        width=max(2, size // 32),
    )

    # "B" in der Mitte
    try:
        font_size = int(size * 0.5)
        font = ImageFont.truetype("segoeuib.ttf", font_size)  # Segoe UI Bold
    except Exception:
        try:
            font = ImageFont.truetype("arial.ttf", int(size * 0.5))
        except Exception:
            font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), "B", font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    x = (size - tw) // 2
    y = (size - th) // 2 - bbox[1]
    draw.text((x, y), "B", fill=(255, 255, 255), font=font)

    return img


# ── Toast Helper ───────────────────────────────────────────────────────────

def toast(title, msg):
    try:
        from winotify import Notification
        t = Notification(app_id="Beleg-Agent", title=title, msg=msg, duration="short")
        t.show()
    except Exception:
        pass


# ── Tray Application ──────────────────────────────────────────────────────

class BelegTray:
    def __init__(self):
        self.process = None
        self.running = False
        self.icon = None

    # ── Agent-Steuerung ────────────────────────────────────────────────

    def _hole_env(self):
        env = os.environ.copy()
        try:
            result = subprocess.run(
                ["powershell", "-Command",
                 "[Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY', 'User')"],
                capture_output=True, text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            key = result.stdout.strip()
            if key:
                env["ANTHROPIC_API_KEY"] = key
        except Exception:
            pass
        return env

    def start_agent(self):
        if self.process and self.process.poll() is None:
            return
        self.process = subprocess.Popen(
            [PYTHON_EXE, AGENT_SCRIPT],
            cwd=AGENT_DIR,
            env=self._hole_env(),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        self.running = True
        self.update_icon("green")

    def stop_agent(self):
        if self.process and self.process.poll() is None:
            self.process.terminate()
            try:
                self.process.wait(timeout=5)
            except Exception:
                self.process.kill()
        self.running = False
        self.update_icon("red")

    def update_icon(self, farbe):
        if self.icon:
            self.icon.icon = erstelle_icon(farbe)
            labels = {"green": "Laeuft", "red": "Gestoppt", "yellow": "Warnung"}
            self.icon.title = f"Beleg-Agent: {labels.get(farbe, farbe)}"

    # ── Script-Runner ──────────────────────────────────────────────────

    def _run_script(self, script_pfad, name):
        def _run():
            toast(name, "Wird ausgefuehrt...")
            try:
                result = subprocess.run(
                    [PYTHON_EXE, script_pfad],
                    cwd=AGENT_DIR,
                    env=self._hole_env(),
                    capture_output=True, text=True,
                    encoding="utf-8", errors="replace",
                    timeout=300,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
                output = result.stdout.strip()
                zeilen = output.split("\n")
                zusammenfassung = ""
                for z in zeilen:
                    if any(kw in z for kw in ["Abgeglichen:", "ergaenzt:", "Ohne Beleg:", "Fertig", "Erfolgreich", "Fehler:"]):
                        zusammenfassung += z.strip() + "\n"
                if not zusammenfassung:
                    zusammenfassung = "\n".join(zeilen[-3:])
                toast(f"{name} fertig", zusammenfassung[:200])
            except subprocess.TimeoutExpired:
                toast(f"{name}", "Timeout nach 5 Minuten")
            except Exception as e:
                toast(f"{name}", str(e)[:200])
        threading.Thread(target=_run, daemon=True).start()

    # ── Monitor ────────────────────────────────────────────────────────

    def monitor_loop(self):
        while True:
            if self.process and self.process.poll() is not None:
                self.running = False
                self.update_icon("red")
            elif self.running and os.path.exists(STATUS_DATEI):
                alter = time.time() - os.path.getmtime(STATUS_DATEI)
                self.update_icon("green" if alter < 90 else "yellow")
            time.sleep(15)

    # ── Menü-Aktionen ──────────────────────────────────────────────────

    def on_status(self, icon, item):
        if os.path.exists(STATUS_DATEI):
            with open(STATUS_DATEI, "r") as f:
                status = f.read().strip()
        else:
            status = "Agent nicht gestartet"
        toast("Beleg-Agent Status", status)

    def on_start(self, icon, item):
        self.start_agent()
        toast("Beleg-Agent", "Agent gestartet")

    def on_stop(self, icon, item):
        self.stop_agent()
        toast("Beleg-Agent", "Agent gestoppt")

    def on_restart(self, icon, item):
        self.stop_agent()
        time.sleep(2)
        self.start_agent()
        toast("Beleg-Agent", "Agent neu gestartet")

    def on_kk_abgleich(self, icon, item):
        self._run_script(ABGLEICH_SCRIPT, "KK-Abgleich")

    def on_bank_abgleich(self, icon, item):
        self._run_script(ABGLEICH_BANK_SCRIPT, "Bank-Abgleich")

    def on_dauerauftraege(self, icon, item):
        self._run_script(DAUERAUFTRAEGE_SCRIPT, "Dauerauftraege")

    def on_open_inbox(self, icon, item):
        os.startfile(os.path.normpath(INBOX_PFAD))

    def on_open_excel(self, icon, item):
        os.startfile(os.path.normpath(EXCEL_PFAD))

    def on_open_abgleich(self, icon, item):
        os.startfile(os.path.normpath(ABGLEICH_PFAD))

    def on_open_dauerauftraege(self, icon, item):
        os.startfile(os.path.normpath(DAUERAUFTRAEGE_PFAD))

    def on_open_log(self, icon, item):
        os.startfile(os.path.normpath(LOG_DATEI))

    def on_beenden(self, icon, item):
        self.stop_agent()
        icon.stop()

    # ── Main ───────────────────────────────────────────────────────────

    def run(self):
        self.start_agent()

        threading.Thread(target=self.monitor_loop, daemon=True).start()

        menu = pystray.Menu(
            pystray.MenuItem("Status anzeigen", self.on_status, default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Abgleich", pystray.Menu(
                pystray.MenuItem("KK-Abgleich starten", self.on_kk_abgleich),
                pystray.MenuItem("Bank-Abgleich starten", self.on_bank_abgleich),
                pystray.MenuItem("Dauerauftraege erfassen", self.on_dauerauftraege),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Abgleich-Ordner oeffnen", self.on_open_abgleich),
                pystray.MenuItem("Dauerauftraege-Ordner oeffnen", self.on_open_dauerauftraege),
            )),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Inbox oeffnen", self.on_open_inbox),
            pystray.MenuItem("Excel oeffnen", self.on_open_excel),
            pystray.MenuItem("Log oeffnen", self.on_open_log),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Agent", pystray.Menu(
                pystray.MenuItem("Starten", self.on_start),
                pystray.MenuItem("Stoppen", self.on_stop),
                pystray.MenuItem("Neustarten", self.on_restart),
            )),
            pystray.MenuItem("Beenden", self.on_beenden),
        )

        self.icon = pystray.Icon(
            "beleg-agent",
            erstelle_icon("green"),
            "Beleg-Agent: Laeuft",
            menu,
        )
        self.icon.run()


if __name__ == "__main__":
    tray = BelegTray()
    tray.run()
