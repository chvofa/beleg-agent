#!/usr/bin/env python3
"""
Beleg-Agent – System Tray + Web Interface
Zeigt ein Status-Icon in der Menüleiste und startet das Web-Interface im Hintergrund.
"""

import os
import sys
import threading
import time
import webbrowser

from PIL import Image, ImageDraw, ImageFont
import pystray

import platform_utils

# ── macOS: Python-Icon aus Dock verstecken (reine Menubar-App) ────────────
if sys.platform == "darwin":
    try:
        from AppKit import NSBundle
        info = NSBundle.mainBundle().infoDictionary()
        info["LSUIElement"] = "1"
    except Exception:
        pass

# ── DPI Awareness (gegen Pixelation auf HiDPI-Displays) ───────────────────
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

# ── Pfade ──────────────────────────────────────────────────────────────────
AGENT_DIR = os.path.dirname(os.path.abspath(__file__))
STATUS_DATEI = os.path.join(AGENT_DIR, "beleg-agent.status")

# ── Web-Interface Port ────────────────────────────────────────────────────
WEB_PORT = 5001
WEB_URL = f"http://localhost:{WEB_PORT}"


# ── Icon erstellen ─────────────────────────────────────────────────────────

def erstelle_icon(farbe="green", size=256):
    """Erstellt ein hochaufgeloestes Icon."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    farben = {
        "green": (46, 160, 67),
        "red": (207, 34, 46),
        "yellow": (227, 149, 31),
        "gray": (110, 118, 129),
    }
    fill = farben.get(farbe, farben["gray"])

    margin = size // 16
    draw.ellipse(
        [margin, margin, size - margin, size - margin],
        fill=fill,
        outline=(255, 255, 255, 200),
        width=max(2, size // 32),
    )

    try:
        font_size = int(size * 0.5)
        primary, fallback = platform_utils.get_font_names()
        font = ImageFont.truetype(primary, font_size)
    except Exception:
        try:
            font = ImageFont.truetype(fallback, int(size * 0.5))
        except Exception:
            font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), "B", font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    x = (size - tw) // 2
    y = (size - th) // 2 - bbox[1]
    draw.text((x, y), "B", fill=(255, 255, 255), font=font)

    return img


# ── Tray + Web Application ───────────────────────────────────────────────

class BelegTray:
    def __init__(self):
        self.icon = None
        self._web_thread = None

    # ── Web-Server starten ────────────────────────────────────────────

    def _start_web_server(self):
        """Startet Flask + AgentController in einem Background-Thread."""
        from web_app import app, agent, WEB_PORT as port
        try:
            agent.start()
        except Exception as e:
            print(f"Agent-Start fehlgeschlagen: {e}")
        app.run(host="127.0.0.1", port=port, threaded=True, debug=False,
                use_reloader=False)

    # ── Icon-Status ───────────────────────────────────────────────────

    def update_icon(self, farbe):
        if self.icon:
            self.icon.icon = erstelle_icon(farbe)
            labels = {"green": "Laeuft", "red": "Gestoppt", "yellow": "Warnung"}
            self.icon.title = f"Beleg-Agent: {labels.get(farbe, farbe)}"

    def monitor_loop(self):
        """Aktualisiert das Icon basierend auf der Status-Datei."""
        while True:
            if os.path.exists(STATUS_DATEI):
                try:
                    alter = time.time() - os.path.getmtime(STATUS_DATEI)
                    if alter < 90:
                        self.update_icon("green")
                    elif alter < 300:
                        self.update_icon("yellow")
                    else:
                        self.update_icon("red")
                except Exception:
                    pass
            else:
                self.update_icon("red")
            time.sleep(15)

    # ── KPI-Daten laden ─────────────────────────────────────────────

    def _get_kpis(self):
        """Liest KPIs aus dem laufenden Web-Backend."""
        try:
            from web_app import _dashboard_stats
            return _dashboard_stats()
        except Exception:
            return {}

    def _kpi_label(self):
        stats = self._get_kpis()
        if not stats:
            return "Keine Daten"
        parts = []
        total = stats.get("total_belege", 0)
        offen = stats.get("nicht_abgeglichen", 0)
        pruefen = stats.get("pruefen_count", 0)
        parts.append(f"{total} Belege")
        if offen:
            parts.append(f"{offen} offen")
        if pruefen:
            parts.append(f"{pruefen} prüfen!")
        return " · ".join(parts)

    # ── Menü-Aktionen ─────────────────────────────────────────────────

    def on_open_web(self, icon, item):
        webbrowser.open(WEB_URL)

    def on_beenden(self, icon, item):
        from web_app import agent
        agent.stop()
        icon.stop()
        os._exit(0)

    # ── Main ──────────────────────────────────────────────────────────

    def run(self):
        # Web-Server im Hintergrund starten
        self._web_thread = threading.Thread(
            target=self._start_web_server, daemon=True
        )
        self._web_thread.start()

        # Monitor fuer Icon-Farbe
        threading.Thread(target=self.monitor_loop, daemon=True).start()

        # Browser oeffnen nach kurzer Wartezeit
        threading.Timer(2.0, lambda: webbrowser.open(WEB_URL)).start()

        # Tray-Menü mit KPI-Zeile
        menu = pystray.Menu(
            pystray.MenuItem(
                "Web-Interface öffnen", self.on_open_web, default=True
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                lambda _: self._kpi_label(), None, enabled=False
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Beenden", self.on_beenden),
        )

        self.icon = pystray.Icon(
            "beleg-agent",
            erstelle_icon("green"),
            "Beleg-Agent: Läuft",
            menu,
        )
        self.icon.run()


if __name__ == "__main__":
    tray = BelegTray()
    tray.run()
