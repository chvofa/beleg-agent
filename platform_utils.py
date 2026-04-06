"""
Plattform-Abstraktion fuer Beleg-Agent (Windows + macOS).
Zentralisiert alle OS-spezifischen Aufrufe.
"""

import sys
import os
import subprocess

IS_WINDOWS = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"

# ── Subprocess-Flags ──────────────────────────────────────────────────────
# Verhindert Konsolen-Fenster auf Windows; auf Mac nicht noetig.
SUBPROCESS_FLAGS = subprocess.CREATE_NO_WINDOW if IS_WINDOWS else 0


# ── Dateien/Ordner oeffnen ────────────────────────────────────────────────

def open_file(path):
    """Oeffnet Datei/Ordner mit dem Standard-Programm des OS."""
    path = os.path.normpath(path)
    if IS_WINDOWS:
        os.startfile(path)
    elif IS_MAC:
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


# ── Fonts ─────────────────────────────────────────────────────────────────

def get_font_names():
    """Gibt (primaer, fallback) Font-Dateinamen fuer Pillow zurueck."""
    if IS_MAC:
        return (
            "/System/Library/Fonts/Helvetica.ttc",
            "/System/Library/Fonts/Supplemental/Arial.ttf",
        )
    return ("segoeuib.ttf", "arial.ttf")


# ── Toast-Benachrichtigungen ─────────────────────────────────────────────

def toast(title, msg, icon_path=None):
    """Plattformuebergreifende Toast-Benachrichtigung."""
    if IS_WINDOWS:
        try:
            from winotify import Notification
            t = Notification(
                app_id="Beleg-Agent", title=title, msg=msg, duration="short"
            )
            if icon_path and os.path.exists(icon_path):
                t.icon = icon_path
            t.show()
        except Exception:
            pass
    elif IS_MAC:
        try:
            # Sonderzeichen escapen fuer osascript
            safe_msg = msg.replace('\\', '\\\\').replace('"', '\\"').replace('\n', ' ')
            safe_title = title.replace('\\', '\\\\').replace('"', '\\"').replace('\n', ' ')
            script = (
                f'display notification "{safe_msg}" '
                f'with title "{safe_title}"'
            )
            subprocess.Popen(
                ["osascript", "-e", script],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        except Exception:
            pass


# ── API-Key Verwaltung ────────────────────────────────────────────────────

def get_api_key_from_env():
    """Liest ANTHROPIC_API_KEY mit OS-spezifischer Methode."""
    if IS_WINDOWS:
        try:
            result = subprocess.run(
                ["powershell", "-Command",
                 "[Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY', 'User')"],
                capture_output=True, text=True,
                creationflags=SUBPROCESS_FLAGS,
            )
            key = result.stdout.strip()
            if key:
                return key
        except Exception:
            pass
    # Fallback: Prozess-Umgebung (funktioniert auf Mac + als Windows-Fallback)
    return os.environ.get("ANTHROPIC_API_KEY", "")


def set_api_key_in_env(api_key):
    """Speichert ANTHROPIC_API_KEY persistent im OS. Gibt Statusmeldung zurueck."""
    if IS_WINDOWS:
        try:
            subprocess.run(
                ["powershell", "-Command",
                 f"[Environment]::SetEnvironmentVariable('ANTHROPIC_API_KEY', '{api_key}', 'User')"],
                capture_output=True,
                creationflags=SUBPROCESS_FLAGS,
            )
            os.environ["ANTHROPIC_API_KEY"] = api_key
            return "API Key als Windows-Umgebungsvariable gespeichert."
        except Exception:
            pass
    elif IS_MAC:
        try:
            zshrc = os.path.expanduser("~/.zshrc")
            marker = "# Beleg-Agent API Key"
            line = f'export ANTHROPIC_API_KEY="{api_key}"  {marker}'
            # Alten Eintrag entfernen falls vorhanden
            if os.path.exists(zshrc):
                with open(zshrc, "r") as f:
                    lines = [l for l in f.readlines() if marker not in l]
                with open(zshrc, "w") as f:
                    f.writelines(lines)
            with open(zshrc, "a") as f:
                f.write(f"\n{line}\n")
            os.environ["ANTHROPIC_API_KEY"] = api_key
            return "API Key in ~/.zshrc gespeichert. Neues Terminal oeffnen fuer Wirkung."
        except Exception:
            pass
    os.environ["ANTHROPIC_API_KEY"] = api_key
    return "API Key in Prozess-Umgebung gesetzt (nicht persistent)."
