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

KEYCHAIN_DIENST = "beleg_agent"
KEYCHAIN_KONTO = "ANTHROPIC_API_KEY"


def _keychain_lesen():
    """Liest den API-Key aus dem macOS-Keychain. Leerstring wenn nicht da."""
    try:
        result = subprocess.run(
            ["security", "find-generic-password",
             "-s", KEYCHAIN_DIENST, "-a", KEYCHAIN_KONTO, "-w"],
            capture_output=True, text=True, timeout=10,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
    return ""


def _keychain_schreiben(api_key):
    """Legt den API-Key im macOS-Keychain ab (-U ersetzt einen bestehenden).

    Der Key steht dabei kurz in der Kommandozeile von `security` und ist
    damit fuer `ps` sichtbar — security liest ihn nicht von stdin. Dafuer
    landet er nicht in der Shell-History wie beim Aufruf von Hand.
    """
    result = subprocess.run(
        ["security", "add-generic-password",
         "-s", KEYCHAIN_DIENST, "-a", KEYCHAIN_KONTO, "-U", "-w", api_key],
        capture_output=True, text=True, timeout=30,
    )
    if result.returncode != 0:
        meldung = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(meldung or "security add-generic-password fehlgeschlagen")


def get_api_key_from_env():
    """Liest ANTHROPIC_API_KEY: Umgebung zuerst, dann der OS-Speicher
    (macOS: Keychain, Windows: User-Umgebungsvariable)."""
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
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if key:
        return key
    # Die Umgebung ist leer, wenn der Agent nicht ueber start_beleg_agent.sh
    # oder direnv gestartet wurde — dann direkt im Keychain nachsehen.
    if IS_MAC:
        return _keychain_lesen()
    return ""


def set_api_key_in_env(api_key):
    """Speichert ANTHROPIC_API_KEY persistent im OS. Gibt Statusmeldung zurueck.

    Wirft bei Fehlschlag: ein stilles "gespeichert" ist schlimmer als ein
    Fehler, weil der Agent dann beim naechsten Start ohne Key hochfaehrt.
    """
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
        # In den Keychain, nicht nur in die Prozess-Umgebung: die stirbt mit
        # dem Setup, und launchd liest keine Shell-Profile. start_beleg_agent.sh
        # holt genau diesen Eintrag (security find-generic-password).
        _keychain_schreiben(api_key)
        os.environ["ANTHROPIC_API_KEY"] = api_key
        return f"API Key im macOS-Keychain gespeichert ({KEYCHAIN_DIENST}/{KEYCHAIN_KONTO})."
    os.environ["ANTHROPIC_API_KEY"] = api_key
    return "API Key in Prozess-Umgebung gesetzt (nicht persistent)."
