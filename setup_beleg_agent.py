#!/usr/bin/env python3
"""
Beleg-Agent – Interaktives Setup
Fragt alle noetigen Einstellungen ab und richtet den Agent ein.
"""

import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import platform_utils


def frage(text, standard=""):
    """Fragt den Benutzer mit optionalem Standardwert."""
    if standard:
        eingabe = input(f"  {text} [{standard}]: ").strip().strip('"').strip("'")
        return eingabe if eingabe else standard
    else:
        while True:
            eingabe = input(f"  {text}: ").strip().strip('"').strip("'")
            if eingabe:
                return eingabe
            print("    Eingabe erforderlich.")


def ja_nein(text, standard=True):
    """Ja/Nein-Frage."""
    hint = "J/n" if standard else "j/N"
    eingabe = input(f"  {text} [{hint}]: ").strip().lower()
    if not eingabe:
        return standard
    return eingabe in ("j", "ja", "y", "yes")


def main():
    print()
    print("=" * 58)
    print("  BELEG-AGENT – Setup")
    print("=" * 58)
    print()

    agent_dir = os.path.dirname(os.path.abspath(__file__))

    # venv-Python fuer macOS, System-Python fuer Windows
    venv_dir = os.path.join(agent_dir, ".venv")
    if sys.platform == "darwin":
        venv_python = os.path.join(venv_dir, "bin", "python3")
    else:
        venv_python = sys.executable

    # ── 1. Abhaengigkeiten ────────────────────────────────────────────────
    print("[1/5] Abhaengigkeiten pruefen...\n")

    req_file = os.path.join(agent_dir, "requirements.txt")
    try:
        if sys.platform == "darwin":
            # macOS: Homebrew-Python verbietet System-Installationen (PEP 668)
            # → venv erstellen und Pakete darin installieren
            if not os.path.exists(venv_dir):
                print("  Erstelle virtuelle Umgebung (.venv)...")
                subprocess.check_call(
                    [sys.executable, "-m", "venv", venv_dir],
                )
            subprocess.check_call(
                [venv_python, "-m", "pip", "install", "-r", req_file],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            print("  Alle Abhaengigkeiten in .venv installiert.\n")
        else:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "-r", req_file],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            print("  Alle Abhaengigkeiten installiert.\n")
    except subprocess.CalledProcessError:
        print("  WARNUNG: Einige Pakete konnten nicht installiert werden.")
        if sys.platform == "darwin":
            print("  Bitte manuell ausfuehren:")
            print("    python3 -m venv .venv && .venv/bin/pip install -r requirements.txt\n")
        else:
            print("  Bitte manuell ausfuehren: pip install -r requirements.txt\n")

    # ── 2. API-Key ────────────────────────────────────────────────────────
    print("[2/5] Anthropic API Key\n")

    # Pruefen ob bereits gesetzt
    bestehender_key = platform_utils.get_api_key_from_env()

    if bestehender_key:
        maskiert = bestehender_key[:7] + "..." + bestehender_key[-4:]
        print(f"  API Key bereits gesetzt: {maskiert}")
        if not ja_nein("Neuen Key eingeben?", standard=False):
            api_key = bestehender_key
        else:
            api_key = frage("Anthropic API Key (sk-ant-...)")
    else:
        print("  Kein API Key gefunden.")
        print("  Erstelle einen unter: https://console.anthropic.com/\n")
        api_key = frage("Anthropic API Key (sk-ant-...)")

    if api_key and api_key != bestehender_key:
        try:
            meldung = platform_utils.set_api_key_in_env(api_key)
            print(f"  {meldung}\n")
        except Exception as e:
            print(f"  WARNUNG: Konnte Key nicht setzen: {e}")
            print("  Bitte manuell setzen (siehe README.md)\n")

    # ── 3. Lokale Konfiguration ───────────────────────────────────────────
    print("[3/5] Belege-Ordner & Bank konfigurieren\n")

    config_local_pfad = os.path.join(agent_dir, "config_local.py")

    if os.path.exists(config_local_pfad):
        # Bestehenden Pfad auslesen
        try:
            # config_local.py als Modul laden
            import importlib.util
            spec = importlib.util.spec_from_file_location("config_local_check", config_local_pfad)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            bestehender_pfad = getattr(mod, "ABLAGE_STAMMPFAD", "")
        except Exception:
            bestehender_pfad = ""

        if bestehender_pfad:
            print(f"  Aktueller Pfad: {bestehender_pfad}")
            if not ja_nein("Aendern?", standard=False):
                belege_pfad = bestehender_pfad
            else:
                belege_pfad = frage("Pfad zum Belege-Ordner")
        else:
            belege_pfad = frage("Pfad zum Belege-Ordner")
    else:
        print("  Der Belege-Ordner ist der Stammordner fuer alle Belege.")
        print("  Unterordner (_Inbox, _Abgleich, etc.) werden automatisch erstellt.\n")
        if sys.platform == "darwin":
            beispiel = "/Users/Max/Belege"
        else:
            beispiel = r"C:\Users\Max\Belege"
        belege_pfad = frage(f"Pfad zum Belege-Ordner (z.B. {beispiel})")

    # ── 3b. Bank-Profil ─────────────────────────────────────────────────
    print("  Welche Bank nutzt du fuer KK- und Bank-Abgleich?\n")
    print("    1) UBS")
    print("    2) Raiffeisen")
    print("    3) PostFinance")
    print("    4) Andere / weiss nicht")
    print()

    bank_wahl = frage("Auswahl (1-4)", "1")
    bank_map = {"1": "ubs", "2": "raiffeisen", "3": "postfinance", "4": "ubs"}
    bank_profil = bank_map.get(bank_wahl, "ubs")
    if bank_wahl == "4":
        print("  Standard-Profil 'ubs' gewaehlt. Kann spaeter in config_local.py angepasst werden.")
        print("  Neue Profile koennen in bank_profile.py ergaenzt werden.\n")
    else:
        print(f"  Bank-Profil: {bank_profil}\n")

    # ── 3c. Kreditkarten ─────────────────────────────────────────────────
    print("  Optional: Kreditkarten konfigurieren (fuer automatische Zahlungsart-Erkennung)")
    print("  Gib die letzten 4 Ziffern deiner KK(s) ein, oder Enter zum Ueberspringen.\n")

    karten = {}
    kk1 = input("  KK mit CHF-Abrechnung (letzte 4 Ziffern, oder Enter): ").strip()
    if kk1 and len(kk1) == 4 and kk1.isdigit():
        karten[kk1] = "KK CHF"
    kk2 = input("  KK mit EUR-Abrechnung (letzte 4 Ziffern, oder Enter): ").strip()
    if kk2 and len(kk2) == 4 and kk2.isdigit():
        karten[kk2] = "KK EUR"
    print()

    # Tilde expandieren (macOS: ~/Belege → /Users/xyz/Belege)
    belege_pfad = os.path.expanduser(belege_pfad)

    # config_local.py schreiben
    with open(config_local_pfad, "w", encoding="utf-8") as f:
        f.write('"""\n')
        f.write("Lokale Konfiguration – NICHT ins Git committen!\n")
        f.write('"""\n\n')
        f.write(f'ABLAGE_STAMMPFAD = r"{belege_pfad}"\n')
        f.write(f'\n# Bank-Profil: "ubs", "raiffeisen", "postfinance"\n')
        f.write(f'BANK_PROFIL = "{bank_profil}"\n')
        if karten:
            f.write(f'\n# Bekannte Kreditkarten (letzte 4 Ziffern → KK-Typ)\n')
            f.write(f'BEKANNTE_KARTEN = {repr(karten)}\n')

    print(f"  config_local.py gespeichert.\n")

    # Ordner erstellen
    for unterordner in ["", "_Inbox", "_Abgleich", "_Dauerauftraege"]:
        pfad = os.path.join(belege_pfad, unterordner)
        os.makedirs(pfad, exist_ok=True)

    print("  Ordnerstruktur erstellt:")
    print(f"    {belege_pfad}")
    print(f"    ├── _Inbox")
    print(f"    ├── _Abgleich")
    print(f"    └── _Dauerauftraege\n")

    # ── 4. Autostart ──────────────────────────────────────────────────────
    print("[4/5] Autostart\n")

    if sys.platform == "win32":
        autostart_frage = "Soll der Agent automatisch bei Windows-Start starten?"
    else:
        autostart_frage = "Soll der Agent automatisch bei Mac-Start starten?"

    if ja_nein(autostart_frage):
        try:
            if sys.platform == "win32":
                startup_pfad = os.path.join(
                    os.environ.get("APPDATA", ""),
                    r"Microsoft\Windows\Start Menu\Programs\Startup"
                )
                vbs_quelle = os.path.join(agent_dir, "start_beleg_agent.vbs")
                vbs_ziel = os.path.join(startup_pfad, "start_beleg_agent.vbs")

                if os.path.exists(startup_pfad) and os.path.exists(vbs_quelle):
                    import shutil
                    shutil.copy2(vbs_quelle, vbs_ziel)
                    print(f"  Autostart-Verknuepfung erstellt.\n")
                else:
                    print("  WARNUNG: Startup-Ordner oder VBS-Datei nicht gefunden.")
                    print(f"  Manuell: {vbs_quelle} nach {startup_pfad} kopieren.\n")
            elif sys.platform == "darwin":
                plist_dir = os.path.expanduser("~/Library/LaunchAgents")
                os.makedirs(plist_dir, exist_ok=True)
                plist_pfad = os.path.join(plist_dir, "com.meocon.beleg-agent.plist")
                python_exe = venv_python
                script_pfad = os.path.join(agent_dir, "tray_agent.py")
                plist_inhalt = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.meocon.beleg-agent</string>
    <key>ProgramArguments</key>
    <array>
        <string>{python_exe}</string>
        <string>{script_pfad}</string>
    </array>
    <key>WorkingDirectory</key>
    <string>{agent_dir}</string>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <false/>
    <key>EnvironmentVariables</key>
    <dict>
        <key>PATH</key>
        <string>/usr/local/bin:/usr/bin:/bin:/opt/homebrew/bin</string>
    </dict>
</dict>
</plist>"""
                with open(plist_pfad, "w") as f:
                    f.write(plist_inhalt)
                print(f"  LaunchAgent erstellt: {plist_pfad}")
                print("  Agent startet automatisch beim naechsten Login.\n")
        except Exception as e:
            print(f"  WARNUNG: Konnte Autostart nicht einrichten: {e}\n")
    else:
        print("  Autostart uebersprungen.\n")

    # ── Fertig ────────────────────────────────────────────────────────────
    print("=" * 58)
    print("  Setup abgeschlossen!")
    print("=" * 58)
    print()
    print("  Starten mit:")
    if sys.platform == "darwin":
        print(f"    .venv/bin/python3 tray_agent.py")
        print("  Oder: ./start_beleg_agent.sh")
    else:
        print(f"    python {os.path.join(agent_dir, 'tray_agent.py')}")
        print("  Oder Doppelklick auf start_beleg_agent.vbs")
    print()
    print("  Belege in _Inbox legen – der Rest passiert automatisch.")
    print()

    if ja_nein("Agent jetzt starten?"):
        print("  Starte Beleg-Agent...\n")
        subprocess.Popen(
            [venv_python, os.path.join(agent_dir, "tray_agent.py")],
            cwd=agent_dir,
        )


if __name__ == "__main__":
    main()
