#!/usr/bin/env python3
"""
Beleg-Agent – Interaktives Setup
Fragt alle noetigen Einstellungen ab und richtet den Agent ein.
"""

import os
import subprocess
import sys


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

    # ── 1. Abhaengigkeiten ────────────────────────────────────────────────
    print("[1/4] Abhaengigkeiten pruefen...\n")

    req_file = os.path.join(agent_dir, "requirements.txt")
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "-r", req_file],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        print("  Alle Abhaengigkeiten installiert.\n")
    except subprocess.CalledProcessError:
        print("  WARNUNG: Einige Pakete konnten nicht installiert werden.")
        print("  Bitte manuell ausfuehren: pip install -r requirements.txt\n")

    # ── 2. API-Key ────────────────────────────────────────────────────────
    print("[2/4] Anthropic API Key\n")

    # Pruefen ob bereits gesetzt
    bestehender_key = ""
    try:
        result = subprocess.run(
            ["powershell", "-Command",
             "[Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY', 'User')"],
            capture_output=True, text=True,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
        )
        bestehender_key = result.stdout.strip()
    except Exception:
        pass

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
            subprocess.run(
                ["powershell", "-Command",
                 f"[Environment]::SetEnvironmentVariable('ANTHROPIC_API_KEY', '{api_key}', 'User')"],
                capture_output=True,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
            )
            os.environ["ANTHROPIC_API_KEY"] = api_key
            print("  API Key als Windows-Umgebungsvariable gespeichert.\n")
        except Exception as e:
            print(f"  WARNUNG: Konnte Key nicht setzen: {e}")
            print("  Bitte manuell setzen (siehe README.md)\n")

    # ── 3. Lokale Konfiguration ───────────────────────────────────────────
    print("[3/4] Belege-Ordner konfigurieren\n")

    config_local_pfad = os.path.join(agent_dir, "config_local.py")

    if os.path.exists(config_local_pfad):
        # Bestehenden Pfad auslesen
        try:
            with open(config_local_pfad, "r", encoding="utf-8") as f:
                inhalt = f.read()
            # Pfad extrahieren
            import re
            match = re.search(r'ABLAGE_STAMMPFAD\s*=\s*[(\s]*r?"([^"]+)"', inhalt)
            if match:
                bestehender_pfad = match.group(1)
                # Mehrzeilige Pfade zusammensetzen
                matches = re.findall(r'r?"([^"]+)"', inhalt)
                bestehender_pfad = "".join(matches)
            else:
                bestehender_pfad = ""
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
        belege_pfad = frage("Pfad zum Belege-Ordner (z.B. C:\\Users\\Max\\Belege)")

    # config_local.py schreiben
    with open(config_local_pfad, "w", encoding="utf-8") as f:
        f.write('"""\n')
        f.write("Lokale Konfiguration – NICHT ins Git committen!\n")
        f.write('"""\n\n')
        f.write(f'ABLAGE_STAMMPFAD = r"{belege_pfad}"\n')

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
    print("[4/4] Autostart\n")

    if ja_nein("Soll der Agent automatisch bei Windows-Start starten?"):
        try:
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
    print(f"    python {os.path.join(agent_dir, 'beleg_agent.py')}")
    print("  Oder Doppelklick auf start_beleg_agent.vbs")
    print()
    print("  Belege in _Inbox legen – der Rest passiert automatisch.")
    print()

    if ja_nein("Agent jetzt starten?"):
        print("  Starte Beleg-Agent...\n")
        subprocess.Popen(
            [sys.executable, os.path.join(agent_dir, "tray_agent.py")],
            cwd=agent_dir,
        )


if __name__ == "__main__":
    main()
