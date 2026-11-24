#!/bin/bash
# Beleg-Agent Starter
#
# Zwei Modi:
#   (ohne Argument)  Startet den Tray im Hintergrund, kehrt sofort zurueck.
#                    Fuer "Beleg-Agent starten.app" und manuellen Aufruf.
#   --foreground     Haelt den Prozess im Vordergrund. Fuer launchd.
#
# Warum --foreground fuer launchd:
# macOS pinnt einen LaunchAgent per Launch Constraint (LWCR) an den cdhash
# seiner Binary. Homebrew-Python ist nur adhoc signiert, d.h. der cdhash
# aendert sich bei jedem Brew-Rebuild -> macOS killt den Agent mit
# OS_REASON_CODESIGNING und verweigert danach den Start. Laeuft stattdessen
# /bin/bash als launchd-Prozess (Apple Platform Binary, stabiler cdhash), bleibt das
# Pinning gueltig; Python ist nur ein Kindprozess. Stirbt Python, endet bash
# -> KeepAlive(SuccessfulExit=false) startet sauber neu.

set -u
cd "$(dirname "$0")" || exit 1

FOREGROUND=0
[ "${1:-}" = "--foreground" ] && FOREGROUND=1

melde() {
    osascript -e "display notification \"$1\" with title \"Beleg-Agent\"" 2>/dev/null || true
}

fehler() {
    echo "ERROR: $1" >&2
    osascript -e "display notification \"$2\" with title \"Beleg-Agent FEHLER\"" 2>/dev/null || true
    exit 1
}

# ── Laeuft schon einer? (Port 5001 vertraegt keinen zweiten) ──────────────
# Kein pgrep -f: macOS schneidet lange Kommandozeilen ab, ein Start mit
# absolutem Pfad zu tray_agent.py wird dadurch uebersehen. ps + grep sieht
# die volle Zeile, der Port-Check faengt den Rest.
LAUFENDE="$(ps -eo pid=,command= | grep -F 'tray_agent.py' \
    | grep -v grep | awk '{print $1}' | tr '\n' ' ')"
PORT_BELEGT=0
lsof -nP -iTCP:5001 -sTCP:LISTEN >/dev/null 2>&1 && PORT_BELEGT=1

if [ -n "$LAUFENDE" ] || [ "$PORT_BELEGT" -eq 1 ]; then
    echo "Beleg-Agent laeuft bereits (PID ${LAUFENDE:-?})."
    if [ "$FOREGROUND" -eq 0 ]; then
        melde "Laeuft bereits — Web-Interface wird geoeffnet."
        open "http://localhost:5001" 2>/dev/null || true
    fi
    exit 0
fi

# ── Start via launchd (fuer "Beleg-Agent starten.app") ────────────────────
# Nicht direkt Python starten: ueber launchd bekommt der Tray KeepAlive und
# denselben Zustand wie beim Autostart, statt als loser Prozess zu haengen.
if [ "${1:-}" = "--via-launchd" ]; then
    LABEL="gui/$(id -u)/com.meocon.beleg-agent"
    PLIST="$HOME/Library/LaunchAgents/com.meocon.beleg-agent.plist"
    UEBER_LAUNCHD=0
    if launchctl print "$LABEL" >/dev/null 2>&1; then
        launchctl kickstart "$LABEL" >/dev/null 2>&1 && UEBER_LAUNCHD=1
    elif [ -f "$PLIST" ]; then
        launchctl bootstrap "gui/$(id -u)" "$PLIST" >/dev/null 2>&1 && UEBER_LAUNCHD=1
    fi

    if [ "$UEBER_LAUNCHD" -eq 1 ]; then
        # Auf den Webserver warten, dann Browser oeffnen
        for _ in $(seq 1 20); do
            if lsof -nP -iTCP:5001 -sTCP:LISTEN >/dev/null 2>&1; then
                melde "Beleg-Agent gestartet."
                open "http://localhost:5001" 2>/dev/null || true
                exit 0
            fi
            sleep 1
        done
        fehler "Start dauerte zu lange — siehe ~/Library/Logs/beleg-agent.launchd.err.log" \
               "Start fehlgeschlagen — Log pruefen"
    fi

    # Kein LaunchAgent vorhanden (Autostart im Setup abgelehnt) oder launchctl
    # verweigert: dann weiter unten direkt starten. Ohne KeepAlive, aber eine
    # App, die nur eine Fehlermeldung zeigt, hilft niemandem. tray_agent.py
    # oeffnet den Browser selbst, der Rest des Skripts erledigt den Start.
    echo "Kein LaunchAgent aktiv — starte direkt (ohne Neustart-Schutz)." >&2
fi

# ── API-Key aus dem macOS-Keychain ────────────────────────────────────────
if command -v direnv >/dev/null 2>&1 && [ -f .envrc ]; then
    eval "$(direnv export bash 2>/dev/null)" || true
fi

if [ -z "${ANTHROPIC_API_KEY:-}" ]; then
    ANTHROPIC_API_KEY="$(security find-generic-password -s beleg_agent \
        -a ANTHROPIC_API_KEY -w 2>/dev/null || true)"
    export ANTHROPIC_API_KEY
fi

if [ -z "${ANTHROPIC_API_KEY:-}" ]; then
    echo "Key im macOS-Keychain hinterlegen und neu starten:" >&2
    echo "  security add-generic-password -s beleg_agent -a ANTHROPIC_API_KEY -U -w \"sk-ant-...\"" >&2
    fehler "ANTHROPIC_API_KEY nicht gesetzt (Keychain leer oder gesperrt)." \
           "ANTHROPIC_API_KEY fehlt — im Keychain hinterlegen und neu starten"
fi

[ -x .venv/bin/python3 ] || fehler \
    ".venv/bin/python3 fehlt oder ist nicht ausfuehrbar." \
    "Python-Umgebung fehlt — .venv neu anlegen"

# ── Start ─────────────────────────────────────────────────────────────────
if [ "$FOREGROUND" -eq 1 ]; then
    .venv/bin/python3 tray_agent.py &
    KIND=$!
    trap 'kill -TERM "$KIND" 2>/dev/null' TERM INT
    wait "$KIND"
    exit $?
else
    # nohup + Umleitung: ohne das nimmt die aufrufende Shell (Terminal oder
    # das "do shell script" der Start-App) den Tray beim Beenden mit.
    # disown allein genuegt nicht — SIGHUP erreicht die Prozessgruppe trotzdem.
    nohup .venv/bin/python3 tray_agent.py >>"$PWD/beleg-agent.start.log" 2>&1 &
    disown
    melde "Beleg-Agent gestartet."
fi
