#!/bin/bash
cd "$(dirname "$0")"

if command -v direnv >/dev/null 2>&1 && [ -f .envrc ]; then
    eval "$(direnv export bash 2>/dev/null)" || true
fi

if [ -z "${ANTHROPIC_API_KEY:-}" ]; then
    echo "ERROR: ANTHROPIC_API_KEY nicht gesetzt (Keychain leer oder .envrc nicht erlaubt)." >&2
    echo "Key im macOS-Keychain hinterlegen und Skript neu starten:" >&2
    echo "  security add-generic-password -s beleg_agent -a ANTHROPIC_API_KEY -U -w \"sk-ant-...\"" >&2
    echo "  direnv allow   # falls .envrc geaendert wurde" >&2
    echo "  $(basename "$0")" >&2
    osascript -e 'display notification "ANTHROPIC_API_KEY fehlt — im Keychain hinterlegen (security add-generic-password) und Tray neu starten" with title "Beleg-Agent FEHLER"' 2>/dev/null || true
    exit 1
fi

.venv/bin/python3 tray_agent.py &
disown
