"""
Beleg-Agent Konfiguration
Alle Pfade und Einstellungen zentral verwaltet.

Persoenliche Pfade werden aus config_local.py geladen.
Kopiere config_local.example.py → config_local.py und passe die Pfade an.
"""

import os
import sys

try:
    from config_local import ABLAGE_STAMMPFAD
except ImportError:
    print("FEHLER: config_local.py nicht gefunden!")
    print("Kopiere config_local.example.py → config_local.py und passe die Pfade an.")
    sys.exit(1)

# ── Pfade (abgeleitet aus ABLAGE_STAMMPFAD) ───────────────────────────────
INBOX_PFAD = os.path.join(ABLAGE_STAMMPFAD, "_Inbox")

EXCEL_PROTOKOLL = os.path.join(ABLAGE_STAMMPFAD, "Belege_Protokoll.xlsx")

LOG_DATEI = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "beleg-agent.log"
)

# ── Claude API ─────────────────────────────────────────────────────────────
ANTHROPIC_MODEL = "claude-sonnet-4-20250514"
# Alternativ: "claude-opus-4-5-20250514"

# ── Schwellenwerte ─────────────────────────────────────────────────────────
CONFIDENCE_AUTO = 0.85        # Ab hier: automatische Ablage
CONFIDENCE_RÜCKFRAGE = 0.60   # Ab hier: Terminal-Rückfrage
# Darunter: [PRÜFEN]-Prefix

RÜCKFRAGE_TIMEOUT_SEKUNDEN = 60
WARTE_NACH_ERKENNUNG_SEKUNDEN = 2

# ── Dateifilter ────────────────────────────────────────────────────────────
ERLAUBTE_ENDUNGEN = {".pdf", ".jpg", ".jpeg", ".png"}

# ── Monatsordner-Mapping ──────────────────────────────────────────────────
MONATE = {
    1: "01_Januar",
    2: "02_Februar",
    3: "03_März",
    4: "04_April",
    5: "05_Mai",
    6: "06_Juni",
    7: "07_Juli",
    8: "08_August",
    9: "09_September",
    10: "10_Oktober",
    11: "11_November",
    12: "12_Dezember",
}

# ── Excel-Spalten ──────────────────────────────────────────────────────────
EXCEL_SPALTEN = [
    "Datum_Rechnung",
    "Rechnungssteller",
    "Betrag",
    "Währung",
    "Typ",               # Rechnung / Gutschrift
    "Zahlungsart",       # KK CHF / KK EUR / Überweisung / leer (= unbekannt)
    "PayPal",            # Ja / Nein
    "Originaldateiname",
    "Ablagepfad",
    "Abgeglichen",       # Nein / Ja – wird erst nach Bank-/KK-Abgleich auf Ja gesetzt
    "Confidence_Score",
    "Verarbeitungsdatum",
    "Bemerkungen",
]

# ── Bank-/KK-Abgleich (für späteres Reconciliation-Feature) ──────────────
# Pfade zu CSV-Exporten der Bank-/KK-Auszüge
ABGLEICH_PFAD = os.path.join(ABLAGE_STAMMPFAD, "_Abgleich")

ABGLEICH_QUELLEN = {
    "KK CHF": os.path.join(ABGLEICH_PFAD, "KK_CHF.csv"),
    "KK EUR": os.path.join(ABGLEICH_PFAD, "KK_EUR.csv"),
    "Bank":   "",   # z.B. os.path.join(ABGLEICH_PFAD, "Bank.csv")
    "PayPal": "",   # z.B. os.path.join(ABGLEICH_PFAD, "PayPal.csv")
}

# CSV-Datumsformat der Bankauszüge (für späteres Parsing)
ABGLEICH_DATUMSFORMAT = "%d.%m.%Y"
