"""
Beleg-Agent Konfiguration
Alle Pfade und Einstellungen zentral verwaltet.

Persoenliche Pfade werden aus config_local.py geladen.
Kopiere config_local.example.py → config_local.py und passe die Pfade an.
"""

import os
import sys
import time
from contextlib import contextmanager

try:
    from config_local import ABLAGE_STAMMPFAD
except ImportError:
    print("FEHLER: config_local.py nicht gefunden!")
    print("Kopiere config_local.example.py → config_local.py und passe die Pfade an.")
    sys.exit(1)

try:
    from config_local import BEKANNTE_KARTEN
except ImportError:
    BEKANNTE_KARTEN = {}  # Keine Karten konfiguriert → Vision erkennt ZA ohne Hinweis

try:
    from config_local import BANK_PROFIL
except ImportError:
    BANK_PROFIL = "ubs"  # Standard-Bankprofil

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
# Reihenfolge: Kerndaten → Finanzen → Abgleich → Bemerkungen → Metadaten
EXCEL_SPALTEN = [
    "Datum_Rechnung",       #  1 - Kerndaten
    "Rechnungssteller",     #  2
    "Typ",                  #  3 - Rechnung / Gutschrift / Dauerauftrag
    "Betrag",               #  4 - Finanzen
    "Währung",              #  5
    "Zahlungsart",          #  6 - KK CHF / KK EUR / Überweisung / eBill / leer
    "PayPal",               #  7 - Ja / Nein
    "Währung_Belastet",     #  8 - KK-Abrechnungswährung bei Fremdwährung
    "Betrag_Belastet",      #  9 - Tatsächlich belasteter Betrag in KK-Währung
    "Abgeglichen",          # 10 - Nein / Ja
    "Bemerkungen",          # 11 - Trinkgeld, Hinweise etc.
    "Originaldateiname",    # 12 - Metadaten
    "Ablagepfad",           # 13
    "Confidence_Score",     # 14
    "Verarbeitungsdatum",   # 15
]

# ── Spalten-Nummern (für Code-Referenzen) ─────────────────────────────────
COL_DATUM = 1
COL_RECHNUNGSSTELLER = 2
COL_TYP = 3
COL_BETRAG = 4
COL_WAEHRUNG = 5
COL_ZAHLUNGSART = 6
COL_PAYPAL = 7
COL_WAEHRUNG_BELASTET = 8
COL_BETRAG_BELASTET = 9
COL_ABGEGLICHEN = 10
COL_BEMERKUNGEN = 11
COL_ORIGINALDATEINAME = 12
COL_ABLAGEPFAD = 13
COL_CONFIDENCE = 14
COL_VERARBEITUNGSDATUM = 15

# ── Bank-/KK-Abgleich (für späteres Reconciliation-Feature) ──────────────
# Pfade zu CSV-Exporten der Bank-/KK-Auszüge
ABGLEICH_PFAD = os.path.join(ABLAGE_STAMMPFAD, "_Abgleich")

# ── File-Lock fuer Excel-Zugriff ─────────────────────────────────────────
EXCEL_LOCK = EXCEL_PROTOKOLL + ".lock"
_LOCK_TIMEOUT = 30      # Max. Sekunden warten
_LOCK_STALE = 120       # Lock gilt als stale nach N Sekunden


@contextmanager
def excel_lock():
    """Context-Manager: verhindert gleichzeitigen Excel-Zugriff durch mehrere Prozesse."""
    start = time.time()
    while True:
        try:
            fd = os.open(EXCEL_LOCK, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
            break
        except FileExistsError:
            # Stale-Lock aufloesen (z.B. nach Absturz)
            try:
                if time.time() - os.path.getmtime(EXCEL_LOCK) > _LOCK_STALE:
                    os.remove(EXCEL_LOCK)
                    continue
            except OSError:
                pass
            if time.time() - start > _LOCK_TIMEOUT:
                raise TimeoutError(
                    f"Excel-Lock nicht verfuegbar nach {_LOCK_TIMEOUT}s. "
                    f"Lock-Datei: {EXCEL_LOCK}"
                )
            time.sleep(0.3)
    try:
        yield
    finally:
        try:
            os.remove(EXCEL_LOCK)
        except OSError:
            pass

ABGLEICH_QUELLEN = {
    "KK CHF": os.path.join(ABGLEICH_PFAD, "KK_CHF.csv"),
    "KK EUR": os.path.join(ABGLEICH_PFAD, "KK_EUR.csv"),
    "Bank":   "",   # z.B. os.path.join(ABGLEICH_PFAD, "Bank.csv")
    "PayPal": "",   # z.B. os.path.join(ABGLEICH_PFAD, "PayPal.csv")
}

# CSV-Datumsformat der Bankauszüge (für späteres Parsing)
ABGLEICH_DATUMSFORMAT = "%d.%m.%Y"
