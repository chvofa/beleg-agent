"""
Beleg-Agent Konfiguration
Alle Pfade und Einstellungen zentral verwaltet.

Persoenliche Pfade werden aus config_local.py geladen.
Kopiere config_local.example.py → config_local.py und passe die Pfade an.
"""

import os
import sys
import time
import zipfile
from contextlib import contextmanager

import openpyxl

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

try:
    from config_local import MIN_JAHR_ABGLEICH
except ImportError:
    MIN_JAHR_ABGLEICH = 2026  # Abgleich ignoriert Transaktionen vor diesem Jahr

# ── Pfade (abgeleitet aus ABLAGE_STAMMPFAD) ───────────────────────────────
ABLAGE_STAMMPFAD = os.path.expanduser(ABLAGE_STAMMPFAD)
INBOX_PFAD = os.path.join(ABLAGE_STAMMPFAD, "_Inbox")

EXCEL_PROTOKOLL = os.path.join(ABLAGE_STAMMPFAD, "Belege_Protokoll.xlsx")

# Lieferant→Konto-Mapping aus den Treuhänder-Kontoblättern (import_kontierung.py)
KONTIERUNG_MAPPING = os.path.join(ABLAGE_STAMMPFAD, "Kontierung_Mapping.json")

LOG_DATEI = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "beleg-agent.log"
)

# ── Claude API ─────────────────────────────────────────────────────────────
# claude-sonnet-4-20250514 (Sonnet 4) wird abgekuendigt und liefert seit
# 2026-06-22 einen 404 not_found_error -> Drop-in-Ersatz: Sonnet 4.6.
ANTHROPIC_MODEL = "claude-sonnet-4-6"
# Alternativ (teurer, mehr Reasoning): "claude-opus-4-8"

# ── Schwellenwerte ─────────────────────────────────────────────────────────
CONFIDENCE_AUTO = 0.85        # Ab hier: automatische Ablage
CONFIDENCE_RÜCKFRAGE = 0.60   # Ab hier: Terminal-Rückfrage
# Darunter: [PRÜFEN]-Prefix

RÜCKFRAGE_TIMEOUT_SEKUNDEN = 60

# ── OneDrive-Sync: warten bis der Beleg wirklich vollstaendig lokal liegt ──
# Die _Inbox liegt in einer OneDrive-Bibliothek. Watchdog meldet eine neue
# Datei, sobald der Verzeichniseintrag da ist — der Inhalt kommt u.U. erst
# Sekunden spaeter. Ein fixes sleep() reichte nicht: grosse Scan-PDFs kamen
# abgeschnitten bei pdfium an ("PDFium: Data format error"). Statt fix zu
# warten wird gepollt, bis die Dateigroesse konstant bleibt.
DATEI_STABIL_INTERVALL = 0.5   # Sekunden zwischen zwei Groessen-Messungen
DATEI_STABIL_CHECKS = 3        # so viele identische Messungen = fertig
DATEI_STABIL_TIMEOUT = 120     # Abbruch nach N Sekunden (grosse Scans/langsamer Sync)

# Files-On-Demand-Platzhalter melden bereits die finale Groesse, haben den
# Inhalt aber nur auf dem Server. Deshalb zusaetzlich: Datei komplett nach
# Python lesen (erzwingt die Hydration) und bei Lesefehlern erneut versuchen.
PDF_LADE_VERSUCHE = 3          # Gesamtversuche fuer das Einlesen/Rendern
PDF_LADE_DELAY = 2.0           # Sekunden Pause zwischen den Versuchen

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
    "Valutadatum",          # 16 - Tatsaechliches Buchungsdatum der Bank/KK-Transaktion (aus Abgleich)
    "Referenz_Nr",          # 17 - Transaktions-/Auftragsnummer fuer eindeutigen Abgleich
    "Konto_Vorschlag",      # 18 - Provisorische Kontierung aus Kontoblatt-Historie, z.B. "4530 Telefon/Internet"
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
COL_VALUTADATUM = 16
COL_REFERENZ = 17
COL_KONTO_VORSCHLAG = 18

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


def load_protokoll(read_only: bool = False, _versuche: int = 5, _delay: float = 0.5):
    """Laedt das Excel-Protokoll robust gegen transiente OneDrive-Sync-Fehler.

    Waehrend OneDrive Belege_Protokoll.xlsx synchronisiert, schlaegt
    openpyxl.load_workbook sporadisch fehl mit zipfile.BadZipFile
    ("File is not a zip file") oder OSError (z.B. [Errno 11]
    "Resource deadlock avoided"). Diese Fehler sind transient — ein kurzer
    Retry sitzt das Sync-Fenster aus. Nach _versuche erfolglosen Versuchen
    wird der letzte Fehler weitergereicht (kein stilles None).

    Aufrufer laden weiterhin innerhalb von excel_lock().
    """
    letzter_fehler = None
    for versuch in range(_versuche):
        try:
            return openpyxl.load_workbook(EXCEL_PROTOKOLL, read_only=read_only)
        except (zipfile.BadZipFile, OSError) as e:
            letzter_fehler = e
            if versuch < _versuche - 1:
                time.sleep(_delay)
    raise letzter_fehler

ABGLEICH_QUELLEN = {
    "KK CHF": os.path.join(ABGLEICH_PFAD, "KK_CHF.csv"),
    "KK EUR": os.path.join(ABGLEICH_PFAD, "KK_EUR.csv"),
    "Bank":   "",   # z.B. os.path.join(ABGLEICH_PFAD, "Bank.csv")
    "PayPal": "",   # z.B. os.path.join(ABGLEICH_PFAD, "PayPal.csv")
}

# CSV-Datumsformat der Bankauszüge (für späteres Parsing)
ABGLEICH_DATUMSFORMAT = "%d.%m.%Y"
