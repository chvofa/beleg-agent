#!/usr/bin/env python3
"""
Beleg-Agent – Provisorische Kontierung

Lookup Lieferant/Buchungstext → Konto-Vorschlag auf Basis der
Kontierung_Mapping.json (gebaut von import_kontierung.py aus den
Kontoblättern des Treuhänders).

Der Vorschlag ist rein informativ: er beeinflusst weder Confidence-Routing
noch Ablage. Fehlt die Mapping-Datei oder ist sie kaputt, degradiert alles
stumm (leere Vorschläge).

CLI:
    python kontierung.py --backfill   # füllt leere Konto-Spalten rückwirkend
"""

import json
import os
import re
import sys
from collections import Counter

import config


# ═══════════════════════════════════════════════════════════════════════════
#  Normalisierung
# ═══════════════════════════════════════════════════════════════════════════

# Rausch-Präfixe der Treuhand-Buchungstexte: "750 8.1 I ", "V75 8.1 I ",
# "230 2.6 I ", "Deb. ", "Kred. " etc.
_PREFIX_RE = re.compile(r"^(v?\d{2,4}\s+\d+(\.\d+)?\s+i\s+|(deb|kred)\.\s+)", re.IGNORECASE)
# Datums-/Referenz-Suffixe: ", 012024", " 12.2023", lange Ziffernfolgen
_SUFFIX_RE = re.compile(r"[,;]?\s*(\d{2}\.\d{4}|\d{2}\.\d{2}\.\d{4}|\d{6,})\s*$")
# Klammern mit Zahlen sind Rauschen ("(EUR 4'269.38)"), reine Wort-Klammern
# sind Semantik ("(Büromiete)", "(Lohn)") und bleiben als Tokens erhalten
_ZAHLEN_KLAMMER_RE = re.compile(r"\s*\((?=[^)]*\d)[^)]*\)")


def normalisiere_text(text: str) -> str:
    """Lowercase, Präfixe/Suffixe strippen, Whitespace kollabieren."""
    t = (text or "").strip().lower()
    vorher = None
    while vorher != t:
        vorher = t
        t = _PREFIX_RE.sub("", t).strip()
        t = _SUFFIX_RE.sub("", t).strip()
        t = _ZAHLEN_KLAMMER_RE.sub("", t).strip()
    t = t.replace("(", " ").replace(")", " ")
    return re.sub(r"\s+", " ", t).strip()


# ═══════════════════════════════════════════════════════════════════════════
#  Mapping laden (mit Cache)
# ═══════════════════════════════════════════════════════════════════════════

_cache: dict = {}
_cache_mtime: float = -1.0


def lade_mapping() -> dict:
    """Lädt Kontierung_Mapping.json, cached per mtime. Fehler → {}."""
    global _cache, _cache_mtime
    try:
        mtime = os.path.getmtime(config.KONTIERUNG_MAPPING)
    except OSError:
        return {}
    if mtime == _cache_mtime and _cache:
        return _cache
    try:
        with open(config.KONTIERUNG_MAPPING, encoding="utf-8") as f:
            _cache = json.load(f)
        _cache_mtime = mtime
    except (OSError, json.JSONDecodeError):
        return {}
    return _cache


# ═══════════════════════════════════════════════════════════════════════════
#  Lookup
# ═══════════════════════════════════════════════════════════════════════════

# Tokens ohne Unterscheidungskraft: Rechtsformen, Geografie, Generisches.
# Ohne diese Liste matcht "Sky Switzerland" via "switzerland" auf
# "Corden Pharma Switzerland LLC" (Ertragskonto!) oder jede "Foo GmbH"
# auf den häufigsten GmbH-Eintrag.
_STOPWORTE = {
    # Rechtsformen
    "gmbh", "sagl", "sarl", "srl", "llc", "ltd", "inc", "corp",
    # Geografie (Länder + Städte aus Bank-Adressen und Restaurant-/Hotel-Keys)
    "schweiz", "switzerland", "suisse", "svizzera", "swiss",
    "deutschland", "germany", "österreich", "austria", "international",
    "bern", "zürich", "zuerich", "zurich", "basel", "gallen", "luzern",
    "lucerne", "thun", "visp", "brig", "naters", "freiburg", "fribourg",
    "zermatt", "kandersteg", "schlieren", "crans", "montana", "amsterdam",
    # Generische Firmenwörter
    "holding", "group", "gruppe", "services", "service",
    "management", "consulting", "center", "centre",
}


def _tokens(text: str) -> set[str]:
    """Wort-Tokens: an JEDEM Nicht-Wortzeichen trennen (auch ;,/|-), damit
    Bank-Texte wie 'HSG Alumni;Dufourstrasse 50' sauber zerfallen.
    Reine Zahlen (Jahre, PLZ, Referenzen) sind keine Match-Beweise."""
    return {t for t in re.split(r"\W+", text)
            if len(t) > 3 and not t.isdigit() and t not in _STOPWORTE}


def lookup_konto(text: str) -> dict | None:
    """Sucht das historische Konto für einen Lieferanten/Buchungstext.

    Konservativ (analog lookup_zahlungsart): 1) exakter Key, 2) Teilstring
    beidseitig (min. 4 Zeichen), 3) Token-Overlap. Bei Mehrfachtreffern
    gewinnt der Eintrag mit der höchsten Anzahl Buchungen.
    """
    daten = lade_mapping()
    if not daten:
        return None
    norm = normalisiere_text(text)
    if len(norm) < 3:
        return None

    # Manuelle Overrides wie normale Einträge behandeln — sie gewinnen bei
    # exaktem Key und nehmen am Fuzzy-Matching teil (Bank-Texte wie
    # "Parkingpay Schlieren CHE" sollen die Regel "parkingpay" treffen)
    mappings = dict(daten.get("mappings", {}))
    for k, v in daten.get("manuell", {}).items():
        mappings[k] = {"konto": str(v), "anzahl": 3, "anteil": 1.0}

    if norm in mappings:
        return dict(mappings[norm])

    # Fuzzy in zwei Stufen: Teilstring-Treffer (präziser) vor Token-Overlap.
    # Innerhalb einer Stufe werden die Kandidaten PRO KONTO aggregiert —
    # sonst gewinnt bei lauter 1x-Treffern zufällig der alphabetisch erste
    # Key ("kanton bern, strassenverkehrsamt" schlug so die drei
    # "kanton zug, steuern"-Buchungen).
    sub_kandidaten, tok_kandidaten = [], []
    tokens = _tokens(norm)
    for key, eintrag in mappings.items():
        if len(norm) >= 4 and len(key) >= 4 and (norm in key or key in norm):
            sub_kandidaten.append(eintrag)
        elif tokens and tokens & _tokens(key):
            tok_kandidaten.append(eintrag)
    kandidaten = sub_kandidaten or tok_kandidaten
    if not kandidaten:
        return None
    zaehler = Counter()
    for e in kandidaten:
        zaehler[e["konto"]] += e["anzahl"]
    konto, top = zaehler.most_common(1)[0]
    total = sum(zaehler.values())
    return {"konto": konto, "anzahl": total, "anteil": round(top / total, 2)}


def format_vorschlag(hit: dict | None) -> str:
    """Formatiert einen Lookup-Treffer: '4530 Telefon / Internet' (+ ' ?' wenn unsicher)."""
    if not hit:
        return ""
    konto = str(hit.get("konto", ""))
    name = lade_mapping().get("kontenplan", {}).get(konto, "")
    unsicher = hit.get("anteil", 1.0) < 0.6 or hit.get("anzahl", 0) == 1
    return f"{konto} {name}".strip() + (" ?" if unsicher else "")


def vorschlag_fuer(text: str) -> str:
    """Bequemer Einzeiler: Text rein, formatierter Vorschlag (oder '') raus."""
    return format_vorschlag(lookup_konto(text))


# ═══════════════════════════════════════════════════════════════════════════
#  Backfill (CLI)
# ═══════════════════════════════════════════════════════════════════════════

def backfill() -> None:
    """Füllt leere Konto_Vorschlag-Zellen im Protokoll und bei Offenen Posten."""
    import offene_posten

    if not lade_mapping():
        print(f"FEHLER: Mapping fehlt/unlesbar: {config.KONTIERUNG_MAPPING}")
        print("Zuerst ausführen: python import_kontierung.py")
        sys.exit(1)
    if not os.path.exists(config.EXCEL_PROTOKOLL):
        print(f"FEHLER: Protokoll nicht gefunden: {config.EXCEL_PROTOKOLL}")
        sys.exit(1)

    with config.excel_lock():
        wb = config.load_protokoll()
        ws = wb.active

        # Header sicherstellen (Agent-Neustart macht das sonst auch)
        if (ws.max_column or 0) < len(config.EXCEL_SPALTEN):
            for col_idx in range((ws.max_column or 0) + 1, len(config.EXCEL_SPALTEN) + 1):
                ws.cell(row=1, column=col_idx).value = config.EXCEL_SPALTEN[col_idx - 1]

        gefuellt = ohne_treffer = 0
        for row_idx in range(2, ws.max_row + 1):
            rs = str(ws.cell(row=row_idx, column=config.COL_RECHNUNGSSTELLER).value or "")
            if not rs.strip():
                continue
            zelle = ws.cell(row=row_idx, column=config.COL_KONTO_VORSCHLAG)
            if str(zelle.value or "").strip():
                continue
            vorschlag = vorschlag_fuer(rs)
            if vorschlag:
                zelle.value = vorschlag
                gefuellt += 1
            else:
                ohne_treffer += 1
        print(f"Protokoll:     {gefuellt} gefüllt, {ohne_treffer} ohne Treffer")

        ws_op = offene_posten.ensure_sheet(wb)
        gefuellt = ohne_treffer = 0
        for row_idx in range(2, ws_op.max_row + 1):
            text = str(ws_op.cell(row=row_idx, column=offene_posten.COL_TEXT).value or "")
            if not text.strip():
                continue
            zelle = ws_op.cell(row=row_idx, column=offene_posten.COL_KONTO)
            if str(zelle.value or "").strip():
                continue
            vorschlag = vorschlag_fuer(text)
            if vorschlag:
                zelle.value = vorschlag
                gefuellt += 1
            else:
                ohne_treffer += 1
        print(f"Offene Posten: {gefuellt} gefüllt, {ohne_treffer} ohne Treffer")

        wb.save(config.EXCEL_PROTOKOLL)
        wb.close()


if __name__ == "__main__":
    if "--backfill" in sys.argv:
        backfill()
    else:
        # Smoke-Test: python kontierung.py "750 8.1 I Sunrise GmbH"
        for arg in sys.argv[1:]:
            print(f"{arg!r} → {vorschlag_fuer(arg)!r}")
        if len(sys.argv) < 2:
            print("Aufruf: python kontierung.py --backfill | <text> [<text> ...]")
