#!/usr/bin/env python3
"""
Beleg-Agent – Kontierungs-Import

Baut Kontierung_Mapping.json aus den Kontoblättern des Treuhänders:
pro Buchung auf den Bank-/KK-Konten steht der Text (Gegenpartei) und das
Gegenkonto — daraus entsteht ein Lieferant→Konto-Mapping mit Häufigkeiten.

Quellen (Default, per CLI-Arg übersteuerbar):
  1. Kontoblätter 2024 als XLSX (definitiv, ein Sheet pro Konto)
  2. prov. Kontoblätter 2025 als PDF (Entwurf; später durch definitive ersetzen)

Wiederholbar: baut kontenplan + mappings komplett neu, erhält aber die
"manuell"-Sektion aus einer bestehenden Mapping-Datei.

Aufruf:
    python import_kontierung.py [xlsx_pfad ...] [pdf_pfad ...]
"""

import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime

import openpyxl

import config
from kontierung import normalisiere_text

ABSCHLUESSE = (
    "/Users/fabiooro/Library/CloudStorage/"
    "OneDrive-FreigegebeneBibliotheken–MeoconGmbH/"
    "Meocon Admin - Documents/02 - Meocon/02 - Financials/03-Abschlüsse"
)
DEFAULT_QUELLEN = [
    os.path.join(ABSCHLUESSE, "2024/Kontoblätter/Kontoblätter 2024_Meocon GmbH.xlsx"),
    os.path.join(ABSCHLUESSE, "2025/ENTWURF/prov. Kontoblätter.pdf"),
]

# Konten deren Buchungstexte die Gegenpartei benennen:
# Bank/Kassa/KK-Durchlauf plus die Vorsteuer-Konten 1065/1066/1068 —
# letztere listen die einzelnen KK-Händler mit echtem Gegenkonto
QUELL_KONTEN = {"1000", "1010", "1020", "1030", "1098", "1099",
                "1065", "1066", "1068"}
# Gegenkonten die keine echte Kontierung sind (interne Transfers, MWSt-Umbuchungen)
SKIP_GEGENKONTEN = QUELL_KONTEN | {"1999", "1025", "1027", "2005"}

# Buchungstexte ohne Kontierungs-Aussagekraft
SKIP_TEXTE = re.compile(
    r"^(saldovortrag|übertrag|saldo\b|zahlungsauftrag|sammelbuchung"
    r"|kauf (eur|chf)|privatant|ausgleich)",  # "privatant" deckt auch Tippfehler "Privatantiel" ab
    re.IGNORECASE,
)

DATUM_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")


def _sammle(eintraege: list, text: str, gegenkonto) -> None:
    """Prüft Skip-Regeln und hängt (normalisierter Text, Konto) an eintraege."""
    text = (text or "").strip()
    if not text or SKIP_TEXTE.match(text):
        return
    konto = str(gegenkonto or "").strip()
    if not re.fullmatch(r"\d{4}", konto) or konto in SKIP_GEGENKONTEN:
        return
    norm = normalisiere_text(text)
    if len(norm) < 3:
        return
    eintraege.append((norm, konto))


# ═══════════════════════════════════════════════════════════════════════════
#  XLSX-Parser (Kontoblätter-Export, ein Sheet pro Konto)
# ═══════════════════════════════════════════════════════════════════════════

def parse_xlsx(pfad: str, kontenplan: dict, eintraege: list) -> int:
    """Liest Kontoblätter-XLSX. Sheet-Namen wie '_4530_Telefon__Internet'."""
    wb = openpyxl.load_workbook(pfad, read_only=True)
    n_vorher = len(eintraege)

    for sheet_name in wb.sheetnames:
        m = re.match(r"^_(\d{4})_(.+)$", sheet_name)
        if not m:
            continue
        konto_nr, konto_name = m.group(1), m.group(2)
        # Sheet-Namen sind verstümmelt (_ statt Sonderzeichen) — nur als
        # Fallback in den Kontenplan, PDF-Titel überschreiben sie später.
        kontenplan.setdefault(konto_nr, re.sub(r"_+", " ", konto_name).strip())

        if konto_nr not in QUELL_KONTEN:
            continue

        ws = wb[sheet_name]
        for row in ws.iter_rows(values_only=True):
            if len(row) < 4:
                continue
            datum, text, gegenkonto = row[0], row[1], row[3]
            if not DATUM_RE.match(str(datum or "").strip()):
                continue
            _sammle(eintraege, str(text or ""), gegenkonto)

    wb.close()
    return len(eintraege) - n_vorher


# ═══════════════════════════════════════════════════════════════════════════
#  PDF-Parser (gleiche Kontoblätter als PDF, via pypdfium2-Textextraktion)
# ═══════════════════════════════════════════════════════════════════════════

KONTO_TITEL_RE = re.compile(r"^Konto (\d{4}) (.+?)\s*$")
ZEILEN_START_RE = re.compile(r"^(\d{2}\.\d{2}\.\d{4})\s+(.*)$")
# Buchungszeile komplett wenn "<Text> <Gegenkonto> F|CF|S|SF ..." erreichbar ist
RECORD_ENDE_RE = re.compile(r"^(.+?)\s+(\d{4})\s?(F|CF|S|SF)(\s|$)")


def parse_pdf(pfad: str, kontenplan: dict, eintraege: list) -> int:
    """Liest Kontoblätter-PDF zeilenweise mit State-Machine (umbruchsicher)."""
    import pypdfium2 as pdfium

    n_vorher = len(eintraege)
    doc = pdfium.PdfDocument(pfad)

    aktives_konto = ""
    buffer = ""  # gesammelter Text der aktuellen Buchungszeile (ohne Datum)

    def buffer_abschliessen():
        nonlocal buffer
        if aktives_konto in QUELL_KONTEN and buffer:
            m = RECORD_ENDE_RE.match(buffer)
            if m:
                _sammle(eintraege, m.group(1), m.group(2))
        buffer = ""

    for page in doc:
        text = page.get_textpage().get_text_bounded()
        for zeile in text.splitlines():
            zeile = zeile.strip()
            if not zeile:
                continue

            m_titel = KONTO_TITEL_RE.match(zeile)
            if m_titel:
                buffer_abschliessen()
                aktives_konto = m_titel.group(1)
                # PDF-Titel sind sauber ("Telefon / Internet") — gewinnen
                kontenplan[m_titel.group(1)] = m_titel.group(2)
                continue

            m_start = ZEILEN_START_RE.match(zeile)
            if m_start:
                buffer_abschliessen()
                buffer = m_start.group(2)
            elif buffer:
                # Fortsetzung einer umbrochenen Buchungszeile
                buffer += " " + zeile

            if buffer and RECORD_ENDE_RE.match(buffer):
                buffer_abschliessen()

            # Saldo-/Übertrag-Zeilen beenden offene Buffer ohne Treffer
            if re.match(r"^(Saldo|Übertrag|Saldovortrag)", zeile):
                buffer = ""

    buffer_abschliessen()
    doc.close()
    return len(eintraege) - n_vorher


# ═══════════════════════════════════════════════════════════════════════════
#  Aggregation + Persistenz
# ═══════════════════════════════════════════════════════════════════════════

def baue_mapping(eintraege: list) -> dict:
    """Aggregiert (text, konto)-Paare zu Mapping-Einträgen mit Häufigkeiten."""
    zaehler: dict[str, Counter] = defaultdict(Counter)
    for norm, konto in eintraege:
        zaehler[norm][konto] += 1

    mappings = {}
    for norm, counter in sorted(zaehler.items()):
        konto, anzahl_top = counter.most_common(1)[0]
        total = sum(counter.values())
        mappings[norm] = {
            "konto": konto,
            "anzahl": total,
            "anteil": round(anzahl_top / total, 2),
            "alternativen": dict(counter.most_common()),
        }
    return mappings


def main(argv: list[str]) -> int:
    quellen = argv or DEFAULT_QUELLEN

    kontenplan: dict[str, str] = {}
    eintraege: list[tuple[str, str]] = []

    for pfad in quellen:
        if not os.path.exists(pfad):
            print(f"FEHLER: Quelle nicht gefunden: {pfad}")
            return 1
        if pfad.lower().endswith(".xlsx"):
            n = parse_xlsx(pfad, kontenplan, eintraege)
        elif pfad.lower().endswith(".pdf"):
            n = parse_pdf(pfad, kontenplan, eintraege)
        else:
            print(f"FEHLER: Unbekanntes Format (erwarte .xlsx/.pdf): {pfad}")
            return 1
        print(f"  {n:4d} Buchungen aus {os.path.basename(pfad)}")

    mappings = baue_mapping(eintraege)

    # Manuell gepflegte Einträge aus bestehender Datei erhalten
    manuell = {}
    if os.path.exists(config.KONTIERUNG_MAPPING):
        try:
            with open(config.KONTIERUNG_MAPPING, encoding="utf-8") as f:
                manuell = json.load(f).get("manuell", {})
        except (OSError, json.JSONDecodeError):
            pass

    daten = {
        "version": 1,
        "erstellt": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "quellen": [os.path.basename(p) for p in quellen],
        "kontenplan": dict(sorted(kontenplan.items())),
        "mappings": mappings,
        "manuell": manuell,
    }
    with open(config.KONTIERUNG_MAPPING, "w", encoding="utf-8") as f:
        json.dump(daten, f, ensure_ascii=False, indent=2)

    ambig = {k: v for k, v in mappings.items() if v["anteil"] < 0.6}
    print(f"\nKontenplan: {len(kontenplan)} Konten")
    print(f"Mappings:   {len(mappings)} Lieferanten/Texte "
          f"({sum(v['anzahl'] for v in mappings.values())} Buchungen)")
    if ambig:
        print(f"\nMehrdeutig (Anteil < 60%, bekommen '?'-Kennzeichnung):")
        for k, v in ambig.items():
            print(f"  {k}: {v['alternativen']}")
    print(f"\nGeschrieben: {config.KONTIERUNG_MAPPING}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
