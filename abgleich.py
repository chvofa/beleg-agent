#!/usr/bin/env python3
"""
Beleg-Agent – KK-Abgleich
Gleicht Kreditkarten-Transaktionen (CSV) mit dem Belege-Protokoll ab.
"""

import csv
import io
import os
import shutil
import sys
from datetime import datetime, timedelta

import openpyxl

import config

# Windows UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def lade_csv_transaktionen(csv_pfad: str) -> list[dict]:
    """Liest KK CSV-Export (sep=;) und gibt Liste von Transaktionen zurueck."""
    transaktionen = []

    # CSV ist oft Latin-1/CP1252 codiert
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(csv_pfad, "r", encoding=enc) as f:
                inhalt = f.read()
            break
        except UnicodeDecodeError:
            continue
    else:
        with open(csv_pfad, "r", encoding="latin-1") as f:
            inhalt = f.read()

    # Erste Zeile "sep=;" ueberspringen
    zeilen = inhalt.strip().split("\n")
    if zeilen[0].strip().startswith("sep="):
        zeilen = zeilen[1:]

    reader = csv.DictReader(zeilen, delimiter=";")

    def get_col(row, *namen):
        """Findet Spalte nach moeglichen Namen (Umlaut-Varianten)."""
        for n in namen:
            if n in row:
                return row[n].strip() if row[n] else ""
        # Fallback: fuzzy match auf Spaltennamen
        for key in row:
            for n in namen:
                if n.lower() in key.lower() or key.lower() in n.lower():
                    return row[key].strip() if row[key] else ""
        return ""

    for row in reader:
        try:
            betrag = float(get_col(row, "Betrag").replace(",", ".") or "0")
        except ValueError:
            betrag = 0

        datum_str = get_col(row, "Einkaufsdatum")
        try:
            datum = datetime.strptime(datum_str, "%d.%m.%Y").date()
        except ValueError:
            datum = None

        transaktionen.append({
            "datum": datum,
            "buchungstext": get_col(row, "Buchungstext"),
            "betrag": betrag,
            "orig_waehrung": get_col(row, "Originalwährung", "Originalw\xe4hrung", "Originalwaehrung"),
            "kk_waehrung": get_col(row, "Währung", "W\xe4hrung", "Waehrung"),
            "belastung": get_col(row, "Belastung"),
            "gutschrift": get_col(row, "Gutschrift"),
            "buchung": get_col(row, "Buchung"),
            "branche": get_col(row, "Branche"),
        })

    return transaktionen


def erkenne_kk_typ(csv_pfad: str) -> str:
    """Erkennt anhand der Waehrungsspalte ob KK CHF oder KK EUR."""
    trans = lade_csv_transaktionen(csv_pfad)
    if not trans:
        return "?"

    waehrungen = [t["kk_waehrung"] for t in trans if t["kk_waehrung"]]
    if not waehrungen:
        return "?"

    # Mehrheitsentscheid
    from collections import Counter
    haeufigste = Counter(waehrungen).most_common(1)[0][0]
    return f"KK {haeufigste}"


def lade_excel_belege() -> tuple[openpyxl.Workbook, list[dict]]:
    """Laedt das Excel-Protokoll und gibt Workbook + Liste von Beleg-Dicts zurueck."""
    wb = openpyxl.load_workbook(config.EXCEL_PROTOKOLL)
    ws = wb.active

    belege = []
    for row_idx in range(2, ws.max_row + 1):
        datum_str = str(ws.cell(row=row_idx, column=1).value or "").strip()
        try:
            datum = datetime.strptime(datum_str, "%Y-%m-%d").date()
        except ValueError:
            datum = None

        try:
            betrag = float(ws.cell(row=row_idx, column=3).value or 0)
        except (ValueError, TypeError):
            betrag = 0

        belege.append({
            "row": row_idx,
            "datum": datum,
            "rechnungssteller": str(ws.cell(row=row_idx, column=2).value or "").strip(),
            "betrag": betrag,
            "waehrung": str(ws.cell(row=row_idx, column=4).value or "").strip(),
            "typ": str(ws.cell(row=row_idx, column=5).value or "").strip(),
            "zahlungsart": str(ws.cell(row=row_idx, column=6).value or "").strip(),
            "abgeglichen": str(ws.cell(row=row_idx, column=10).value or "").strip(),
        })

    return wb, belege


def match_transaktion_zu_beleg(trans: dict, belege: list[dict], kk_typ: str) -> dict | None:
    """Versucht eine KK-Transaktion einem Beleg zuzuordnen (fuzzy matching)."""
    buch = trans["buchungstext"].upper()
    t_betrag = trans["betrag"]
    t_datum = trans["datum"]

    beste_matches = []

    for beleg in belege:
        b_betrag = beleg["betrag"]
        b_datum = beleg["datum"]
        b_rs = beleg["rechnungssteller"].upper()

        # 1. Betrag muss stimmen (kleine Toleranz fuer Rundung)
        if abs(t_betrag - b_betrag) > 0.10:
            continue

        # 2. Datums-Toleranz: KK-Einkaufsdatum kann +/- 5 Tage vom Rechnungsdatum abweichen
        if t_datum and b_datum:
            diff = abs((t_datum - b_datum).days)
            if diff > 30:
                continue
            datum_score = max(0, 30 - diff) / 30  # 1.0 bei gleichem Tag, 0.0 bei 30 Tagen
        else:
            datum_score = 0.3  # Kein Datum -> schwacher Match

        # 3. Name-Matching (Teilstring)
        name_score = 0
        rs_teile = [t for t in b_rs.split() if len(t) > 2]
        for teil in rs_teile:
            if teil in buch:
                name_score += 1
        if rs_teile:
            name_score = name_score / len(rs_teile)

        # Mindestens Namens- ODER Datumsaehlichkeit
        gesamt_score = (name_score * 0.6) + (datum_score * 0.4)
        if name_score > 0 or (datum_score > 0.8 and abs(t_betrag - b_betrag) < 0.02):
            beste_matches.append((gesamt_score, beleg))

    if not beste_matches:
        return None

    # Bester Match
    beste_matches.sort(key=lambda x: x[0], reverse=True)
    return beste_matches[0][1]


def main():
    print()
    print("=" * 60)
    print("  BELEG-AGENT - KK-Abgleich")
    print("=" * 60)
    print()

    if not os.path.exists(config.EXCEL_PROTOKOLL):
        print("FEHLER: Excel-Protokoll nicht gefunden!")
        return

    # 1. CSV-Dateien im Abgleich-Ordner finden
    abgleich_pfad = config.ABGLEICH_PFAD
    os.makedirs(abgleich_pfad, exist_ok=True)

    csv_dateien = [f for f in os.listdir(abgleich_pfad) if f.lower().endswith(".csv")]
    if not csv_dateien:
        print("Keine CSV-Dateien im Abgleich-Ordner gefunden.")
        return

    print(f"Gefunden: {len(csv_dateien)} CSV-Datei(en)\n")

    # 2. KK-Typ erkennen und Dateien umbenennen
    archiv_pfad = os.path.join(abgleich_pfad, "archiv")
    os.makedirs(archiv_pfad, exist_ok=True)

    kk_dateien = {}  # {"KK CHF": pfad, "KK EUR": pfad}

    for datei in csv_dateien:
        pfad = os.path.join(abgleich_pfad, datei)
        kk_typ = erkenne_kk_typ(pfad)
        print(f"  {datei} -> {kk_typ}")
        if kk_typ in ("KK CHF", "KK EUR"):
            kk_dateien[kk_typ] = pfad
        else:
            print(f"    WARNUNG: Konnte KK-Typ nicht erkennen, ueberspringe.")

    print()

    # 3. Excel laden
    wb, belege = lade_excel_belege()
    ws = wb.active

    gesamt_matches = 0
    gesamt_ohne_beleg = 0
    gesamt_neu_za = 0
    ohne_beleg_liste = []

    for kk_typ, csv_pfad in sorted(kk_dateien.items()):
        print(f"--- {kk_typ} ---")
        transaktionen = lade_csv_transaktionen(csv_pfad)

        # Nur 2026er Transaktionen (passend zu unseren Belegen)
        trans_2026 = [t for t in transaktionen if t["datum"] and t["datum"].year >= 2026]
        print(f"  {len(transaktionen)} Transaktionen total, {len(trans_2026)} ab 2026\n")

        for trans in trans_2026:
            # Gebuehren/Zuschlaege ueberspringen
            if not trans["buchungstext"] or "ZUSCHLAG" in trans["buchungstext"].upper():
                continue

            match = match_transaktion_zu_beleg(trans, belege, kk_typ)

            datum_str = trans["datum"].strftime("%d.%m.%Y") if trans["datum"] else "?"
            orig_w = trans["orig_waehrung"] or trans["kk_waehrung"]

            if match:
                row = match["row"]
                bereits = match["abgeglichen"] == "Ja"

                if bereits:
                    # Schon abgeglichen - still ueberspringen
                    continue

                # Match gefunden -> Abgeglichen = Ja
                ws.cell(row=row, column=10).value = "Ja"

                # Zahlungsart ergaenzen wenn leer
                alte_za = ws.cell(row=row, column=6).value or ""
                if not alte_za:
                    ws.cell(row=row, column=6).value = kk_typ
                    gesamt_neu_za += 1
                    za_info = f" [Zahlungsart -> {kk_typ}]"
                else:
                    za_info = ""

                # PayPal ergaenzen
                if "PAYPAL" in trans["buchungstext"].upper():
                    ws.cell(row=row, column=7).value = "Ja"

                # Beleg als abgeglichen markieren (in-memory auch updaten)
                match["abgeglichen"] = "Ja"

                print(f"  NEU:  {datum_str} {orig_w} {trans['betrag']:>10.2f}  {trans['buchungstext'][:40]}")
                print(f"        -> {match['rechnungssteller']} ({match['waehrung']} {match['betrag']}){za_info}")
                gesamt_matches += 1
            else:
                print(f"  KEIN BELEG: {datum_str} {orig_w} {trans['betrag']:>10.2f}  {trans['buchungstext'][:50]}")
                ohne_beleg_liste.append({
                    "kk": kk_typ,
                    "datum": datum_str,
                    "betrag": trans["betrag"],
                    "waehrung": orig_w,
                    "text": trans["buchungstext"][:60],
                })
                gesamt_ohne_beleg += 1

        print()

    # 4. Excel speichern
    wb.save(config.EXCEL_PROTOKOLL)
    wb.close()

    # 5. CSVs archivieren
    datum_str = datetime.now().strftime("%Y-%m-%d")
    for kk_typ, csv_pfad in kk_dateien.items():
        ziel_name = f"{kk_typ.replace(' ', '_')}_{datum_str}.csv"
        ziel = os.path.join(archiv_pfad, ziel_name)
        shutil.move(csv_pfad, ziel)
        print(f"Archiviert: {os.path.basename(csv_pfad)} -> archiv/{ziel_name}")

    # 6. Zusammenfassung
    print(f"\n{'='*60}")
    print(f"  ZUSAMMENFASSUNG")
    print(f"{'='*60}")
    print(f"  Abgeglichen:              {gesamt_matches}")
    print(f"  Zahlungsart ergaenzt:      {gesamt_neu_za}")
    print(f"  KK-Transaktionen ohne Beleg: {gesamt_ohne_beleg}")

    if ohne_beleg_liste:
        print(f"\n  Transaktionen OHNE passenden Beleg:")
        print(f"  {'KK':<8} {'Datum':<12} {'Betrag':>10} {'Text'}")
        print(f"  {'-'*8} {'-'*12} {'-'*10} {'-'*50}")
        for t in ohne_beleg_liste:
            print(f"  {t['kk']:<8} {t['datum']:<12} {t['waehrung']} {t['betrag']:>8.2f} {t['text']}")

    print(f"\nExcel aktualisiert: {config.EXCEL_PROTOKOLL}")


if __name__ == "__main__":
    main()
