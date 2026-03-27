#!/usr/bin/env python3
"""
Beleg-Agent – Bank-Abgleich
Gleicht Bankkonto-Transaktionen (CSV) mit dem Belege-Protokoll ab.
"""

import csv
import os
import re
import shutil
import sys
from datetime import datetime
from collections import defaultdict

import openpyxl

import config

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def lade_bank_transaktionen(csv_pfad: str) -> tuple[str, list[dict]]:
    """Liest Bankkonto CSV-Export. Gibt (waehrung, transaktionen) zurueck."""

    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(csv_pfad, "r", encoding=enc) as f:
                inhalt = f.read()
            break
        except UnicodeDecodeError:
            continue

    zeilen = inhalt.strip().split("\n")

    # Header-Block parsen (Zeilen 1-8 vor der leeren Zeile)
    waehrung = ""
    for zeile in zeilen[:10]:
        if zeile.startswith("Bewertet in"):
            waehrung = zeile.split(";")[1].strip()
            break

    # Datenzeilen finden (nach der Header-Zeile mit "Abschlussdatum")
    daten_start = None
    for i, zeile in enumerate(zeilen):
        if zeile.startswith("Abschlussdatum;"):
            daten_start = i
            break

    if daten_start is None:
        return waehrung, []

    # CSV parsen ab Datenzeilen
    reader = csv.DictReader(zeilen[daten_start:], delimiter=";")

    transaktionen = []
    aktuelle_sammel = None  # Fuer Sammelauftraege (Zeilen ohne Datum)

    for row in reader:
        datum_str = (row.get("Abschlussdatum") or "").strip()
        beschr1 = (row.get("Beschreibung1") or "").strip().strip('"')
        beschr2 = (row.get("Beschreibung2") or "").strip().strip('"')
        beschr3 = (row.get("Beschreibung3") or "").strip().strip('"')

        belastung_str = (row.get("Belastung") or "").strip()
        gutschrift_str = (row.get("Gutschrift") or "").strip()
        einzelbetrag_str = (row.get("Einzelbetrag") or "").strip()

        # Betrag bestimmen
        betrag = 0
        ist_gutschrift = False
        if gutschrift_str:
            try:
                betrag = abs(float(gutschrift_str.replace(",", ".")))
                ist_gutschrift = True
            except ValueError:
                pass
        elif belastung_str:
            try:
                betrag = abs(float(belastung_str.replace(",", ".").replace("-", "")))
            except ValueError:
                pass
        elif einzelbetrag_str:
            try:
                val = float(einzelbetrag_str.replace(",", "."))
                betrag = abs(val)
                ist_gutschrift = val > 0
            except ValueError:
                pass

        if betrag == 0:
            continue

        # Datum parsen
        if datum_str:
            try:
                datum = datetime.strptime(datum_str, "%Y-%m-%d").date()
            except ValueError:
                datum = None

            # Bankgebuehren und Saldo-Abschlüsse ueberspringen
            if "Dienstleistungspreisabschluss" in beschr1 or "Depotpreis" in beschr1:
                continue
            # KK-Rechnungen ueberspringen (die werden separat abgeglichen)
            if "KREDITKARTEN-RECHNUNG" in beschr3:
                continue
            # FX Spot ueberspringen
            if "Kauf FX Spot" in beschr2:
                continue

            transaktionen.append({
                "datum": datum,
                "betrag": betrag,
                "ist_gutschrift": ist_gutschrift,
                "beschreibung": f"{beschr1} {beschr2}".strip(),
                "details": beschr3,
                "waehrung": waehrung,
            })
        else:
            # Unterzeile eines Sammelauftrags
            transaktionen.append({
                "datum": transaktionen[-1]["datum"] if transaktionen else None,
                "betrag": betrag,
                "ist_gutschrift": ist_gutschrift,
                "beschreibung": f"{beschr1} {beschr2}".strip(),
                "details": beschr3,
                "waehrung": waehrung,
            })

    return waehrung, transaktionen


def erkenne_bank_typ(csv_pfad: str) -> str:
    """Erkennt Waehrung aus Header."""
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(csv_pfad, "r", encoding=enc) as f:
                for line in f:
                    if line.startswith("Bewertet in"):
                        return line.split(";")[1].strip()
            break
        except UnicodeDecodeError:
            continue
    return "?"


def match_bank_transaktion(trans: dict, belege: list[dict]) -> dict | None:
    """Versucht eine Bank-Transaktion einem Beleg zuzuordnen."""
    t_betrag = trans["betrag"]
    t_datum = trans["datum"]
    t_beschr = trans["beschreibung"].upper()
    t_details = trans["details"].upper()
    t_ist_gs = trans["ist_gutschrift"]

    beste_matches = []

    for beleg in belege:
        b_betrag = beleg["betrag"]
        b_datum = beleg["datum"]
        b_rs = beleg["rechnungssteller"].upper()
        b_typ = beleg["typ"]

        # Betrag muss stimmen
        if abs(t_betrag - b_betrag) > 1.0:  # Etwas mehr Toleranz fuer Bank
            continue

        # Typ-Check: Gutschrift in Bank = Gutschrift im Beleg (oder umgekehrt)
        if t_ist_gs and b_typ == "Rechnung":
            continue
        if not t_ist_gs and b_typ == "Gutschrift":
            continue

        # Datums-Toleranz: Bank kann ein paar Tage abweichen
        if t_datum and b_datum:
            diff = abs((t_datum - b_datum).days)
            if diff > 45:
                continue
            datum_score = max(0, 45 - diff) / 45
        else:
            datum_score = 0.3

        # Name-Matching in Beschreibung
        name_score = 0
        rs_teile = [t for t in b_rs.split() if len(t) > 2]
        for teil in rs_teile:
            if teil in t_beschr or teil in t_details:
                name_score += 1
        if rs_teile:
            name_score = name_score / len(rs_teile)

        # Betrag-Genauigkeit
        betrag_score = 1.0 if abs(t_betrag - b_betrag) < 0.05 else 0.8

        gesamt = (name_score * 0.5) + (datum_score * 0.3) + (betrag_score * 0.2)

        if name_score > 0 or (datum_score > 0.8 and betrag_score > 0.9):
            beste_matches.append((gesamt, beleg))

    if not beste_matches:
        return None

    beste_matches.sort(key=lambda x: x[0], reverse=True)
    return beste_matches[0][1]


def main():
    print()
    print("=" * 60)
    print("  BELEG-AGENT - Bank-Abgleich")
    print("=" * 60)
    print()

    abgleich_pfad = config.ABGLEICH_PFAD
    os.makedirs(abgleich_pfad, exist_ok=True)
    archiv_pfad = os.path.join(abgleich_pfad, "archiv")
    os.makedirs(archiv_pfad, exist_ok=True)

    # Bank-CSVs finden (nicht KK-Archiv-Dateien)
    csv_dateien = []
    for f in os.listdir(abgleich_pfad):
        if not f.lower().endswith(".csv"):
            continue
        pfad = os.path.join(abgleich_pfad, f)
        # Check ob es ein Bank-Export ist (Header "Kontonummer:")
        try:
            with open(pfad, "r", encoding="cp1252") as fh:
                erste_zeile = fh.readline()
            if "Kontonummer" in erste_zeile:
                csv_dateien.append(f)
        except Exception:
            pass

    if not csv_dateien:
        print("Keine Bank-CSV-Dateien gefunden.")
        return

    print(f"Gefunden: {len(csv_dateien)} Bank-CSV-Datei(en)\n")

    # Excel laden
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
            "typ": str(ws.cell(row=row_idx, column=5).value or "Rechnung").strip(),
            "zahlungsart": str(ws.cell(row=row_idx, column=6).value or "").strip(),
            "abgeglichen": str(ws.cell(row=row_idx, column=10).value or "").strip(),
        })

    gesamt_matches = 0
    gesamt_neu_za = 0
    ohne_beleg_liste = []

    for datei in sorted(csv_dateien):
        pfad = os.path.join(abgleich_pfad, datei)
        waehrung = erkenne_bank_typ(pfad)
        bank_typ = f"Bank {waehrung}"

        print(f"--- {bank_typ} ({datei}) ---")
        waehrung_csv, transaktionen = lade_bank_transaktionen(pfad)

        # Nur 2026
        trans_2026 = [t for t in transaktionen if t["datum"] and t["datum"].year >= 2026]
        print(f"  {len(transaktionen)} Transaktionen total, {len(trans_2026)} ab 2026\n")

        for trans in trans_2026:
            match = match_bank_transaktion(trans, belege)

            datum_str = trans["datum"].strftime("%d.%m.%Y") if trans["datum"] else "?"
            gs = "+" if trans["ist_gutschrift"] else "-"
            beschr_kurz = trans["beschreibung"][:50]

            if match:
                bereits = match["abgeglichen"] == "Ja"
                if bereits:
                    continue

                row = match["row"]
                ws.cell(row=row, column=10).value = "Ja"

                alte_za = ws.cell(row=row, column=6).value or ""
                if not alte_za:
                    ws.cell(row=row, column=6).value = "Überweisung"
                    gesamt_neu_za += 1
                    za_info = " [Zahlungsart -> Überweisung]"
                else:
                    za_info = ""

                match["abgeglichen"] = "Ja"

                print(f"  NEU:  {datum_str} {gs}{trans['betrag']:>10.2f} {waehrung}  {beschr_kurz}")
                print(f"        -> {match['rechnungssteller']} ({match['waehrung']} {match['betrag']}){za_info}")
                gesamt_matches += 1
            else:
                ohne_beleg_liste.append({
                    "bank": bank_typ,
                    "datum": datum_str,
                    "betrag": trans["betrag"],
                    "gs": gs,
                    "text": beschr_kurz,
                })

        print()

    # Speichern
    wb.save(config.EXCEL_PROTOKOLL)
    wb.close()

    # Archivieren
    datum_str = datetime.now().strftime("%Y-%m-%d")
    for datei in csv_dateien:
        pfad = os.path.join(abgleich_pfad, datei)
        waehrung = erkenne_bank_typ(pfad)
        ziel_name = f"Bank_{waehrung}_{datum_str}.csv"
        ziel = os.path.join(archiv_pfad, ziel_name)
        shutil.move(pfad, ziel)
        print(f"Archiviert: {datei} -> archiv/{ziel_name}")

    # Zusammenfassung
    print(f"\n{'='*60}")
    print(f"  ZUSAMMENFASSUNG")
    print(f"{'='*60}")
    print(f"  Abgeglichen:           {gesamt_matches}")
    print(f"  Zahlungsart ergaenzt:   {gesamt_neu_za}")
    print(f"  Ohne Beleg:            {len(ohne_beleg_liste)} (nicht alle brauchen einen)")


if __name__ == "__main__":
    main()
