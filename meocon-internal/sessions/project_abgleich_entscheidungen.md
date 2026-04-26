---
name: Abgleich-Entscheidungen und Rolling-Export-Logik
description: Warum der Matcher so tickt wie er tickt — Datumsfenster, Cleanup, Pass 2, Daueraufträge, Sammelauftraege
type: project
originSessionId: 73c7fa4b-622e-4279-bd84-e0946fb1d309
---
Der Abgleich-Teil ist iterativ gewachsen. Die wichtigsten Entscheidungen:

**Rolling-Exports sind der Default, nicht die Ausnahme.** UBS Bank-CSV enthaelt immer den kompletten rollenden Zeitraum (z.B. 2024-01 bis heute), nicht nur neue Transaktionen seit dem letzten Lauf. Der Matcher muss damit rechnen, dass dieselbe Transaktion in mehreren CSVs vorkommt.

**Why:** Vor dem Fix landeten bei jedem Re-Import hunderte Duplikate in `offene_posten`. Die Loesung ist ein 2-Pass-Matcher:
- Pass 1: sucht gegen nicht-abgeglichene Belege. Erfolg → Match eintragen.
- Pass 2 (Recall, include_matched=True): sucht gegen ALLE Belege. Erfolg → leise uebergehen (es ist eine Wiederholung aus einem frueheren Lauf).
- Nur wenn beide Paesse nichts finden → neuer offener Posten.

**Cleanup hat 120 Tage Fenster.** B2B-Rechnungen (OXYGY, Baloise, HRD Stucky) kommen mit 60-90 Tagen Zahlungsfrist. Cleanup (offene_posten.resolve) und Bank-Matcher nutzen 120 Tage. KK-Matcher 90 Tage. Schutz vor False Positives: Name-Substring-Match ist Pflicht bei weitem Fenster.

**How to apply:**
- Beim Fehlersuchen in `offene_posten`: zuerst schauen ob der Beleg im Protokoll existiert und `Abgeglichen=Ja` hat. Wenn ja, ist es ein Cleanup-Thema, kein Matching-Thema.
- Toleranzen nicht unbedacht aufweiten — bei B2B-Rundungsdifferenzen (z.B. Baloise 10758.75 vs 10758.70) ist 0.10 CHF die Schmerzgrenze. Mehr → Name-Check wird wichtiger.
- Daueraufträge (typ="Dauerauftrag") sind wiederverwendbare Templates: Matcher ignoriert den Datums-Check komplett und matcht nur auf Name+Betrag, damit sie monatliche Wiederholungen abdecken.
- Sammelauftrag-Master (UBS: Master-Zeile mit Aggregat-Betrag, danach Kind-Zeilen ohne Datum): `ist_sammelauftrag_master=True` wird strukturell gesetzt und der Master im Matching uebersprungen.

**Referenz-Nr ist ab Commit b381e5b der Priority-1-Match-Signal.** Wenn Bank-Transaktion und Beleg dieselbe Transaktions-Nr haben, ist das der kuerzeste Pfad und exakt. Nur UBS Bank, nicht UBS KK — deren CSV hat keine Ref.
