---
name: Offene Themen die Fabio genannt hat
description: Dinge die besprochen wurden aber (noch) nicht entschieden oder umgesetzt sind. Stand 2026-04-16
type: project
originSessionId: 3ac89a16-e4af-4b66-b12c-ad640a10dbdf
---
**Wiederkehrende Zahlungen als Dauerauftrag erfassen** — Fabio hat erwaehnt, dass folgende Posten eigentlich Daueraufträge sein sollten (eigenes PDF ins `_Dauerauftraege/`-Verzeichnis, einmal erfasst, monatlich absorbiert):
- Sunrise (derzeit pro Monat als Rechnung)
- YouTube Premium / Google One (KK CHF Abos)
- Salt Mobile (laut Offene Posten via eBill auf Bank, nicht KK — Buchungstext: "EBILL-RECHNUNG")

**Why:** Damit diese Posten nicht jeden Monat als offener Posten auftauchen.
**How to apply:** Fabio muss die PDF-Vorlage in `_Dauerauftraege/` legen. Nicht automatisch loesbar — fragen, nicht raten.

**NICHT Dauerauftrag:**
- **HSG Alumni**: jaehrliche Rechnung (150 CHF/Jahr). In offene_posten zwei Eintraege (2026-01-21 + 2026-03-30) — moeglicherweise Duplikatzahlung oder doppelter Upload. Fabio muss manuell klaeren.
- **Jahrespreis Kreditkarte** (200 CHF / 150 EUR): kein Beleg noetig → als "Ignoriert" markieren (bereits erledigt fuer 2026).

**Parkingpay-Sammelbeleg-Workflow:**
- Fabio laedt monatlich einen Parkingpay-Sammelbeleg (1 PDF, Tabelle mit Einzeltransaktionen).
- Ab Commit b381e5b: OCR erkennt `ist_sammelbeleg`, extrahiert `einzelposten` als Liste, schreibt N Zeilen ins Protokoll.
- Noch nicht auf echtem Parkingpay-PDF getestet — Fabio hat seit dem Feature noch keinen neuen hochgeladen.
- Zwei Parkingpay-Eintraege (7.42 + 4.50 CHF) sind noch offen in Offene Posten (April 2026).

**OpenAI 5.81 vs 5.83 Ambiguitaet (gleicher Tag, minimal unterschiedliche Betraege):**
- KK-CSV hat keine Ref-Nr. → Toleranz 0.10 CHF + Reihenfolge-Konsumption als Sicherung.
- Falls Fabio bessere Loesung will: OCR extrahiert Ref auch bei KK-Belegen und Matcher prueft gegen Buchungstext.

**Architektur-Grundsatz:** Fabio will "Lager A" — mehr AI am Ingest, deterministischer Matcher. Kein Agent-SDK, kein Claude pro Matching-Schritt. Grund: Audit-Spur + Kosten.

**ERLEDIGT diese Session (2026-04-16):**
- Debitoren-Modul: Getharvest-Import, Web-UI, Abschreiben, automatische OP-Bereinigung (Commits 4a6b055 + 56a624d)
- Zentiva (EUR 1345.95) als "Abgeschrieben" markiert (zahlt nicht mehr)
- 3 Corden Pharma Kundenzahlungen (CHF 168k + 88k + 13k) aus Offene Posten bereinigt
- bereinige_offene_posten() laeuft automatisch nach jedem Getharvest-Import UND nach jedem Bank-Abgleich
