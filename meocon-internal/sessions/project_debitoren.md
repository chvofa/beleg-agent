---
name: Debitoren-Modul (Ausgangsrechnungen)
description: Wie das Debitoren-Modul funktioniert — Getharvest-Import, Bereinigung, Abschreiben
type: project
originSessionId: 3ac89a16-e4af-4b66-b12c-ad640a10dbdf
---
Debitoren-Modul wurde 2026-04-16 gebaut. Quelldatei: `abgleich_debitoren.py`.

**Datenquelle:** Getharvest CSV-Export (`harvest_invoice_report.csv`).
**Sheet:** "Debitoren" in `Belege_Protokoll.xlsx`.
**Spalten:** Harvest_ID (String, z.B. "CPI-2025-15"), Rechnungsdatum, Faellig, Kunde, Betreff, Betrag, Waehrung, Bezahlt_am, Status.

**Import-Logik:**
- Balance > 0 → neu als "Offen" erfassen (falls noch nicht im Sheet)
- Balance = 0 + in letzten 12 Monaten bezahlt → als "Bezahlt" erfassen (Kundennamen-DB fuer Bereinigung)
- Balance = 0 + aelter als 12 Monate → uebersprungen
- Bestehende "Offen"-Eintraege die bezahlt wurden → auf "Bezahlt" aktualisieren
- Am Ende: bereinige_offene_posten() aufrufen

**Automatische Bereinigung (bereinige_offene_posten()):**
- Laedt alle Kundennamen aus dem Debitoren-Sheet
- Sucht in Offene_Posten nach Bank-Eintraegen mit "Gutschrift" im Buchungstext
- Wenn Kundenname (mind. 2 signifikante Woerter) im Buchungstext → Status = "Ignoriert", Grund = "Debitorenzahlung"
- Wird aufgerufen: nach Getharvest-Import UND nach jedem Bank-Abgleich (abgleich_bank.py)

**Abschreiben:**
- `abschreiben_standalone(harvest_id, notiz)` → Status = "Abgeschrieben", Notiz im Betreff-Feld
- Eintraege bleiben sichtbar (Audit-Spur), zaehlen nicht mehr als offen

**Web-UI:**
- Seite: /debitoren
- API: GET /api/debitoren, POST /api/debitoren/<id>/abschreiben, POST /api/reconciliation/debitoren
- Upload-Endpoint: POST /api/upload-csv-inbox (CSV direkt in _Inbox)
- Dashboard-Kachel: "Debitoren offen" (gelb wenn > 0)

**Aktueller Stand (2026-04-16):**
- ZTV-2024-01: Zentiva, EUR 1345.95 → Abgeschrieben
- CPI-2025-15: Corden Pharma International GmbH, CHF 7417.50 → Offen (faellig 2026-05-11)
- 19 weitere Eintraege als "Bezahlt" (letzte 12 Monate, fuer Kundennamen-DB)

**Why:** Fabio sendet monatlich Rechnungen an viele Kunden (hauptsaechlich Corden Pharma International + Switzerland). Zahlungseingang via Bank (Gutschriften). Getharvest ist Quelle der Wahrheit fuer Paid-Status — kein Bank-Abgleich auf Debitor-Seite noetig (Option A).
