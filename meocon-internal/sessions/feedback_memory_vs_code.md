---
name: Memory ist kein Produkt-Feature
description: Was in Memory gehoert und was in Code, weil der Beleg-Agent auch fuer andere Kunden gebaut wird
type: feedback
originSessionId: 73c7fa4b-622e-4279-bd84-e0946fb1d309
---
Produkt-Entscheidungen gehoeren ins Code/Config, nicht in meine Memory.

**Why:** Fabio baut den Beleg-Agent auch als Produkt fuer andere Kunden weiter. Die haben kein Claude-Gespraech — meine Memory hilft denen nichts. Als er mich fragte "wie generieren wir das?" war die Antwort klar: Entscheidungen muessen im Code oder in Config-Dateien landen, nicht in einem Notizblock, der nur mir gehoert.

**How to apply:**
- Matching-Logik, Schwellwerte, Vendor-spezifisches Verhalten, OCR-Prompts, Splitting-Regeln → **committed Code**.
- Kunden-spezifische Settings (Bank-Profil, Pfade, API-Keys) → `config_local.py` (gitignored, pro Installation).
- Meine Memory NUR fuer: Fabios persoenliches Setup, Workflow-Praeferenzen, Hinweise damit ich in der naechsten Session nicht blind starte.
- Wenn ich mich dabei ertappe, "wir hatten besprochen dass X" zu memoisieren — Stopp. Pruefen ob X im Code landen muss.
