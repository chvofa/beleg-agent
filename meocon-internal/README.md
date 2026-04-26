# Meocon Internal

Meocon-eigenes Wissen — niemals shared mit externen Tenants.

## Folder

- `sessions/` — Memory-Files aus Claude-Code-Sessions (migriert von `~/.claude/projects/-Users-fabiooro-Developer-beleg-agent/memory/`)

## Status

- 2026-04-26: Initial-Migration der 7 Memory-Files (Snapshot-Stand)
- Auto-Sync-Hook (`~/.claude/.../memory/` → `meocon-internal/sessions/`) noch nicht eingerichtet — neue Memory-Updates müssen vorerst manuell mitkopiert oder der Sync-Hook später aufgesetzt werden

## Pattern-Referenz

Siehe `q_pilot` Repo für die ausführliche Doku des Memory-Migrations-Patterns:
`docs/q_alizer/q_pilot_g1_cleanup_g2_foundation_2026-04.md`
