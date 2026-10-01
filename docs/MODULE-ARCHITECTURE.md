# Module Architecture

**Status:** Canonical product/module boundary rules. The detailed approved session, availability, reporting, and migration design is in [PRODUCT-MODULE-ARCHITECTURE.md](PRODUCT-MODULE-ARCHITECTURE.md); implementation sequencing is maintained in [TAURI-MIGRATION-PLAN.md](TAURI-MIGRATION-PLAN.md).

## Product and Session

The product is **Music Program Scheduler**. A **Scheduling Session** is the one active local project and represents exactly one institution/program and one academic term. “Project” is a user-facing synonym only, not a separate persisted entity. Session metadata is: internal application-generated UUID, institution display name, program display name, term display label, optional year/start/end dates, and created/modified timestamps. No session timezone is required yet.

Only one session/database is active at a time initially. `.mpsession` is the versioned portable archive contract; the internal SQLite schema is not the public file-format contract. Open/Restore replaces the active session after confirmation and creation of a local recovery snapshot. Selective historical import is a separate future feature. Session archives are unencrypted in v1 and may contain student educational information; handle them appropriately.

The initial product shell is:

```text
Music Program Scheduler
  -> Open/Restore or New Session
  -> Session Dashboard
  -> Implemented Scheduling Modules
       -> Accompanist Scheduling
       -> Clinical Placements (when implemented)
       -> Performance Juries (when implemented)
```

The current Accompanist workflow stays inside its module. Do not display empty placeholder workflows to imply functionality that does not exist.

## Module Ownership

Modules are statically registered built-in workflows, not dynamically loaded plugins. A lightweight module descriptor may provide identity/navigation, module schema versions, validation, optimizer operation, report definitions, and declared dependencies on result contracts.

Each module owns its inputs, persistence records/profiles, constraints, validation policy, optimizer, workflow, output model, and report definitions:

- **Accompanist Scheduling:** lessons, pianist-specific profiles/availability, workload, fit/scoring, overlap and manual-assignment policy, assignments, and accompanist reports.
- **Clinical Placements:** placement requirements/history, clinical sites and capacity, transportation, placement validation/scoring, and placement reports. Its optimizer must not be designed until detailed constraints and policies are gathered from the actual Music Therapy workflow owner.
- **Performance Juries:** jury roster, areas/panels, rooms, duration/slot rules, breaks, availability, jury-specific validation/optimizer, and reports. Characterize and test the existing standalone Python jury scheduler before integrating or rewriting it.

Never create a universal optimizer, universal assignment/schedule type, giant universal Person/Student record, or one universal spreadsheet schema.

## Shared Domain and Infrastructure

Share only demonstrated cross-module concepts and mechanics: application-generated UUID Person identity (optional institutional id/email; never merge people solely by display-name match), institution/program and term metadata, time primitives, minimal location/resource identity, import parsing/mapping machinery, validation issue representation, session file handling, availability normalization, report rendering/export primitives, and result envelopes.

Keep role/profile data module-owned around shared identity. A pianist, student, faculty member, room, or clinical site must not accumulate every field used by every module.

### Availability

The shared availability value is a window with a start/end, either a recurring weekday or a dated occurrence, status (`Available`, `Tentative`, `Unavailable`), and optional import provenance. It supports multiple windows on one day and does not assume 30-minute granularity. Missing availability is distinct from explicit `Unavailable`. Same-status overlapping imported windows may be normalized/merged; conflicting-status overlaps produce a validation/review issue and are not silently merged.

Modules retain their own status meaning, scoring, owner associations, validation, and entry UI. The Accompanist grid remains available and retains current solver semantics while becoming the first consumer of shared availability infrastructure. Import CSV/XLS/XLSX exported locally from Microsoft Forms; do not add a Forms/Graph API. Imported rows become normal editable application data.

## Cross-Module Results

Modules exchange structured, versioned results through a session-level result registry rather than importing another module's optimizer or passing spreadsheets. The envelope identifies session, module, result contract/version, state (`draft`, `finalized`, `superseded`), and provenance/revisions; its payload remains module-specific.

Jury may consume a finalized Accompanist assignment result using stable internal Person IDs. Do not use display-name joins as the permanent contract. When a consumed finalized result changes, mark dependent results as potentially stale and surface that state to the user using simple terminology.

## Reporting

Report infrastructure may be shared for field selection, validated filters, sorting/grouping, titles, preview, and local Excel/print/PDF-oriented output. Each module explicitly provides its report definitions, allowed fields, row projection, and semantics. Do not expose arbitrary SQL/database queries or build a universal report designer.

## Persistence and Platform

The existing SQLite/FastAPI service remains transitional. Use versioned schema migrations before session/schema evolution, preserve the current local data as a legacy session with no inferred term, and keep migrations/backups reversible. The `.mpsession` manifest/archive is stable; SQLite remains private implementation payload. Opening/restoring stages and validates the archive before replacing the active session.

The application must work offline and keep student data on-device. Desktop-vs-browser and local-vs-cloud are separate choices. The eventual browser version needs local session storage/file import/export and local execution; it must not depend on a vendor server. No telemetry, analytics, student-data upload, or cloud account is implied.

## Implementation Order

Follow the approved order in [PRODUCT-MODULE-ARCHITECTURE.md](PRODUCT-MODULE-ARCHITECTURE.md): versioned migrations and legacy fixture; session metadata and `.mpsession` new/export/open/restore with recovery snapshots; shared availability using Accompanist first; shared reports adapted from existing Accompanist output; useful product dashboard; independent Jury characterization; Jury integration through finalized results; Clinical data workflow; Clinical optimizer only after owner requirements are gathered.
