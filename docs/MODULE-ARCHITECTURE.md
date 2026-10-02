# Module Architecture

**Status:** Canonical product/module boundary rules. Detailed session, availability, and reporting decisions are in [PRODUCT-MODULE-ARCHITECTURE.md](PRODUCT-MODULE-ARCHITECTURE.md); current runtime and schema details are in [ARCHITECTURE.md](ARCHITECTURE.md).

## Product and Session

The product is **Music Program Scheduler**. A **Scheduling Session** is the one active local project and represents exactly one institution/program and one academic term. “Project” is a user-facing synonym only, not a separate persisted entity. Session metadata is: internal application-generated UUID, institution display name, program display name, term display label, optional year/start/end dates, and created/modified timestamps. No session timezone is required yet.

Only one session/database is active at a time initially. `.mpsession` is the versioned portable archive contract; the internal SQLite schema is not the public file-format contract. Open/Restore replaces the active session after confirmation and creation of a local recovery snapshot. Selective historical import is a separate future feature. Session archives are unencrypted in v1 and may contain student educational information; handle them appropriately.

The initial product shell starts on the active-session dashboard. Session operations are owned by the product level, not repeated inside module toolbars:

```text
Music Program Scheduler
  -> Active-session dashboard (session identity + Edit/New/Open/Export)
  -> Implemented Scheduling Modules
       -> Accompanist Scheduling
       -> Clinical Placements (module shell)
      -> Performance Juries (lesson-based setup and readiness)
```

The Accompanist module has its own header, return-to-dashboard control, and internal Import Lessons, Pianists & Availability, Schedule, and Reports navigation. Session actions do not remain in that module header. Performance Juries provides lesson-based setup and readiness tabs in Overview, Panels, Pianist Availability, Lessons, and Schedule order; Overview is the initial view. Schedule generation is available through the backend module API; the Schedule UI remains a placeholder. Clinical Placements remains a minimal honest shell and must not display fake workflow state or statistics. Replacing/restoring a session returns the user to the dashboard. Export does not change navigation state.

## Module Ownership

Modules are statically registered built-in workflows, not dynamically loaded plugins. A lightweight module descriptor may provide identity/navigation, module schema versions, validation, optimizer operation, report definitions, and declared dependencies on result contracts.

Each module owns its inputs, persistence records/profiles, constraints, validation policy, optimizer, workflow, output model, and report definitions:

- **Accompanist Scheduling:** lessons, the authoritative Needs pianist? Boolean, optional Specific pianist name, source-owned Jury Required value, pianist profiles/availability, workload, fit/scoring, overlap and manual-assignment policy, assignments, and accompanist reports.
- **Clinical Placements:** placement requirements/history, clinical sites and capacity, transportation, placement validation/scoring, and placement reports. Its optimizer must not be designed until detailed constraints and policies are gathered from the actual Music Therapy workflow owner.
- **Performance Juries:** lesson-based entries/roster, manually assigned Panels, one Scheduling Date per Panel, duration/break/meal configuration, date-scoped pianist Availability Windows, readiness validation, a pure-domain optimizer core, persisted typed schedule results, dependency/staleness tracking, and Jury-owned reports. Jury Required and Panel selection belong to each source lesson entry, not globally to a student. The standalone scheduler remains characterized and unused; the domain and integration boundaries are documented in [JURY-OPTIMIZER-CONTRACT.md](JURY-OPTIMIZER-CONTRACT.md). Schedule UI/manual editing/finalization remain unimplemented.

Never create a universal optimizer, universal assignment/schedule type, giant universal Person/Student record, or one universal spreadsheet schema.

## Shared Domain and Infrastructure

Share only demonstrated cross-module concepts and mechanics: application-generated UUID Person identity (optional institutional id/email; never merge people solely by display-name match), institution/program and term metadata, time primitives, minimal location/resource identity, import parsing/mapping machinery, validation issue representation, session file handling, availability normalization, report rendering/export primitives, and result envelopes.

Keep role/profile data module-owned around shared identity. A pianist, student, faculty member, room, or clinical site must not accumulate every field used by every module.

### Availability

The shared availability value is a window with a start/end, either a recurring weekday or a dated occurrence, status (`Available`, `Tentative`, `Unavailable`), and optional import provenance. It supports multiple windows on one day and does not assume 30-minute granularity. Submission completeness is separate from status: absent windows are unknown for an incomplete horizon and derive `Unavailable` for a valid complete horizon. Same-status overlapping imported windows may be normalized/merged; conflicting-status overlaps produce a validation/review issue and are not silently merged.

Modules retain their own status meaning, scoring, owner associations, validation, and entry UI. The Accompanist grid remains available and retains current solver semantics while becoming the first consumer of shared availability infrastructure. The shared parser accepts local CSV/XLSX/XLS, normalized rows or configurable wide Forms-style rows, and rejects conflicting overlap. Accompanist adapts only 30-minute-aligned windows into existing pianist slots and records weekly completeness in an Accompanist-specific marker; no shared owner table or session foreign keys are introduced. Pianist Availability import groups rows only by supplied Pianist Name after trimming surrounding whitespace and collapsing repeated spaces. It does not reconcile against the old roster: applying creates fresh Pianist/Person identities, replaces the entire Pianist roster and weekly Availability Windows, and clears Lesson-to-Pianist assignments. New Pianists have blank email unless explicitly mapped and start with 40 Max Hours Per Week. Internal identifiers are not shown or requested in user workflows. Import CSV/XLS/XLSX exported locally from Microsoft Forms; do not add a Forms/Graph API. Imported rows become normal editable application data.

## Cross-Module Results

Modules exchange structured, versioned results through a session-level result registry rather than importing another module's optimizer or passing spreadsheets. The envelope identifies session, module, result contract/version, state (`draft`, `finalized`, `superseded`), and provenance/revisions; its payload remains module-specific.

Jury consumes the typed, versioned `accompanist.assignment-result` contract v2 using stable Person and Lesson UUIDs. Each source lesson retains its Jury Required value, Needs pianist? Boolean, optional Specific pianist name, and assigned Pianist. Accompanist owns these source Lesson fields; Jury Required is editable in both modules through the same source-owned service, with no Jury-owned Boolean copy. Jury owns only the per-lesson Panel selection, each Panel's one-day Jury Date, and date-keyed Jury-Day Availability Windows. A Jury-required lesson with Needs pianist? = Yes but no finalized assignment is a blocking readiness problem. Jury never substitutes a pianist. Jury-Day availability is a Jury-owned, binary, closed-world input, separate from Accompanist Tentative semantics. A declaration is complete only when it contains one or more Availability Windows; empty or invalid submissions remain incomplete. A stale Accompanist snapshot blocks readiness until reviewed and finalized again.

## Reporting

Report infrastructure may be shared for field selection, validated filters, sorting/grouping, titles, preview, and local Excel/print/PDF-oriented output. Each module explicitly provides its report definitions, allowed fields, row projection, and semantics. Do not expose arbitrary SQL/database queries or build a universal report designer.

## Persistence and Platform

The existing SQLite/FastAPI service remains transitional. Use versioned schema migrations before session/schema evolution, preserve the current local data as a legacy session with no inferred term, and keep migrations/backups reversible. The `.mpsession` manifest/archive is stable; SQLite remains private implementation payload. Opening/restoring stages and validates the archive before replacing the active session.

The application must work offline and keep student data on-device. Desktop-vs-browser and local-vs-cloud are separate choices. The eventual browser version needs local session storage/file import/export and local execution; it must not depend on a vendor server. No telemetry, analytics, student-data upload, or cloud account is implied.

## Implementation Order

Follow the approved order in [PRODUCT-MODULE-ARCHITECTURE.md](PRODUCT-MODULE-ARCHITECTURE.md): versioned migrations and session archives; shared availability and reporting; product dashboard; Jury characterization (complete) and Jury data integration through finalized Accompanist results/readiness (complete); later Jury optimizer and UI milestones; Clinical data workflow; Clinical optimizer only after owner requirements are gathered.
