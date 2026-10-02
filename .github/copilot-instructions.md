# Music Program Scheduler - Copilot Instructions

## Product

This project is a modular scheduling application for university and conservatory music programs.

The current production module schedules accompanists for student lessons.

Planned future modules include:
- Music Therapy Clinical Placements
- Performance Jury Scheduling

Other music-program scheduling modules may be added later.

Do not treat this application internally as only an "Accompanist Scheduler."

The accompanist scheduler is a module within the larger Music Program Scheduler product.

One Scheduling Session represents exactly one institution/program and one academic term. “Project” is a user-facing synonym, not a separate persisted entity. Initially, only one session/database is active at a time; do not build an in-app session history library.


## Core Architectural Principle

GENERALIZE SHARED DOMAIN CONCEPTS AND INFRASTRUCTURE.

DO NOT CREATE ONE UNIVERSAL OPTIMIZATION ALGORITHM.

Different scheduling modules may have fundamentally different:
- input data
- constraints
- objective functions
- algorithms
- workflows
- outputs

Shared concepts may include:
- academic terms
- people
- students
- faculty
- staff
- locations
- rooms
- sites
- dates/times
- availability
- travel
- assignments
- imports/exports
- validation

Keep module-specific concepts inside their modules.


## Application Structure

Prefer conceptually:

Music Program Scheduler
    |
    +-- Shared Domain
    |
    +-- Scheduling Session
    |     +-- Institution / Program Context
    |     +-- Academic Term
    |
    +-- Scheduling Modules
    |     +-- Accompanists
    |     +-- Clinical Placements
    |     +-- Juries
    |
    +-- Shared Infrastructure
    |
    +-- Platform Abstraction
          +-- Tauri
          +-- Future Browser

Major scheduling modules should be separate workflows/modules.

Tabs may be used INSIDE modules.

Do not model the major scheduling modules merely as tabs of one universal scheduler.


## Shared vs Module-Specific Data

Do not create giant universal Student, Person, Assignment, or Schedule types containing every field that any future module might need.

Prefer:
- small shared identities/concepts
- module-specific data layered on top

Only promote a concept into the shared domain when it is genuinely shared.

Use an application-generated UUID as the authoritative internal Person identity. Institutional identifiers and email are optional. Never automatically merge people solely because display names match. Keep Accompanist, Clinical, and Jury profiles/data module-specific around shared identity.


## Optimization

Each scheduling problem may have its own optimizer.

Examples:
- AccompanistOptimizer
- ClinicalPlacementOptimizer
- JuryOptimizer

Reusable primitives/utilities are encouraged where genuinely useful.

Do not distort one scheduling problem to make it fit another algorithm.

Keep scheduling logic independent of:
- UI
- Tauri
- HTTP
- persistence
- remote servers

If implemented in Rust, prefer pure Tauri-independent solver/domain crates.


## Cross-Module Data

Modules may consume finalized results from other modules.

Known example:

Accompanist Assignments -> Performance Jury Scheduling

Prefer structured internal result sharing over exporting and reimporting spreadsheets.

Avoid tightly coupling one module's implementation directly to another module.

Represent result lifecycle internally as draft/finalized/superseded while keeping user-facing labels simple. If a finalized result consumed by another module changes, mark dependent results potentially stale and surface that state. Use structured, versioned result contracts rather than spreadsheet handoffs.


## Sessions and Academic Terms

Each Scheduling Session is scoped to one academic term and one institution/program display context. Session metadata includes an internal UUID, institution display name, program display name, term display label, optional year/start/end dates, and created/modified timestamps. Exact term dates and timezone are not required initially.

Examples:
- Fall 2026
- Spring 2027

Module data and results belong to the active session/term. Do not infer terms from lesson dates or filenames. The portable extension is `.mpsession`; its versioned archive/manifest is the contract, not internal SQLite tables. Open/Restore replaces the active session after confirmation and an automatic local recovery snapshot. Selective historical import is separate future work. V1 archives are unencrypted and may contain student educational information; handle them appropriately.


## Platform Architecture

This is currently a Tauri 2 + React + TypeScript + Rust desktop application.

Do not tightly couple React to Tauri.

Platform-specific capabilities must go through clear service/platform abstractions where practical.

Examples:
- FileService
- SessionStorage
- SettingsService
- PlatformService
- UpdateService

UI components should not directly invoke native/platform functionality throughout the codebase.


## Future Web Version

The architecture must preserve a future browser/SaaS deployment.

Treat these as independent choices:
- Desktop vs Web
- Local vs Cloud

A future browser version should be capable of:
- processing imports locally
- running scheduling locally
- storing projects locally
- generating exports locally

A web version must not require student scheduling data to be stored on our servers.

Do not introduce a server requirement for functionality that can operate locally.

If Rust scheduling engines are used, preserve reasonable separation so future WebAssembly compilation remains possible.

Do not implement WebAssembly prematurely.

Future browser support must use local file import/export, local session storage, and local execution; it must not require a vendor server or upload student scheduling data.


## Privacy

This application handles student scheduling information.

Fundamental principle:

> The application provider should not need possession of student scheduling data in order to provide the scheduling service.

Therefore:

- Core scheduling must work offline.
- Student scheduling data remains local by default.
- `.mpsession` exports may contain student educational information; keep them local or share them only through institution-approved channels.
- Do not add telemetry.
- Do not add analytics.
- Do not add advertising SDKs.
- Do not send student data to AI/LLM services.
- Do not send student data to remote services without an explicitly requested feature.
- Do not introduce cloud storage without explicit instruction.
- Avoid logging personally identifiable student information.
- Use synthetic/anonymized data in tests.

Licensing must remain independent from student scheduling data.


## Tauri Security

Use least privilege.

Do not broadly enable Tauri capabilities.

Do not expose generic shell execution to the frontend.

Validate data crossing the frontend/native boundary.

Prefer narrow, typed native commands.

Do not commit:
- passwords
- tokens
- API keys
- signing keys
- signing certificates
- notarization credentials
- developer-account credentials


## Python Migration

Existing working Python scheduling code must not be rewritten merely for architectural purity.

Before replacing a working optimizer:
1. create regression/equivalence tests
2. verify hard constraints
3. verify relevant objective behavior
4. verify calculations
5. test edge cases

Temporary bundled Python sidecars are acceptable if necessary.

End users must not need Python installed.

Prefer eliminating persistent localhost HTTP/uvicorn architecture from production where practical.

Characterize and test the standalone Python Jury scheduler before integration or rewrite. Do not design the Clinical Placement optimizer until detailed constraints and policies have been gathered from the actual Music Therapy workflow owner.

SQLite schema changes must use the application-owned ordered, forward-only migration registry in `webapp/backend/app/database.py`. The schema version is `PRAGMA user_version`; do not add ad hoc startup column checks or silently run `create_all` against a current database. Migrations must be transactional where SQLite permits, reject newer unsupported schemas, support the recognized legacy Accompanist database, and be covered by synthetic tests. The migration API must accept a staged database engine for future session restore. Never use or commit real student databases as fixtures.


## Development Practices

- Preserve existing functionality unless explicitly changing it.
- Prefer incremental changes over repository-wide rewrites.
- Do not silently remove features.
- Keep UI changes separate from solver changes where practical.
- Keep Tauri command handlers thin.
- Prefer typed models over stringly-typed/unstructured data.
- Return structured errors for recoverable failures.
- Avoid unwrap()/expect() for normal user input/file-processing failures.
- Prefer boring, maintainable solutions over clever ones.
- Do not build speculative future modules before they are needed.


## Imports and Exports

Share parsing, validation, and export infrastructure where useful.

Allow modules to have different input schemas.

Do not create one giant universal spreadsheet format merely to make modules look alike.

Availability is a demonstrated shared concept. A shared `AvailabilityWindow` may represent recurring weekday or dated windows with Available/Tentative/Unavailable status and import provenance, but status meaning/scoring, owner associations, validation policy, and entry UI remain module-specific. Missing availability is not equivalent to Unavailable. Same-status imported overlaps may be merged; conflicting-status overlaps must be surfaced for review. Import Microsoft Forms exports locally from CSV/XLS/XLSX; do not add Forms/Graph API integration. Imported availability is normal editable data, and manual graphical editing must remain available.

Accompanist Availability import is a full Pianist-roster replacement. Within one import, group rows only by Pianist Name after trimming surrounding whitespace and collapsing repeated spaces; do not match against the previous roster by name, email, or any identifier. Preview incoming names and Availability Windows, warn that the roster and Accompanist weekly availability will be replaced, Lesson assignments cleared, and all Jury Availability Windows removed, then apply atomically. Preserve Lessons, Jury Required values, Panels, Panel dates, and Panel selections. Create a fresh internal Pianist/Person identity for each distinct imported name. Pianist IDs and other internal identifiers are never part of user-facing create/edit/import/export workflows. New email defaults blank unless explicitly mapped; Max Hours Per Week defaults to 40.

Share report rendering/export infrastructure, but modules own report definitions, fields, filters, and row semantics. Do not create a universal report designer or permit arbitrary SQL.


## Distribution

The application will eventually be commercially distributed on Windows and macOS.

Maintain compatibility with:
- Windows Microsoft Store distribution
- signed Windows direct distribution
- Apple Developer ID signing
- Apple notarization
- Apple Silicon
- Intel macOS where reasonably practical
- Linux AppImage and deb packaging where reasonably practical

Do not add real production signing credentials to the repository.

Development builds must not require signing credentials.


## Repository Documentation

Important architectural documentation lives in:

- docs/ARCHITECTURE.md
- docs/MODULE-ARCHITECTURE.md
- docs/PRIVACY-ARCHITECTURE.md
- docs/DISTRIBUTION.md
- docs/BUILDING.md

During the Electron-to-Tauri migration also follow:

- docs/TAURI-MIGRATION-PLAN.md
- docs/MIGRATION-ELECTRON-TO-TAURI.md
- docs/PRODUCT-MODULE-ARCHITECTURE.md

Keep documentation synchronized with significant architectural changes.


## Decision Rule

When uncertain whether something belongs in shared infrastructure or a module, default to keeping it inside the module until there is demonstrated reuse.

It is easier to generalize two proven implementations later than to undo a premature abstraction.

Preserve correctness above architectural purity.

## Accompanist -> Jury Dependency

Performance Jury Scheduling is downstream of Accompanist Scheduling.

Where applicable, Jury Scheduling should consume authoritative student,
lesson, and finalized pianist-assignment information produced by the
Accompanist module rather than requiring users to enter or import the same
information again.

Accompanist Scheduling owns accompanist assignments.

Jury Scheduling must not maintain an independent editable copy of
Accompanist-owned assignment data.

The eventual Jury optimizer will combine:

1. typed finalized results from Accompanist Scheduling
2. shared session/person information where appropriate
3. additional Jury-specific inputs

to produce Jury-specific scheduling results.

Do not tightly couple Jury to Accompanist database tables, UI state, or
internal solver objects. Use the approved typed finalized-result boundary.

If a finalized Accompanist result changes after a Jury result has consumed
it, the dependent Jury result must be detectable as stale or superseded
and the user must be warned.

Jury Scheduling must still support participants for whom no Accompanist
assignment exists.

The Accompanist-to-Jury typed result boundary is
`accompanist.assignment-result` contract v2. Contract v1 finalized results
remain historical and immutable.

Jury Required is an Accompanist-owned source Lesson Boolean, independent
from the Accompanist-owned Needs pianist? Boolean. Specific pianist is a
separate optional required-Pianist name. Never infer either Boolean from the
other, from Specific pianist, or from an assignment. The finalized
Accompanist result publishes both Booleans per stable Lesson UUID.
Accompanist Schedule and Jury Lesson Entries both edit the same Jury
Required source field by stable Lesson UUID; do not create a Jury-owned copy.
Jury owns manual Jury Panel selection per source lesson;
turning Jury Required off does not delete its stored Panel selection. Do not
infer panels from instrument, teacher, lesson area, or program.

Jury Setup navigation and keyboard order is Panels, Lesson Entries, Pianist
Availability, Schedule, then Overview. Panels is the initial view. Schedule
is a placeholder until an optimizer is separately authorized.

If Jury Required and Needs pianist? are both Yes for the same source
lesson but no finalized Accompanist pianist assignment exists, treat that
lesson as a blocking Jury-readiness problem. Do not proceed as though a
pianist were not required, and do not substitute another pianist for a
finalized assignment.

Each Jury Panel has its own one-day Jury Date. Jury-day pianist availability
is Jury-specific, binary Available/Unavailable, and keyed by pianist and
date. Tentative is not part of the Jury availability UI or solver semantics.
Available windows are hard legal intervals; all other times in a complete
declaration are unavailable. A declaration is complete only when it contains
one or more valid Available windows. Empty or invalid submissions are
incomplete and must not be interpreted as a complete unavailable schedule.

Any future Jury result must retain enough source-result/session revision
provenance to detect when its consumed finalized Accompanist result changes
and mark the Jury result stale/superseded.

## Scheduling Terminology

Use **Availability Window** as the standard product and architectural term
for a continuous period during which a person or resource is available
for scheduling.

Examples include:
- Pianist Availability Windows
- Student Availability Windows
- Jury-Day Availability Windows
- Faculty Availability Windows
- Resource Availability Windows

An Availability Window represents a continuous interval with a start time
and end time. Module-specific rules determine how that availability is
interpreted.

Use this terminology consistently in:
- user-facing UI
- documentation
- domain/service naming where appropriate
- validation and import messages

Do not unnecessarily replace this established term with alternatives such
as:
- availability range
- availability period
- time range
- available block

The term **Time Slot** has a different meaning and remains valid where the
application actually operates on discrete scheduling units.

Distinction:

- **Availability Window** = a continuous human-facing availability
  interval, such as "9:00 AM–12:00 PM."
- **Time Slot** = a discrete scheduling unit used by an algorithm or
  interface, such as a 30-minute 9:00–9:30 slot.

Availability semantics may differ by module.

For example:
- Accompanist Scheduling may distinguish Available and Tentative
  availability.
- Performance Jury Scheduling uses binary, hard availability: a pianist
  is available within the declared Jury-Day Availability Windows and
  unavailable outside them.

Do not force module-specific availability semantics into the shared
Availability Window concept.