# Music Program Scheduler
# Electron-to-Tauri Migration and Architecture Plan

## PROJECT VISION

This project is evolving from a single-purpose accompanist scheduling application into a modular suite of scheduling tools designed specifically for university and conservatory music programs.

The existing application schedules pianists to student lessons according to pianist availability, lesson times, workload, and schedule-block consolidation. The current accompanist optimizer does not model actual room/site distance or travel time.

This will become the first scheduling module in a larger product.

Known future modules include:

1. Accompanist Scheduling
   - Assign pianists to student lessons.
   - Respect pianist availability.
   - Avoid conflicts.
   - Minimize schedule fragmentation and, when location-aware scheduling is implemented, unnecessary travel.
   - Prefer contiguous blocks.
   - Balance workload.
   - Track overlapping assignments correctly.

2. Music Therapy Clinical Placements
   - Assign every music therapy student to an appropriate clinical site.
   - Consider student availability.
   - Consider clinical-site availability.
   - Consider population type and placement requirements.
   - Consider prior placement history.
   - Consider transportation requirements.
   - Consider whether the student has access to a car.
   - Consider travel feasibility and travel burden.
   - Optimize the overall quality/equity of placements.

3. Performance Jury Scheduling
   - Schedule students into end-of-semester performance jury time slots.
   - Incorporate data produced by accompanist scheduling where relevant.
   - Consider pianist assignments and availability.
   - Consider students, faculty, rooms, panels, times, durations, and other jury constraints.
   - Preserve and eventually migrate the functionality of the existing Python jury-scheduling program.

Additional university music scheduling modules may be developed later.

Examples might eventually include:
- auditions
- recital scheduling
- ensemble placement
- chamber music
- studio classes
- room scheduling
- masterclasses
- festival scheduling

Do NOT implement speculative future modules now.

Architect the application so new scheduling modules can be added cleanly without restructuring the entire product.

The approved post-shell product/session architecture is detailed in [PRODUCT-MODULE-ARCHITECTURE.md](PRODUCT-MODULE-ARCHITECTURE.md), with concise module ownership rules in [MODULE-ARCHITECTURE.md](MODULE-ARCHITECTURE.md). Follow those documents for session metadata, availability, reports, and cross-module results.


## FUNDAMENTAL ARCHITECTURAL PRINCIPLE

GENERALIZE SHARED DOMAIN CONCEPTS AND INFRASTRUCTURE.

DO NOT FORCE DIFFERENT OPTIMIZATION PROBLEMS INTO ONE UNIVERSAL ALGORITHM.

Shared concepts may include:

- academic terms
- people
- students
- faculty
- staff
- pianists
- locations
- rooms
- off-campus sites
- dates
- days
- times
- time ranges
- availability
- travel
- assignments
- preferences
- constraints
- imports
- exports
- project persistence
- validation
- schedule visualization

But each scheduling module should own its own:

- module-specific inputs
- module-specific data
- constraints
- objective/scoring model
- optimization algorithm
- validation
- workflow
- output representation
- reports

Conceptually:

                    Music Program Scheduler
                            |
                    Common Domain Layer
                            |
       +--------------------+--------------------+
       |                    |                    |
 Accompanist Module    Clinical Module       Jury Module
       |                    |                    |
 Accompanist Solver    Clinical Solver        Jury Solver
       |                    |                    |
       +--------------------+--------------------+
                            |
                    Shared Infrastructure
                            |
                  Platform Abstraction
                     /            \
                Tauri/Desktop    Browser/Web


## PRODUCT ORGANIZATION

Do not model major scheduling modules as simple tabs in one giant scheduler.

Use a product-level dashboard and module-based workflow. One Scheduling Session represents one institution/program and one academic term; “Project” is only a user-facing synonym. The initial product has one active session/database at a time, not an in-app session history library.

Conceptually:

Music Program Scheduler
         |
         +-- Open/Restore or New Session
                 |
                 +-- Institution / Program / Academic Term
                         |
                         +-- Session Dashboard
                                 +-- Accompanist Scheduling
                                 +-- Clinical Placements (when implemented)
                                 +-- Performance Juries (when implemented)

Each module may have its own internal tabs/views.

For example:

Accompanist Scheduling:
- Students/Lessons
- Pianists
- Availability
- Assignments
- Schedule
- Reports

Clinical Placements:
- Students
- Clinical Sites
- Requirements
- Availability
- Placements
- Reports

Performance Juries:
- Students
- Jury Requirements
- Faculty/Panels
- Rooms/Times
- Schedule
- Reports

Do not assume these exact tabs are final.

Base actual UI decisions on workflow usability.


==================================================
1. CURRENT MIGRATION GOAL
==================================================

Migrate the current Electron application to Tauri 2 while preserving:

- existing UI
- scheduling behavior
- data formats
- import/export functionality
- manual adjustment functionality
- current accompanist scheduling functionality

At the same time, restructure the project so the existing accompanist scheduler becomes a MODULE within the larger Music Program Scheduler architecture.

Do not build the Clinical Placement or Jury modules as part of the Electron-to-Tauri migration unless code already exists that can be integrated safely.

The immediate objective is:

1. establish the shared architecture
2. migrate the existing accompanist module into it
3. establish clean extension points for future modules

Do not let future extensibility prevent completion of the current accompanist scheduler.


==================================================
2. EXISTING PROJECT ANALYSIS
==================================================

Before changing application code, inspect the entire repository.

Inspect:

- package.json
- Electron main process
- preload code
- Electron IPC
- frontend
- React components
- TypeScript/JavaScript
- Python code
- FastAPI
- uvicorn
- scheduling/optimization code
- Excel handling
- CSV handling
- persistence
- saved projects
- build scripts
- tests
- sample data
- documentation

Document the existing architecture.

Determine:

- what runs in Electron
- what runs in the frontend
- what runs in Python
- what currently requires HTTP
- how scheduling data flows through the application
- how imports and exports work
- how manual assignments work
- what parts are truly accompanist-specific
- what concepts can reasonably become shared domain concepts

Do not begin with a mechanical Electron-to-Tauri translation.


==================================================
3. TARGET TECHNOLOGY STACK
==================================================

Use:

- Tauri 2
- existing React frontend where practical
- TypeScript where practical
- Rust for Tauri/native functionality
- npm unless the existing project deliberately uses something else
- current stable Tauri APIs/plugins

The installed desktop application must not require:

- Python installation
- Node installation
- Rust installation
- uvicorn
- an external web server
- developer tools

Temporary bundled sidecars are acceptable during migration if necessary.

Ordinary end users must receive a self-contained application.


==================================================
4. PRODUCT DOMAIN MODEL
==================================================

Introduce a product-level model appropriate to university music programs.

Conceptually:

MusicProgram
    |
    +-- AcademicTerms
    |
    +-- People
    |     +-- Students
    |     +-- Faculty
    |     +-- Staff
    |     +-- Pianists
    |
    +-- Locations
    |     +-- Buildings
    |     +-- Rooms
    |     +-- OffCampusSites
    |
    +-- SharedSchedulingData
    |
    +-- ModuleData
          +-- Accompanist
          +-- ClinicalPlacements
          +-- Juries

Do NOT build an enormous universal Person or Student object containing every field every future module might possibly need.

Prefer:

shared person/student identity
        +
module-specific records

For example:

Student
- internal application-generated UUID
- name
- optional institutional identifier/email
- shared/basic identity properties

AccompanistStudentData
- lesson
- instrument
- accompanist requirement
- assigned pianist

ClinicalPlacementStudentData
- placement requirements
- placement history
- transportation/car access
- clinical availability/preferences

JuryStudentData
- jury duration
- panel requirements
- repertoire/instrument requirements
- accompanist relationship
- jury scheduling requirements

Avoid speculative fields.

Only move information into the shared model when it is genuinely shared.


==================================================
5. ACADEMIC TERMS
==================================================

Use Academic Term as an important organizing concept.

Examples:

Fall 2026
Spring 2027
Summer 2027

Each Scheduling Session represents one institution/program and one academic term; “Project” is a user-facing synonym only. Session metadata is an internal UUID, institution display name, program display name, term display label, optional year/start/end dates, and created/modified timestamps. Exact dates and timezone are not required initially. Keep one active session/database at a time; do not add an in-app historical session library.

Conceptually:

Fall 2026
    |
    +-- Accompanist Assignments
    +-- Clinical Placements
    +-- Jury Schedule

Module data should be capable of referencing shared term-level data.


==================================================
6. CROSS-MODULE DEPENDENCIES
==================================================

Modules must remain independent but should be capable of consuming finalized outputs from other modules.

Known example:

Accompanist Scheduling
        |
        v
Performance Jury Scheduling

The jury scheduling module may need pianist assignments generated by the accompanist module.

Do not require:

export spreadsheet
        ->
reimport spreadsheet

when both modules live inside the same project and can safely share structured data.

Establish a clean mechanism for module outputs to be consumed by another module. Use a session-level versioned result envelope with draft/finalized/superseded states. If a consumed finalized result changes, mark dependent results as potentially stale and surface that condition to the user. Never merge people solely because their display names match.

Avoid direct module-to-module implementation dependencies where practical.

Prefer something conceptually like:

AccompanistModule
      |
published/finalized assignments
      |
Term Data / Module Result API
      |
JuryModule

This allows modules to evolve independently.


==================================================
7. MODULE INTERFACE
==================================================

Create a lightweight architectural concept for scheduling modules.

Do not over-engineer this into an elaborate plugin platform.

A module should conceptually be able to define:

- identity
- display name
- routes/navigation
- input model
- persisted data
- validation
- optimization operation
- results
- reports
- dependencies on other module results

The main application should not contain large amounts of module-specific conditional logic such as:

if module == "accompanist"
else if module == "jury"
else if module == "clinical"

Prefer module boundaries that allow functionality to remain self-contained.

However, do not create an unnecessary dynamic plugin loader.

These modules will initially be built into the application.


==================================================
8. OPTIMIZATION ARCHITECTURE
==================================================

Do NOT create one universal optimizer for all scheduling problems.

Use separate optimization engines.

Conceptually:

scheduling/
    common/
    accompanist/
    clinical/
    jury/

Common code may provide reusable primitives such as:

- TimeRange
- Availability
- conflicts
- Location
- travel time
- Assignment
- constraint results
- scoring helpers
- solver utilities

Availability is a demonstrated shared concept. Use a normalized `AvailabilityWindow` with Available/Tentative/Unavailable vocabulary while keeping owner associations, status meaning/scoring, validation, and UI module-specific. Track submission completeness separately from status: absent windows are unknown for an incomplete horizon and derive Unavailable for a valid complete horizon. Same-status overlaps may be merged; conflicting-status overlaps require review. Import Microsoft Forms exports locally as CSV/XLS/XLSX; do not integrate Forms/Graph APIs.

But:

AccompanistOptimizer
ClinicalPlacementOptimizer
JuryOptimizer

must be allowed to implement fundamentally different models and algorithms.

The correct algorithm for one module must not be distorted merely to make it fit a generalized framework.


==================================================
9. ACCOMPANIST MODULE
==================================================

The existing application becomes the first production module.

Preserve current functionality.

The module should own concepts specific to accompanist assignment, such as:

- lessons
- pianist requirements
- pianist availability
- pianist load
- assignment preferences
- schedule blocks
- overlapping lessons
- merged working-time calculations
- accompanist-specific objective weights

Do not leak these concepts unnecessarily into the product-wide shared domain model.


==================================================
10. FUTURE CLINICAL PLACEMENT MODULE
==================================================

Do not implement this module during the Tauri shell migration. Afterward, build its data-entry/import/manual-edit workflow only from requirements gathered from the actual Music Therapy workflow owner. Do not design or implement its optimizer until those constraints and policies are understood.

Architectural requirements should support a future module that matches students to clinical placements based on factors including:

- student availability
- site availability
- site capacity
- population type
- placement requirements
- previous placements
- student experience/history
- travel requirements
- travel time
- access to a car
- transportation feasibility
- preferences
- equity/fairness

This is not a conventional time-slot scheduling problem and must not be forced into the accompanist algorithm.

Conceptually:

Students + Clinical Sites + Constraints
                |
        Clinical Optimizer
                |
         Student -> Site


==================================================
11. FUTURE JURY MODULE
==================================================

The standalone Python Jury scheduler is characterized in [JURY-SCHEDULER-CHARACTERIZATION.md](JURY-SCHEDULER-CHARACTERIZATION.md). Jury Integration Milestone A now provides identities, finalized Accompanist result contract v2, Jury-owned setup persistence, and readiness. Do not call, port, or rewrite the legacy scheduler until a separately approved optimizer milestone.

Preserve awareness that a working Python jury scheduler already exists.

Eventually:

- inspect that program separately
- identify its inputs
- preserve its tested behavior
- migrate/integrate it carefully
- use accompanist module output directly where appropriate

Conceptually:

Students
+ Faculty
+ Pianist assignments
+ Jury availability
+ Rooms
+ Time slots
+ Jury constraints
        |
    Jury Optimizer
        |
Student -> Jury time


==================================================
12. FRONTEND ARCHITECTURE
==================================================

The frontend should support:

Application Shell
    |
    +-- Dashboard
        +-- Open/Restore or New Scheduling Session
        +-- Session Dashboard (institution/program and academic term)
    +-- Modules
          |
          +-- Accompanists
          +-- Clinical Placements
          +-- Juries

Do not hard-code the overall application around accompanist terminology.

Accompanist terminology belongs inside the accompanist module.

The product shell should use general music-program terminology.


==================================================
13. FUTURE WEB / SAAS COMPATIBILITY
==================================================

The product may later be offered as a web/SaaS application.

The architecture must allow the React frontend, domain model, scheduling modules, imports/exports, and application logic to be reused in:

1. Tauri desktop
2. browser-hosted web application

The future web application must be capable of operating while student scheduling data stays on the user's device.

A server may deliver:

- frontend code
- static assets
- updates
- product information
- documentation
- licensing/account information

Core student scheduling data must not need to be uploaded.


==================================================
14. LOCAL-FIRST ARCHITECTURE
==================================================

Treat these as separate decisions:

Desktop vs Web
Local vs Cloud

Do not assume SaaS means cloud-stored student data.

Core scheduling should operate locally.

Optional future services may include:

- licensing
- cloud sync
- backup
- collaboration
- multi-device access

These should be optional layers.

The core scheduler should not depend on them.


==================================================
15. PLATFORM ABSTRACTION
==================================================

Do not scatter Tauri APIs throughout React components.

Create explicit platform/service abstractions.

Potential interfaces include:

- FileService
- SessionStorage
- SettingsService
- PlatformService
- UpdateService
- LicensingService

Tauri implements these using native capabilities.

A future browser application can implement them using browser capabilities.

React/domain code should normally depend on interfaces, not Tauri directly.


==================================================
16. SCHEDULING CORE AND WEBASSEMBLY
==================================================

If schedulers are implemented in Rust, keep optimization/domain crates independent of Tauri.

Example:

crates/
    scheduling-common/
    accompanist-solver/
    clinical-solver/
    jury-solver/

These names are illustrative.

A future browser version may potentially compile appropriate solver crates to WebAssembly.

Do not implement WebAssembly now merely for architectural purity.

Maintain reasonable compatibility where practical.


==================================================
17. PYTHON / UVICORN MIGRATION
==================================================

The existing prototype uses Python/FastAPI/uvicorn.

Inspect exactly what the backend does.

Separate:

HTTP transport
from
domain/scheduling logic

Preferred final desktop architecture:

React
   |
Application Layer
   |
Module
   |
Scheduling Solver
   |
Local Data

Eliminate the localhost HTTP server eventually where practical.

Do NOT rewrite proven Python optimization code prematurely.

Migration strategy:

1. inventory Python endpoints
2. identify pure scheduling/domain logic
3. migrate simple native operations to Rust
4. preserve reliable complex Python logic temporarily if required
5. use a bundled sidecar if necessary
6. require no Python installation by users
7. eliminate unnecessary HTTP
8. create regression tests
9. migrate solver functionality only when equivalent behavior can be verified

Never sacrifice correct scheduling for architectural purity.


==================================================
18. PRIVACY AND SECURITY
==================================================

The fundamental privacy principle is:

"The application provider should not need possession of student scheduling data in order to provide the scheduling service."

Therefore:

- core scheduling works offline
- project data remains local by default
- no telemetry
- no analytics
- no advertising SDKs
- no student data sent to AI/LLM services
- no unnecessary network communication
- least-privilege Tauri permissions
- no arbitrary frontend shell execution
- validate all frontend/native inputs
- do not log personally identifiable student information when avoidable
- use anonymized/synthetic test data

Portable `.mpsession` archives are unencrypted in v1 and may contain student educational information. Keep them local or share them only through institution-approved channels. Before Open/Restore replaces the active session, create a local recovery snapshot and ask for user confirmation.


==================================================
19. APPROVED POST-SHELL IMPLEMENTATION ORDER
==================================================

After the current Tauri shell validation milestone, proceed in this order:

1. Establish an application-owned, forward-only SQLite migration registry and synthetic legacy fixture. Store the version in `PRAGMA user_version`; recognize an empty v0 database and the supported unversioned Accompanist schema as v0, migrate sequentially to the latest version, and reject unknown/newer schemas safely. Run schema DDL and version updates transactionally where SQLite permits. See [ARCHITECTURE.md](ARCHITECTURE.md#sqlite-schema-migrations).
2. Implement session metadata, New Session, `.mpsession` export, and Open/Restore with staged validation/migration and an automatic local recovery snapshot before confirmed replacement. Keep one active session/database; do not add a history library.
3. Introduce shared availability value/import infrastructure with Accompanist Scheduling as the first consumer. Preserve the current manual pianist editor and solver semantics; track complete-submission state separately, treat valid blank periods as Unavailable only within a complete submission, reject malformed imports without mutation, and flag conflicting-status overlaps.
4. Introduce shared report infrastructure by adapting existing Accompanist reports without changing their established semantics. Do not create an arbitrary query designer.
5. Introduce the product module registry/dashboard once the session and shared infrastructure are useful. Do not add placeholder module workflows.
6. Independently characterize and test the standalone Jury scheduler. Complete: see [JURY-SCHEDULER-CHARACTERIZATION.md](JURY-SCHEDULER-CHARACTERIZATION.md).
7. Establish the Jury identity/result/input/readiness boundary. Complete through schema v9 and contract v2; Jury Setup provides Panels, Lesson Entries, date-keyed Pianist Availability, and stale-readiness handling. Optimizer integration and schedule generation remain later work.
8. Build the Clinical Placement data-entry/import/manual-edit workflow.
9. Design the Clinical Placement optimizer only after requirements and policies are gathered from the Music Therapy workflow owner.

The detailed approved design is in `docs/PRODUCT-MODULE-ARCHITECTURE.md`; module ownership rules are in `docs/MODULE-ARCHITECTURE.md`.
