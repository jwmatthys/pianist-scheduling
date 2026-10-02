# Jury Module Data Design

**Status:** Jury data/readiness design implemented. Phase 2 optimizer and Phase 3 generation service, typed result persistence, API, stale-result detection, and read-only Schedule view are documented in [JURY-OPTIMIZER-CONTRACT.md](JURY-OPTIMIZER-CONTRACT.md). Manual schedule editing and finalization remain out of scope.

## 1. Scope and Design Basis

The original Jury data-design milestone established stable identities, an immutable finalized Accompanist result, Jury-owned configuration and availability, dependency provenance, and readiness validation; schedule generation was outside that milestone. Subsequent phases implemented the bounded optimizer, generation service, persistence/API, and read-only Schedule view. The optimizer must terminate with every Jury-required lesson entry either scheduled or explicitly unscheduled with understandable reasons.

Phase 1 established the optimizer contract and DTOs; Phase 2 implemented the pure-domain core; Phase 3 integrates generation through readiness, result-registry persistence, API contracts, and revision-based stale detection. The optimizer itself remains separate from persistence, API, and UI workflows.

The design started from schema version 3; the current schema is version 9. Historical migrations remain unchanged. Accompanist retains numeric `Pianist` and `Lesson` primary keys and current availability behavior while adding internal UUID identity mappings, source revisions, finalized-result history, and an Accompanist-owned per-lesson Jury Required extension. The legacy `pianist_code` column is not part of user-facing identity or import workflows. Jury persistence is module-owned and lesson-based.

## 2. Shared Person Identity

Add a minimal shared `PersonIdentity` keyed by an application-generated UUID. The UUID is the only cross-module identity key. Its small shared fields may include a display name and creation/update timestamps; optional institutional identifiers or email are metadata, never required keys. A display name is presentation data and is never a join key. Do not create a universal student/person record containing Accompanist, Jury, Clinical, and other module fields.

Each module retains its own profile and data around that identity. The existing Accompanist `Pianist` row remains the Accompanist pianist profile and receives a unique `person_uuid` foreign key. An Accompanist student profile stores Accompanist-side display data and the external Student ID metadata; each lesson refers to the student's `person_uuid`. Jury participation is attached to a source lesson UUID and stores that lesson's student UUID, not a student-global participation flag. Jury pianist availability refers to the pianist UUID without duplicating an Accompanist profile.

## 3. Required Identity Migration

The migration must preserve current rows and must not merge people based only on names. Within one Scheduling Session, identical nonblank institutional Student ID values normally identify the same student. Trim surrounding whitespace for ID comparison, but preserve the source value as external metadata; never derive a Person UUID from that value.

- Each existing `Pianist` row can be assigned a generated UUID independently and deterministically retained after migration; copy its current name to the shared identity's display name and keep all existing integer foreign keys and Accompanist behavior intact. Schema v9 historically backfilled the unused legacy `pianist_code` column from the numeric primary key; new records do not generate or expose values in that column.
- Create one UUID-backed Accompanist student profile for each nonblank Student ID within the session and link corresponding lessons to it when the ID group has no conflict signal. This is the normal path and does not require manual reconciliation. Preserve Student ID as external metadata, not as the key.
- Blank-ID lesson rows are not merged by display name; assign separate UUIDs unless a user explicitly reconciles them. For a repeated nonblank ID, compare nonblank student display names after trimming, collapsing whitespace, and case-folding. Formatting-only differences are equivalent. Materially different normalized names flag the ID group for review and block downstream publication until resolved; do not automatically split the group or guess which row is correct. Keep raw source values available for review. This check is a conflict signal, not a claim that a student's name cannot change.
- Add a generated UUID to each lesson source record for stable result provenance. Preserve its existing integer primary key as a module-local implementation detail.

Identity reconciliation must define a safe merge operation for reviewed conflicts: choose a surviving UUID, repoint module references transactionally, and retain an audit/tombstone or equivalent mapping so stale UUIDs cannot silently resolve to a different person. Automatic grouping by display name remains prohibited.

## 4. Finalized Accompanist Result Model

Jury consumes a typed, immutable `AccompanistAssignmentResult` through the result service. It receives a snapshot, not ORM rows, live Accompanist state, solver objects, or a copy of the Accompanist database.

The result envelope contains:

- `session_uuid`, result UUID, module ID, contract ID/version, monotonically increasing Accompanist source revision, state, created timestamp, and finalized timestamp;
- the source revision and input/result provenance used to create the snapshot;
- an immutable payload of lesson-level Accompanist facts.

Each relevant source lesson entry in contract v2 contains `source_lesson_uuid`, `student_person_uuid`, student display name, instrument, teacher, `pianist_required`, `jury_required`, and either no assigned pianist or the finalized assigned pianist's Person UUID and display name. Instrument and teacher are source-lesson facts displayed read-only in Jury. Jury Required is also authoritative Accompanist source data, but both Accompanist Schedule and Jury Lesson Entries edit that same value by stable Lesson UUID; finalized payloads remain immutable. Instrument is not a universal student property. As a Lesson Entries activation convenience only, an unassigned entry may receive a Panel when its full Instrument value exactly matches one Panel Name case-insensitively; partial or ambiguous matches are not applied, and an existing Panel selection is never changed. Preserve one result entry per source lesson; a student with multiple lessons can have different Jury Required values, instruments, teachers, pianist requirements, and assigned pianists on those lessons.

The standalone characterization consumes student name, instrument, pianist-required state, and the fixed pianist assignment. Contract v2 adds authoritative lesson-level Jury Required alongside stable student, lesson, and pianist identity. Teacher is included as read-only result presentation metadata and is not currently a Jury optimizer constraint. Do not aggregate lessons into one student-level Jury flag, pianist, or instrument, and do not join by name. Contract v1 remains immutable and decodable for historical results; new finalizations use v2.

The payload excludes Jury Panel, Jury Date, Jury-Day availability, Jury schedule fields, and other Jury-owned policy. It does not expose max-hours data, Accompanist availability, fit scoring, notes, or unrelated database fields.

## 5. Finalization and Revision Semantics

Finalization means: “This is the Accompanist schedule the human approves for downstream use.” It is an explicit action over committed state. Manual assignment edits and human-approved overrides remain valid inputs. Finalize atomically validates identity resolution and source references, snapshots the current source revision, stores an immutable lesson-level result payload, and marks that result `finalized`. Jury can only resolve a result in this state; it never reads the current editable lessons or pianists as a fallback.

Every persisted Accompanist edit that can affect the result advances its source revision: lesson create/update/delete/import, student identity reconciliation, student/instrument/teacher/requirement edits, pianist profile identity/display changes, pianist assignment edits, and assignment clearing/reruns. A finalized result is not edited in place. A later edit makes it no longer the current finalized result for the current source revision and records it as superseded or otherwise non-current; historical result snapshots remain available for dependent provenance. Non-result-affecting changes need not advance this revision.

Unassigned Accompanist work remains representable: `pianist_required=true` with no assigned pianist is explicit in that lesson's result entry and becomes a Jury readiness blocker only when that lesson is Jury-required. Ordinary solver-quality warnings do not automatically block finalization; finalization records the human-approved state rather than asserting a perfect solver result. Structural ambiguity that prevents a meaningful typed result, including unresolved identity or invalid source references, does block finalization. This distinction does not introduce a broader finalization policy or remove existing validation/reporting.

## 6. Jury Dependency and Staleness

Every generated Jury result records an immutable dependency reference: source session UUID, Accompanist result UUID, result contract version, and Accompanist source revision. It also records the Jury input revision and Jury result revision. The dependency points to the finalized snapshot, not a table row or a name.

When Accompanist source revision advances or its current finalized result changes, the result registry marks dependent Jury results potentially stale (or computes the same condition from source revision references). Keep the old dependency intact; do not silently switch it to a newer Accompanist snapshot or reconcile assignments automatically. The dependency/version architecture is in scope; exact user-facing terminology and refresh workflow are deferred and do not block this data milestone.

## 7. Jury Lesson Participation Data

Jury Required is authoritative Accompanist source lesson data, not a Jury-owned copy. It is editable in both Accompanist Schedule and Jury Lesson Entries through the Accompanist source service. Persist one Jury-owned `jury_lesson_entries` row per source lesson, keyed by `(session_uuid, source_lesson_uuid)`, containing:

- `student_person_uuid`;
- `panel_uuid NULL` until a user selects a panel.

The source lesson UUID and student UUID come from the typed result boundary; the current Jury Required value is read and written through the authoritative Accompanist service. Jury persistence must not store a second requirement Boolean. Jury otherwise edits only the lesson's Panel selection, which remains stored when Jury Required is false and is ignored by readiness in that state. Therefore one student may have a Jury for one lesson but not another, separate panels for separate lessons, and different finalized pianist assignments for those lessons. Jury Required is independent of that lesson's Needs pianist? Boolean and optional Specific pianist name. The only automatic Panel assignment is the explicit Lesson Entries activation convenience above; teacher, Area, program, and Accompanist assignment are never used to infer a Panel. A Jury-required lesson with Needs pianist? = No needs no assigned pianist.

On module activation and after source replacement/deletion operations, Jury reconciliation consumes an Accompanist-owned UUID identity snapshot. It removes lesson-entry rows whose source Lesson UUID no longer exists, updates the projected student UUID for surviving source lessons without changing their Panel selection, and removes availability declarations/windows whose Pianist UUID is no longer in the Accompanist roster. If a current finalized result exists, the roster is also synchronized from that immutable result. Reconciliation returns aggregate counts and detects stale dependent results through source result/revision comparison; structured logs contain counts and event codes only, never student or pianist names.

The initial Jury participation roster is built from lesson entries in the current finalized result. Every Jury entry must reference an authoritative source Lesson UUID; Jury-only participants are out of scope and would require a future explicit design.

## 8. Jury Panel Model

Persist `jury_panels` with UUID primary key and session UUID, plus:

- nonblank panel name and descriptive room;
- earliest start time and preferred start time as local minutes since midnight;
- positive jury length in integer minutes;
- `break_needed` Boolean;
- nullable `break_every_x_juries` and `break_length_minutes`, required only when breaks are enabled;
- `meal_break` Boolean;
- nullable `meal_start_minute` and `meal_end_minute`, required only when meal break is enabled.

No latest start, panel end time, or room-occupancy constraint is stored. Room is descriptive only: panels sharing a room do not create warnings or constraints. Panel UUID, not name or room, is the stable reference.

Validation: name must be nonblank; times must be integer minutes within the Jury day; earliest start must be before day end; preferred start must be within the day and not earlier than earliest start; jury length must be positive and fit at least one interval within the day. If breaks are enabled, `break_every_x_juries` must be a positive integer and break length must be positive. Enabling Break Needed initializes Break Length to the current Jury Length; after initialization it is an independent stored/editable value. Changing Jury Length later must not change the stored Break Length. If no break is needed, periodic-break parameters are null/ignored, not accidentally applied. If meal break is enabled, both endpoints are required and `0 <= start < end <= 1440`; otherwise both are null/ignored. The meal interval resets the periodic-break counter. An explicit 60-minute jury length is valid data and must not depend on legacy arithmetic such as `(60 - 10) // duration`.

Whether preferred-before-earliest should be rejected or normalized was open in the request. This design chooses a blocking panel validation error; silently normalizing would hide a contradictory saved configuration. No optimizer objective or time placement is specified here.

## 9. Jury-Day Pianist Availability

Jury availability belongs to Jury and is keyed by the Accompanist-provided `PersonIdentity` UUID. Do not create a second Jury pianist profile. Persist a per-session, per-pianist, per-date internal availability record with derived completeness, timestamps, and positive Available windows. The user-facing feature is named **Jury Availability Windows**; do not expose declaration or completeness terminology in the ordinary UI. Windows are local half-open intervals `[start, end)` within the applicable Panel date; the UI presents binary Available/Unavailable semantics but persistence need not store Unavailable windows or status rows. There is no Tentative state.

Closed-world behavior is explicit: positive `Available` intervals define legal time; every other relevant time in a complete record is derived as `Unavailable`. An internal record is complete when it contains at least one valid Availability Window. No windows means incomplete/unknown and blocks readiness when the pianist is required; it is not treated as a complete unavailable day. Reject out-of-day, zero/negative, or overlapping-invalid intervals and normalize overlapping/adjacent Available windows to their union. Records remain keyed by date so a pianist may have different Availability Windows on dates used by different Panels. Accompanist roster replacement removes all Jury availability records and windows for the active session in the same transaction; it never transfers availability to replacement people by name.

The only Jury Availability Windows required for readiness are those of pianists referenced by a Jury-required lesson entry's finalized assignment where that lesson's Needs pianist? Boolean is Yes; readiness checks the date of the Panel assigned to that lesson. Missing availability issues should say that the named Pianist has no Jury Availability Windows for the relevant date.

## 10. Panel Date

Each Panel has exactly one required `jury_date`; different Panels may occur on different dates. No Panel spans more than one day. A missing date blocks readiness. This supports multiple one-day Panels without introducing multi-day Panel schedules, timezone handling, or cross-panel booking logic. The legacy `jury_configurations.jury_date` column is retained only as the schema-v8 migration source for backfilling existing Panels and is no longer an active setting.

## 11. Readiness Validation

Provide a pure/module-domain readiness service that accepts a consistent snapshot of Jury configuration, lesson-based Jury entries, finalized Accompanist result, and Jury availability declarations. It returns typed issues (`code`, `severity`, message, and relevant entity UUIDs); it does not call an optimizer. Do not pass structurally incomplete input to the future optimizer.

Blocking errors include:

- no current finalized Accompanist result or a result from a different session/unsupported contract;
- unresolved or ambiguous lesson/student identity/source projection;
- a Jury-required lesson entry with no selected existing panel;
- `Jury Required` plus Needs pianist? = Yes on the same lesson entry with no finalized assigned pianist;
- a referenced pianist identity missing from the source result or with no Jury Availability Windows for the assigned Panel's date;
- a Jury-required lesson's selected Panel has no Jury Date;
- invalid panel fields, break parameters, meal interval, or preferred start earlier than earliest start;
- an invalid or conflicting source identity that prevents an authoritative lesson-level result projection.

Warnings may include a finalized Accompanist result whose source revision has changed (stale dependency), no Jury-required lesson entries, Jury-required lesson entries not assigned to a pianist because pianist is not required for that lesson, panels with no Jury-required entries, or descriptive duplicate room names. Same-room panels specifically produce no warning. Warnings do not make missing/invalid data schedulable. Readiness reports must state which lesson/person/panel/pianist is affected without leaking unrelated student details into logs.

## 12. Proposed Database Schema

Names are conceptual; implementation should follow the existing SQLAlchemy and explicit migration conventions.

| Table | Core columns and constraints |
| --- | --- |
| `person_identities` | `person_uuid UUID PK`, `display_name`, optional institutional ID/email metadata, timestamps; no module behavior fields. |
| `accompanist_students` | `person_uuid UUID PK/FK`, session-scoped external Student ID metadata and Accompanist-specific reconciliation fields; unique `(session_uuid, student_id)` for nonblank IDs, with conflicting source-name evidence flagged before publication. |
| `pianists` | Existing columns retained, including the legacy `pianist_code` compatibility column, which is not used or exposed by normal workflows; application identity is provided through the internal `AccompanistPianistIdentity` to `PersonIdentity` mapping. Schema v9 historically backfilled legacy code values from the prior numeric primary key. |
| `lessons` | Existing columns retained; add stable `lesson_uuid UUID UNIQUE NOT NULL` and `student_person_uuid UUID FK person_identities`. Do not remove integer IDs or raw `student_id` during this milestone. |
| `accompanist_lesson_jury_requirements` | `lesson_id` PK scoped to the active database, `jury_required BOOLEAN NOT NULL DEFAULT false`; Accompanist-owned source lesson extension added in v7. |
| `module_results` | Result UUID PK, session UUID, module ID, contract ID/version, source revision, state, created/finalized timestamps, typed payload JSON/text and payload schema version; index current results by session/module/state. Enforce immutable finalized payloads in the result service. |
| `module_result_dependencies` | Dependent result UUID, source result UUID, source session UUID, source contract/version, source revision; typed registry relation, not a polymorphic SQL foreign key to arbitrary module tables. |
| `accompanist_state` (or equivalent revision row) | Session UUID PK, current source revision and current finalized result UUID; incremented transactionally for result-affecting Accompanist edits. |
| `jury_configuration` | Session UUID PK/FK, Jury input revision, timestamps. The previous session-wide date remains only for migration compatibility. |
| `jury_panels` | Panel UUID PK, session UUID, all panel fields in section 8 except date; index by session/name. |
| `jury_panel_dates` | Panel UUID PK/FK, session UUID, one required Jury Date; separate extension preserves the released v6 Panel table shape. |
| `jury_lesson_entries` | Composite PK `(session_uuid, source_lesson_uuid)`, student Person UUID, nullable Panel UUID; no authoritative Jury Required Boolean. `source_lesson_uuid` resolves through the finalized result contract and is not a direct Accompanist-table dependency. The unused v6 Boolean is legacy-only and not read or written by current services. |
| `jury_pianist_availability_declarations` | Composite key `(session_uuid, pianist_person_uuid, jury_date)`, derived complete Boolean, timestamps; FK to shared identity/session. Complete exactly when at least one valid Available window exists. |
| `jury_pianist_available_windows` | Window UUID PK, declaration key including Jury Date, start/end minute representing only positive Available intervals; check `0 <= start < end <= 1440`. Other time is derived unavailable; no status field or Tentative value. |

Use JSON only for versioned module-owned result payloads if that matches repository conventions; indexed lifecycle/provenance fields remain relational. Do not introduce a generic `owner_type + owner_id` relation for people or availability. Enforce one active session as the current product does, while retaining session UUID in module-owned records and results for future session isolation.

## 13. Proposed Service and API Boundaries

Keep FastAPI routes thin and define service interfaces around domain operations:

- Accompanist service: edit/persist existing Accompanist data, advance source revision, validate projection, and `finalize_accompanist_result()`; its implementation may read Accompanist tables.
- Result registry service: list/resolve only finalized typed contract versions and record/mark dependencies; it owns lifecycle/provenance, not optimizer logic.
- Jury input service: create/update Panels and their dates, lesson Panel assignments, and pianist declarations using UUID identities; it does not query Accompanist tables.
- Jury readiness service: receive typed finalized-result data plus Jury-owned DTOs and return structured issues.
- Jury synchronization service: compare Jury UUID references against an Accompanist-owned source identity snapshot, prune only orphans, preserve surviving Panel choices, and return PII-free aggregate counts plus stale-result counts.
- Jury generation service: resolve a current finalized typed result, require readiness, build optimizer DTOs from the typed boundary and Jury-owned input records, invoke the pure-domain optimizer, then write a versioned immutable payload and dependency through the module result registry. The optimizer receives no ORM, UI, Tauri, HTTP, or persistence objects.

Local Jury schedule API operations are `POST /api/jury/generate`, `GET /api/jury/results/current`, `GET /api/jury/results/history`, and `GET /api/jury/results/{result_id}`. Generation returns structured readiness blockers and never invokes the optimizer when readiness fails. Jury services resolve source data through the Accompanist result service, never direct SQL joins to `lessons`/`pianists`. Keep the application service transport-independent so a future browser/local-storage host can supply an equivalent adapter.

## 14. `.mpsession` Implications

No archive format change is required. Format 1 already packages a SQLite snapshot, records `PRAGMA user_version`, and restores by migrating a staged database before activation. The new identities, immutable results, Jury configuration, and Jury availability live in that same database payload and therefore round-trip together through Export, New Session, and Open/Restore.

Add forward-only database migrations after v3; do not edit historical migrations. Keep the v1 manifest/archive contract and existing checksums. Ensure a fresh session initializes the latest schema; old format-1 archives migrate on staged restore, while export reports the latest schema number in the existing manifest. New-session replacement naturally starts with empty Jury data. No duplicated Jury JSON archive entry or archive-version bump is proposed.

## 15. Privacy and Local Data

Person names, student identifiers, lesson facts, assignments, panels, and availability are student/staff scheduling data. Persist and process them locally; do not add telemetry, analytics, cloud storage, remote identity resolution, or AI/LLM processing. Do not log names, external IDs, or availability details. UUIDs are internal identifiers, not anonymization. `.mpsession` remains unencrypted in v1 and may contain educational information; preserve local recovery and institution-approved sharing practices.

## 16. Future Browser Compatibility

The contract is based on UUIDs, typed DTOs, versioned local result payloads, explicit validation, and local session persistence. Keep result projection, revision logic, readiness, and Jury data operations independent from Tauri and HTTP. Browser implementation may use local SQLite/WASM or another local store behind the same interfaces; importing, finalizing, validating, and later optimizing must not require a vendor server or upload student data. Do not implement WASM or cloud services in this milestone.

## 17. Implementation Sequence for This Data Milestone

1. Define conflict detection for repeated Student IDs and write synthetic migration cases before changing production schema.
2. Add schema v4 for shared UUID identities, Accompanist student profile links, pianist identity links, lesson UUIDs, and source revision tracking. Backfill pianist UUIDs; group identical nonblank Student IDs within the session; flag clearly conflicting uses for review; keep blank-ID rows distinct regardless of name.
3. Add schema v5 for typed result records/dependencies and Accompanist finalization/revision service. Keep current manual-edit operations; ensure every relevant edit advances revision and finalized snapshots are immutable.
4. Add schema v6 for Jury configuration/date, lesson-based Jury entries, panels, and complete binary pianist availability.
5. Add schema v7 for Accompanist-owned lesson Jury Required, defaulting existing lessons false; do not edit v1-v6 migrations.
6. Publish immutable Accompanist result contract v2 with lesson Jury Required, retaining a v1 decoder for historical results.
7. Keep Jury entry persistence Panel-only; both modules edit the same Accompanist Jury Required source field by Lesson UUID, and stale finalized snapshots block readiness until refreshed.
8. Verify session archive export/new/restore round trips and staged migration from v3 and older supported versions. Keep `.mpsession` format 1.
9. Record the Phase 1 optimizer contract and ownership boundary before separately authorizing optimizer implementation/integration; see [JURY-OPTIMIZER-CONTRACT.md](JURY-OPTIMIZER-CONTRACT.md). The contract does not authorize or contain scheduling logic.

Migration grouping may change during implementation, but versions remain explicit, ordered, transactional where SQLite permits, forward-only, and tested using synthetic fixtures. Do not change v1-v3 definitions.

## 18. Tests Required Before Optimization

- Migration from schema v6 with existing lessons and v1 results; assert v7 defaults Jury Required false, preserves lesson data, supersedes the current v1 pointer without deleting history, and still decodes the old typed result.
- Migration from schema v3 with multiple lessons, duplicate display names, blank/repeated external student IDs, existing manual assignments, pianist availability/completeness, and session metadata; assert identical nonblank IDs normally share one UUID, blank-ID rows remain distinct despite matching names, conflicting ID use is flagged, and no Accompanist data is lost.
- Fresh and legacy-v0 migrations through the new latest schema; transactional failure rollback; rejection of newer unsupported schema; no changes to v1-v3 migration definitions.
- UUID stability through reopen, new session, archive export/restore, and staged archive migration; all identity/result/Jury rows retained in a format-1 round trip.
- Identity tests prove internal UUIDs are generated independently of Student ID; ordinary identical nonblank IDs need no manual reconciliation; clearly conflicting use blocks downstream publication pending review; blank IDs never merge on matching names.
- Finalization produces immutable typed snapshots; manual Accompanist edits remain supported; relevant edits advance revision; non-relevant edits do not; old result remains addressable and is marked non-current after edits.
- Result contract v2 tests preserve lesson UUID, student UUID/name, instrument, teacher, both independent requirement Booleans, and that lesson's assigned-pianist UUID/name without aggregation across lessons; contract v1 history remains interpretable.
- Jury lesson entries persist Panel only; editable Jury Required writes the Accompanist source field, and refresh keeps the Panel choice even when the flag is false. A stale finalized snapshot blocks readiness.
- Lesson import tests cover yes/no, y/n, true/false, 1/0, case/whitespace variants, blank false defaults, old files without the column, and invalid nonblank value issues.
- Panel validation boundaries, including 60-minute jury duration, default break length, positive break interval/count, meal endpoint pairing/order/day bounds, meal reset configuration, preferred-before-earliest, and descriptive room duplicates without warnings.
- Availability accepts only Available/Unavailable, rejects Tentative, uses complete closed-world semantics with completeness derived from one or more Available windows, and treats missing/incomplete submissions as blocking unknown.
- Readiness emits deterministic structured blockers/warnings, catches missing fixed pianist and stale source revision, and never invokes any optimizer.
- Local API/service tests prove Jury reads only finalized typed result contracts, not Accompanist ORM/SQLAlchemy state, React state, or solver internals.
- No test in this milestone calls `build_schedule()` or asserts an optimizer implementation; future optimizer tests must separately prove finite termination and explicit unscheduled reasons for impossible availability, as well as arithmetic safety for 60-minute duration and configurable breaks.

Use synthetic/anonymized data only.

## 19. Unresolved Product Decisions
- Jury Availability Windows remain stored by date; changing a Panel date does not transfer or delete windows for its previous date. Clearing one Pianist/date does not affect other dates; replacing the Accompanist Pianist roster clears Jury availability for all dates in the active session.
- Confirm whether Earliest Start and Preferred Start are both required, and accept the proposed blocking validation when Preferred Start is earlier.
- What is the acceptable migration reconciliation UX and behavior when an identity is merged after Jury settings or finalized results exist? Preserve UUID lineage and mark affected results stale; do not silently rewrite historical result payloads.
- Exact stale-result terminology and refresh workflow are deferred; they are not prerequisites for this data milestone.

## 20. Concepts That Remain Module-Specific

- Accompanist owns lesson facts, the Needs pianist? Boolean, optional Specific pianist name, Jury Required, pianist profiles and assignments, Accompanist availability/workload/scoring, source revisions, and the finalized assignment-result projection.
- Jury owns manual Panel selection, each Panel's one-day Jury Date, panel timing/duration/break/meal configuration, binary date-keyed pianist Availability Windows, Jury readiness policy, and future Jury schedules/results. Jury Required remains Accompanist-owned source data editable in both modules.
- Clinical Placement owns its placement requirements, sites, capacities, transport/history rules, optimizer, and result payload.
- Shared infrastructure owns only stable person UUIDs, session metadata, result envelopes/dependency mechanics, time primitives, migration/session archive mechanics, and structured validation issue representation.

There is no universal Person profile, Student profile, Assignment, Schedule, Availability policy, or optimizer. Jury consumes Accompanist's typed finalized contract; it does not own or edit Accompanist assignment data.

## Explicitly Out of Scope

The approved integration does not call or port `build_schedule()`, implement manual schedule reordering, add cross-panel room booking, or design multi-day Juries. The Schedule view is read-only; manual schedule edits and finalization remain deferred. The implemented optimizer separately guarantees bounded termination and an explicit scheduled/unscheduled outcome for every Jury-required lesson.