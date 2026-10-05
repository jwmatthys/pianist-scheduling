# Jury Optimizer Contract

**Status:** Phase 2 core, Phase 3 backend integration, and read-only Schedule UI implemented. Manual schedule editing and finalization remain out of scope.

## Responsibilities

The Jury optimizer receives a readiness-approved, immutable snapshot and produces one complete outcome for every Jury-required source Lesson. It places Jury entries in time and reports explicit unscheduled outcomes. It preserves source identity, manual Panel selection, Panel dates, and finalized Accompanist assignments.

The pure-domain DTOs and `JuryOptimizer` protocol live in `webapp/backend/app/jury_optimizer_contract.py`. That module uses only Python's standard library; it has no ORM, API, persistence, React, FastAPI, Tauri, HTTP, or UI dependencies.

## Non-Responsibilities

The optimizer does not read Accompanist ORM tables, current editable Accompanist data, Jury persistence, workbook files, UI state, or solver internals. It does not perform readiness checks, change source assignments, substitute pianists, infer Panels, change Jury Required, persist results, or publish/finalize a result. An adapter outside the pure domain resolves contract data and readiness approval into DTOs.

The Phase 2 implementation is a deterministic earliest-feasible core in `webapp/backend/app/jury_optimizer.py`. It does not call or port `generate_jury_schedule.py` and does not implement a universal optimizer.

## Inputs

`JuryOptimizerInput` contains:

- a finalized `accompanist.assignment-result` contract v2 reference, including session/result UUIDs and source revision;
- the contract's lesson-level finalized facts, including source Lesson UUID, student Person UUID, independent Jury Required and pianist-required Booleans, and the optional fixed assigned Pianist Person UUID;
- Jury lesson entries and their explicit Panel UUID assignments;
- one date and one timing configuration per referenced Panel;
- earliest and preferred start, positive Jury duration, optional Meal Break interval, and optional every-X-Juries Periodic Break configuration;
- date-keyed, binary Jury Availability Windows for fixed Pianist UUIDs; and
- readiness approval tied to the exact session, finalized result UUID, source revision, and Jury input revision.

Display names are presentation values only. Lesson, Person, and Panel associations use UUIDs. No workbook row ordering or display-name matching participates in identity or tie-breaking.

Construction rejects non-v2/non-finalized results, mismatched or unapproved readiness, stale source revisions, duplicate Lesson/Panel identities, missing required lesson entries or Panel assignments, invalid panel/time intervals, absent fixed assignments for pianist-required lessons, and missing availability declarations for required fixed pianists. This is a typed contract guard, not a replacement for the Jury readiness service.

## Outputs

`JuryOptimizerResult` contains:

- scheduled Jury entries with Lesson, student, Panel, date, optional fixed Pianist, and start/end minutes;
- Meal-Break entries;
- Periodic-Break entries;
- unscheduled entries with a typed reason code and nonblank explanation;
- typed optimizer diagnostics; and
- structured conflict explanations referencing relevant UUIDs.

The result carries source result/session identity and source/Jury input revisions for downstream stale-dependency detection. Its constructor enforces that each Jury-required Lesson UUID occurs exactly once across scheduled and unscheduled entries, with no duplicates, omissions, or double disposition.

## Invariants

### Hard Constraints

- No overlapping Jury entries for the same Pianist on the same date, across all Panels and rooms.
- A Jury using a Pianist is wholly contained in one of that Pianist's declared Availability Windows for the Panel date.
- A finalized Accompanist assignment is fixed. Never replace it with another Pianist; a missing required assignment is not an invitation to schedule without one.
- The assigned Panel is authoritative; do not infer or change it.
- The assigned Panel's Jury Date is authoritative.
- Meal Break and configured Periodic Break intervals are respected; the Meal Break resets the periodic count.
- Jury duration equals the configured positive integer duration, including 60 minutes.
- No scheduled entry or break extends beyond minute 1440. Do not schedule after the Jury day.
- Only inputs approved by Jury readiness are accepted.

### Outcome Constraints

- Every Jury-required source Lesson receives exactly one disposition: scheduled or explicitly unscheduled.
- Output is deterministic for identical typed inputs, independent of database/workbook row order and display names.
- UUIDs are the only identity and association keys.
- Execution terminates finitely, including when no feasible interval exists.
- Every unscheduled entry includes a stable reason code and actionable explanation; no lesson is silently dropped.
- Result provenance retains the exact finalized Accompanist result and Jury input revision so later changes are detectable.

DTO constructors enforce structural input provenance, interval bounds, and complete/disjoint lesson dispositions. The Phase 2 behavior suite exercises availability, cross-Panel fixed-Pianist conflicts, authoritative assignments/dates, breaks, deterministic output, day bounds, and total outcomes.

## Termination Guarantee

The implementation scans monotonically increasing integer minutes bounded by the one-day horizon and processes each required Lesson at most once as scheduled or unscheduled. Candidate advancement is strictly forward to a meal end, break end, next availability start, conflicting booking end, or the end of the search horizon. It cannot retry indefinitely or use after-midnight placement as an escape hatch. Exhausting feasible placements produces explicit unscheduled outcomes for all remaining Jury-required Lessons.

## Unscheduled Reasons

Every unscheduled entry must carry a typed reason and a user-understandable explanation. The initial reason vocabulary includes no feasible interval, fixed-Pianist conflict, availability too short, and day-bound exceeded. Implementations may add stable codes when characterization demonstrates a distinct user-actionable cause. A generic empty result or silent omission is invalid.

## Determinism

Stable UUID ordering must resolve any otherwise equivalent choices. Input order, ORM query order, dictionary insertion order, workbook row order, display names, and room labels must not decide results. Identical input DTOs must yield equal outputs, including diagnostics and unscheduled explanations.

## Objective Strategy

The core follows the authorized priority: preserve hard constraints, keep placing feasible lessons, select earliest completion, use tighter fixed-Pianist availability as a deterministic tie-break, then UUIDs. Each Panel's timeline advances only when an entry or break is committed, avoiding unnecessary internal gaps. Room names and Preferred Start do not add scoring or constraints. This is a simple deterministic strategy, not a weighted objective system.

## Phase 1 and 2 Test Artifacts

Synthetic scenario builders and active DTO/behavior tests are in `webapp/backend/tests/fixtures/jury_optimizer_fixtures.py` and `webapp/backend/tests/test_jury_optimizer_contract.py`. No optimizer tests are skipped. Coverage includes total outcomes, impossible schedules, shared-Pianist conflicts, authoritative Panel/date, hard availability, break reset, day bounds, and deterministic output. All fixture identities and names are synthetic.

## Phase 3 Integration

`webapp/backend/app/services/jury_results.py` is the application boundary. It resolves the current finalized Accompanist v2 result, invokes the existing Jury readiness service, projects typed DTOs, calls `DeterministicJuryOptimizer`, and stores a versioned payload in the shared `module_results` registry with a `module_result_dependencies` row. No Jury schedule tables or database migration are added.

`webapp/backend/app/services/jury_sync.py` reconciles Jury lesson and Pianist UUID references against an Accompanist-owned identity snapshot on module activation and after source replacement/deletion operations. It removes only orphaned Jury references, keeps surviving Panel selections, refreshes a roster only from a current finalized result, and reports stale dependent results by revision/result comparison. Structured log events contain aggregate counts only, not student or Pianist names.

Generated results start as `draft`; generating again creates a new immutable result and marks the previous current draft `superseded`. The module revision's current-result pointer identifies the current snapshot. Historical results remain queryable. Result provenance includes source session/result/contract/version/revision, Jury input revision, and result version.

Staleness is computed from the stored dependency against the current Accompanist result UUID/source revision and current Jury input revision. A stale result is not relinked or rewritten; current/history views report stale reasons and retain the original source reference.

The read-only Schedule view consumes Panel timelines, scheduled and unscheduled entries, Meal/Periodic Break events, diagnostics, readiness warnings, and stale/lifecycle status from `GET /api/jury/results/current`, `GET /api/jury/results/history`, and `GET /api/jury/results/{result_id}`. `POST /api/jury/generate` is a thin service-backed operation. No manual edit/finalization workflow is included.