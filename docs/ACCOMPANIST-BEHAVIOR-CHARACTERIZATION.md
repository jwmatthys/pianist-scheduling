# Phase 2: Accompanist Behavior Characterization

**Status:** Phase 2.5 policy decisions are implemented and covered by synthetic regression tests. No persistence schema, application architecture, or Electron implementation changes were made.

## Test Run

The backend environment has the application dependencies but does not have `pytest`. The tests use Python's standard-library `unittest`; no test dependency was added.

From the repository root:

```sh
PYTHONPATH=webapp/backend webapp/backend/.venv/bin/python -m unittest discover -s webapp/backend/tests -v
```

The final run passed **24 tests**. SQLAlchemy emitted a non-failing `datetime.utcnow()` deprecation warning from the existing `Organization.created_at` default.

Test files:

- [`test_scheduling_characterization.py`](../webapp/backend/tests/test_scheduling_characterization.py)
- [`test_import_persistence_characterization.py`](../webapp/backend/tests/test_import_persistence_characterization.py)

Tests call solver functions and synchronous route handlers directly. They do not start the HTTP server or exercise the React UI.

## Baseline Coverage Retained

- Full, Partial, Near, and None availability fit scoring in both accompanist engines.
- Required-pianist substring matching and required assignments when the requested pianist is unavailable or over cap.
- Fit ranking: Full beats Partial when both candidates are under cap. The current ordinary-assignment policy prefers an under-cap Partial candidate over an over-cap Full candidate.
- CLI overlap-fit assignment, merged hours, desktop merged hours after manual overlap, and double-booking detection.
- Ordinary no-candidate outcomes and the zero-pianist edge case; both engines now leave these lessons unassigned with reasons.
- Required-pianist overlaps and manual desktop assignment locks.
- Schedule block/gap penalties with merged intervals; lesson locations still do not alter assignment scoring.
- Desktop validation after manual edits, repeated assignment runs, clear-assignment behavior, and persistence of assignment/manual-edit state.
- CSV mapping, XLSX time-cell parsing, day aliases, default 50-minute end time, optional `Need Pianist` default, invalid-row warnings, replacement-import behavior, and saved import-profile persistence.
- Availability bulk replacement persistence and startup migration of legacy SQLite rows missing `teacher_email` and `student_id`.

## Agreement Between CLI and Desktop

| Behavior | Characterized result |
|---|---|
| Availability fit tiers | The four tested fit outcomes and their scores agree. |
| Required-name substring match | The tested partial name resolves to the same pianist. |
| Required assignment with unavailable pianist | Both assign the requested pianist and mark unavailable. |
| Required assignment over cap | Both keep the required assignment and derive an over-cap warning from current union hours. |
| Ordinary fit/cap preference | Both prefer Full over Partial when both are under cap, and prefer an under-cap Partial candidate over an over-cap Full candidate. |
| Simple schedule penalty | Active-day, contiguous-block, and separated-gap examples agree. |
| Union working hours | Both calculate union time rather than double-counting overlap. CLI and desktop automatic overlap cases, plus manual desktop assignments, are covered. |
| Location influence | Neither uses room/site distance in assignment scoring. The CLI retains location as output data; the desktop engine record has no location field. |

These are fixture-level results, not a proof of full algorithm equivalence.

## Original Behavior, Policy, and Correction

| Scenario | Phase 2 baseline | Adopted policy | Corrected behavior |
|---|---|---|
| Required pianist conflicts | CLI retained both required assignments; desktop left the later one unassigned. | A required assignment is mandatory unless the user changes it; conflicts must be prominent. | Both retain the assignment. Both lessons are flagged; desktop validation also returns the conflict. |
| Ordinary lesson with no valid candidate | CLI assigned a conflicting best-effort candidate; desktop left it unassigned. | Never create an invalid conflict just to improve coverage. | Both leave it unassigned with a specific explanation; no candidate roster also returns a reason without throwing. |
| Compatible overlap | CLI allowed a characterized overlap when the combined window was covered; desktop had no automatic overlap path. | Preserve compatible overlap; distinguish it from invalid conflict and count union time. | Both use the existing CLI overlap rule. Solver-created pairs are marked `Overlap` and excluded from invalid-conflict reports; manual edits become `Manual` and are validated as potential conflicts. |
| Over-cap warning | Desktop recomputation could remove an existing warning; CLI warnings were not recomputed from final union totals in one centralized pass. | Derive warnings from current assignments and union working hours, not notes. | Both recalculate warnings from final union hours; stale warnings are removed when assignments no longer exceed cap. |
| Nested/overlapping schedule intervals | Both computed a 45-minute gap for the 09:00-10:00, 09:15-09:30, 10:15-11:00 example. | Compute gaps between merged intervals. | Both merge intervals first and report the correct 15-minute gap. |
| Location/travel | Neither engine used actual location distance; README wording overstated the behavior. | Do not claim actual travel optimization. | Current behavior is described as schedule consolidation/block and gap minimization. Location-based travel modeling remains future work. |
| Manual assignment and validation | CLI has no manual-edit workflow. | Keep manually imposed overlap distinct from optimizer-approved overlap. | Desktop locks edited assignments, labels edits `Manual`, and validation flags invalid overlaps while preserving union hours. |
| Import/persistence | CLI uses workbook-specific inputs and outputs. | Preserve current desktop import and local persistence contracts. | Desktop CSV/XLSX mapping, replacement import, saved profiles, availability persistence, and legacy-column upgrade behavior are covered. |

The automatic overlap policy is intentionally the characterized CLI rule, not a new generalized overlap model: it applies only when no standard fit is selected, the current and prior lesson are not required-pianist assignments, their overlap is at most 30 minutes, and pianist availability covers the combined window at Full or Partial fit. This threshold and rule remain fixed for now; whether to make them configurable is future work.

## Remaining Ambiguities

1. **Overlap configurability.** The existing 30-minute, combined-availability rule is preserved, but its threshold and eligibility rules are not user-configurable. No broader policy was invented.
2. **Zero-hour cap.** Existing truthiness checks treat a configured cap of `0` as no cap. The product decision did not define zero as a valid hard cap, so that behavior remains unchanged and untested.

Actual room/site travel distance and transition-time scoring remain explicitly out of scope.

## Remaining Regression Gaps

- No HTTP-level tests, React interaction tests, or end-to-end desktop workflow tests; current tests call service functions and route handlers directly.
- No end-to-end CLI workbook fixture validates generated Excel sheet names, formatting, timestamp rows, or the assignment workbook contract consumed by the jury script.
- The standalone jury, room-schedule, and Markdown report generators remain outside this accompanist characterization suite.
- CSV and XLSX parsing are covered, but malformed/corrupt workbooks, duplicate headers, ambiguous required-name matches, and upload-cache lifecycle/error cases are not.
- Persistence tests cover import replacement/profile storage, availability replacement, manual assignment state, and the two existing legacy columns; they do not characterize every deletion/cascade or database corruption/recovery path.
- Cross-platform installer builds and real desktop startup were not run as part of this behavior suite.
- Tests invoke solver functions and synchronous route handlers directly; they do not exercise HTTP transport, React workflows, or desktop interaction.
- The standalone jury, room-schedule, and Markdown report generators remain outside this accompanist suite.
- Generated CLI Excel workbook structure and the jury script's workbook handoff are not covered end to end.
- Cross-platform installer builds and real desktop startup were not run.

**Phase 2.5 exit:** adopted policies are implemented in the existing Python engines, corrected behavior is covered by 24 deterministic tests, and the remaining overlap configurability question is documented. No Tauri scaffolding or architecture migration has begun.