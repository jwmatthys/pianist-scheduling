# Performance Jury Scheduler Characterization

**Status:** Characterization and future design only. The standalone scheduler remains unchanged and is not integrated into the application.

## Current Script

`generate_jury_schedule.py` is a pandas/openpyxl command-line program. `load_students` joins a Lessons workbook to a separate Accompanist assignment workbook by exact, case-sensitive student display name. `load_pianist_unavailability` reads `Pianist - <Name>` sheets. `build_schedule` creates all panel schedules in memory, and `write_excel` emits a summary, one worksheet per Area, and one per pianist. The output file is written beside the Lessons workbook.

The input files are:

- Lessons sheet: Student Name, Instrument, Need Pianist, Jury, and Area.
- Schedule - Assignments sheet: Student and Accompanist.
- Pianist - <Name> sheets: time rows and Jury Day Availability.
- Jury Information sheet: Area, Jury Length, Start Time, Hourly Break, Lunch Break, and Location.

The output is a timestamped XLSX. It has no stable internal person identity, persisted result lifecycle, readiness report, manual editor, or typed dependency contract.

## Characterized Behavior

Thirteen synthetic tests in `webapp/backend/tests/test_jury_characterization.py` call the standalone loader and scheduler functions directly. No real student data or workbook fixture is used.

- An assigned pianist is inflexible. The script schedules with the assigned pianist when they become free and never considers an alternate pianist.
- Pianist availability is treated as a hard minute-level block. Each non-exact-`Available` value, including Tentative and blank, blocks the row's 30-minute interval. Rows missing from the sheet are not blocked, so the current loader is not a reliable closed-world completeness model.
- `pianist_booked` is shared across all Areas and prevents overlapping bookings for the same pianist, including across different rooms.
- Areas with the same Location share a room cursor and are serialized. Areas in distinct rooms can overlap unless they share a pianist. This room constraint is explicitly removed in the integrated module.
- Students are grouped by assigned pianist, retain original order within each group, and groups are ordered by the first available slot then the latest minute in the union of pianist unavailability and prior bookings. Synthetic tests confirm compact blocks and availability-driven order.
- When the next student is blocked, the scheduler scans the remaining list for the first currently feasible student. It schedules that student and revisits the blocked student on a later pass. If none is feasible, it advances to the earliest next-free minute found for the remaining assigned pianists.
- `find_actual_start` can delay a panel to reduce a large early gap. With two pianist-free students and a pianist blocked until 10:30, a 9:00 earliest start, and 10-minute juries, the test confirms a 9:40 actual start and a 30-minute gap before the pianist becomes available. This is useful evidence for Earliest Start as a hard bound and Preferred Start as a future soft objective, but the current three-slot gap threshold is not a product scoring rule.
- The hourly break is not elapsed-hour based. It derives a cycle as `(60 - 10) // Jury Length`, inserts a fixed 10-minute break after each cycle, and skips a break when only one student remains. A 60-minute jury length with hourly breaks enabled divides by zero; this is a confirmed edge bug, not a behavior to preserve.
- Lunch is hard-coded to one 30-minute interval beginning at the current schedule time once that time is at or after noon. It may begin after noon (the synthetic 11:50, 20-minute jury case inserts it at 12:10), and it resets the hourly-break counter. This fixed lunch model is replaced.
- Different Areas can have different jury lengths. `Area` is currently the panel-like grouping key, but it is not a flexible explicit panel identity.
- Repeating `build_schedule` with identical inputs returns identical in-memory schedules for the covered synthetic case. Some tie decisions remain dependent on workbook row and dictionary insertion order.
- There is no panel end-time constraint. A fully unavailable assigned pianist is eventually scheduled at 24:00 after the bounded `next_free` search ends; `fmt(1440)` renders that as `12:00 PM`. This is a confirmed horizon/formatting bug. The integrated module must report an unschedulable/readiness problem rather than put work beyond the day.

## Input and Identity Findings

`Jury` is included only for the literal string forms `1`, `1.0`, `True`, and `TRUE`. `Need Pianist` uses the same narrow values. Jury participation and accompaniment need are already separate workbook columns, but the integrated design requires a typed Jury Required Boolean, independent from the Accompanist-owned Pianist Required value.

Assignments are joined by stripped display name. The join is case-sensitive, duplicate student names overwrite earlier assignment rows in the dictionary, and students without a resolved assignment are omitted from the assignment map. `load_students` then sets `needs_pianist` to `bool(pianist)`: a student whose source row requires a pianist but whose assignment is missing/UNASSIGNED is silently changed to not needing one. This is a confirmed behavior to replace with a blocking readiness issue. Do not use this name join as the integrated identity boundary.

The pianist availability loader treats only the exact string `Available` as available; Tentative and every other populated status are unavailable. It expands each source row to 30 minutes. It does not check whether the source contains a complete day. The future Jury input is one-day, binary, closed-world Jury-Day Availability with available windows as hard legal intervals; Accompanist Tentative scoring must not leak into it.

## Classification

### Preserve Conceptually

- Finalized Accompanist pianist assignments are fixed; Jury may move jury times but never substitute another pianist.
- Pianist availability and global simultaneous-pianist conflicts are hard constraints.
- Compact pianist blocks, feasible-student deferral, and useful panel-start delay are valuable scheduling objectives/strategies to characterize against the future requirements.
- Different panels may use different jury lengths, and each panel runs until its schedulable students finish; there is no fixed panel end time.

### Replace

- Excel handoff and student-name joins become structured inputs and stable identity references.
- Area becomes an explicitly defined Jury Panel with manual student-to-panel assignment; no inference from instrument, teacher, lesson, or program.
- The old Jury flag becomes Jury Required, independent of Pianist Required. Jury-exempt students are excluded.
- A required pianist with no finalized assignment is a blocking readiness issue, not a student with no pianist requirement.
- Ambiguous Start Time becomes hard Earliest Start plus soft Preferred Start. The current delay heuristic informs objectives but does not define the final weight.
- Fixed hourly-break cycle and fixed noon lunch are replaced with optional every-X-juries breaks, configurable length (default Jury Length), and a configurable hard Meal Break interval. Meal Break resets the periodic count.
- Room/Location becomes descriptive only: shared room names do not serialize or warn.
- Availability is explicitly complete, binary Available/Unavailable, and closed-world within the Jury day. Tentative is not exposed.
- No scheduling past the supported Jury day to hide an unschedulable assigned pianist.

### New

- Jury Required Boolean and manual panel assignment for each required student.
- Jury Panel fields: Panel Name, descriptive Room, Earliest Start Time, Preferred Start Time, Jury Length, optional periodic-break configuration, and optional Meal Break start/end.
- A panel-specific vertical timeline with typed Jury, periodic-break, and meal-break entries; no single global time column is assumed.
- Manual reordering of students within a panel. Recompute times after a move, preserve the move even when it creates a hard conflict, surface that conflict immediately, and do not treat conflicted output as clean/finalized.
- Readiness validation before optimization, including missing required pianist assignments and incomplete Jury-day availability.

## Future Accompanist Dependency

The current script actually consumes Accompanist-owned Student Name, Instrument, Need Pianist, and a separate Student-to-Accompanist assignment export. The integrated module must consume authoritative lesson/student identity, accompaniment requirement, and finalized pianist assignment information from the Accompanist module; the jury optimizer must not maintain an editable assignment copy or depend on Accompanist tables, UI, or solver objects.

Before choosing the typed contract, document the characterized fields and validate the actual integrated Jury workflow requirements. A generated Jury result will need enough provenance to identify the Scheduling Session and the finalized Accompanist result version/revision it consumed, plus its own Jury input/result revision. If the consumed Accompanist result changes, the Jury result must be detectable as stale/superseded. This document does not define the final envelope schema.

## Future Model Requirements

- Students with Jury Required = No are excluded, regardless of Pianist Required.
- Every Jury-required student is manually assigned to a named Jury Panel; no inferred panel mapping or batch assignment is assumed.
- Jury-day pianist availability is Jury-owned, one-day, binary, and complete. Available windows are legal intervals; all other times are unavailable.
- A pianist may not accompany overlapping juries in any panels/rooms. An existing finalized Accompanist assignment is not substitutable.
- Missing required pianist assignment is a blocking readiness problem. Manual override policy remains undecided.
- Each panel has its own timeline, jury duration, breaks, and meal interval. No panel end time or room exclusivity is introduced.

## Risks and Open Decisions

- Select the future optimizer objective/weighting among Preferred Start proximity, compactness, pianist conflicts, and panel delays; do not copy the current three-slot threshold as an authoritative weight.
- Define explicit horizon bounds and user-facing handling for a panel/pianist combination with no legal slot. The current 24:00 behavior is invalid.
- Define tie-breaking independently of workbook order and retain deterministic regression tests.
- Decide how duplicate source identities and students in multiple lesson rows map to one Jury participant without display-name joins.
- Gather Jury-day availability collection and completion requirements, readiness override policy, and manual conflict/finalization rules from workflow owners.
- Characterize behavior with synthetic equivalence tests before replacing or porting the standalone algorithm. Do not integrate it in this phase.
