# LESSON-SCHEDULING-MODULE-SPEC.md

**Status:** Approved  
**Product:** Music Program Scheduler  
**Module:** Lesson Scheduling

---

# 1. Overview

Lesson Scheduling assigns students to instructors for recurring weekly lessons.

Key characteristics:

- Students receive at most one lesson assignment.
- Lessons assigned to the same instructor may not overlap.
- Students requiring multiple lessons are represented as multiple student records.
- Lesson Scheduling is independent from all other modules.

---

# 2. Module Structure

Tabs:

1. Instructor Roster
2. Student Roster
3. Schedule
4. Reports

## Shared Workflow and UI Consistency

Lesson Scheduling intentionally follows the Clinical Placement workflow:

```text
Data Entry
→ Readiness
→ Generate Schedule
→ Manual Adjustments
→ Reports
```

The modules should feel like parts of the same application. Reuse
well-established UI patterns, terminology, status presentation, revision and
stale-state behavior, manual-override concepts, report structure, and import
flows wherever practical. Differences belong in module-owned scheduling
rules, data fields, validation, and report meaning only when the scheduling
problem genuinely requires them. Do not copy Clinical's domain models,
catalog, tag meanings, or optimizer.

## Implementation Roadmap

- ✅ Foundation
- ✅ Instructor Roster
- ✅ Student Roster
- ✅ Candidate Generation
- ✅ Readiness Analysis
- ✅ Optimizer / Schedule Generation Review
- ✅ Manual Placement Management
- ⬜ Reports and Export

The optimizer, automatic generation, and Clinical-aligned schedule lifecycle
are implemented. Reports and Export remain pending.

## Phase Completion Notes

### Phase 2: Instructor Roster — Complete

Includes Instructor CRUD, Tag Manager, Availability editing, Instructor
import, and revision/staleness handling.

### Phase 3: Student Roster — Complete

Includes Student CRUD, Student Availability editing, Student import,
Required/Preferred/Restricted tag states, and revision/staleness handling.
Student import replaces student identity and Availability Window data in one
workflow.

### Phase 4 Completion Notes

Candidate Generation is complete.

Candidate generation is implemented and includes:

- `LessonSchedulingInput` snapshots
- Revision-stamped candidate generation
- 15-minute lesson start enumeration
- Required tag evaluation
- Restricted tag evaluation
- Preferred-tag counting
- Tentative-minute calculations

Candidate generation does not:

- Assign lessons
- Enforce capacity
- Enforce instructor overlap
- Perform optimization

At Phase 4 completion, readiness analysis, scheduling, and reporting remained
pending.

### Phase 5: Readiness Analysis — Implemented

Readiness analysis is complete and includes revision-stamped reports,
machine-readable warning codes, `GET /api/lesson/readiness`, and grouped
warnings in the Schedule view. It determines whether schedule generation is
possible; it does not create assignments.

### Phase 6 Completion Notes

The Lesson-specific optimizer and automatic schedule generation are
implemented. The optimizer selects only feasible candidates and enforces
one lesson per student, instructor non-overlap, effective capacity, and
available lesson time. Its strict lexicographic objectives and deterministic
tie-break remain Lesson-owned as described above.

Schedule generation preserves Manual assignments and Manually Unplaced
decisions, excludes Manually Unplaced students from automatic generation, and
replaces only Automatic assignments. Preserved manual assignments are fixed
inputs that consume instructor capacity and block overlapping automatic
lessons. Schedule results retain source-data provenance and advance
`schedule_revision` after generation.

Source-data changes retain placements and the current schedule, advance
`source_data_revision`, and mark the schedule stale. Placement verification
compares retained assignments with current student/instructor availability,
lesson length, tags, effective capacity, and instructor overlap. It reports
warnings without deleting or rewriting assignments.

Manual assignment, reassignment, Manually Unplaced, and Remove Override are
implemented through typed schedule APIs and Schedule-tab controls. Removing
an override returns the student to the automatic pool and does not regenerate
the schedule. The Schedule view shows assignment status and verification
warnings alongside readiness and schedule summary.

### Optimizer and Automatic Schedule Generation — Implemented

The Lesson-specific optimizer consumes one `LessonSchedulingInput` snapshot
and its matching `LessonCandidateSet`. It selects only feasible candidates
and enforces one lesson per student, instructor non-overlap, instructor
capacity, and available lesson time. It does not assign beyond the
instructor's effective capacity: the configured capacity is bounded by the
number of non-overlapping lesson-length intervals that fit within recurring
Availability Windows.

Optimization is strict lexicographic, not a weighted score:

1. Maximize the number of students scheduled.
2. Maximize total Preferred-tag matches.
3. Minimize candidate tentative minutes, using candidate values unchanged.
4. Minimize instructor gap minutes.
5. Choose a deterministic schedule by canonical student name/UUID, instructor
   name/UUID, weekday order, and start minute. The ordering is independent of
   import order.

The optimizer performs an exact branch-and-bound search over feasible
candidates, processing students in fewest-candidate-first order. Candidate
ordering and the final canonical assignment signature use stable student and
instructor names/UUIDs and weekday/start values; neither depends on roster or
import order. Candidate tentative-minute values are consumed as supplied and
are not recomputed.

Gap minutes are counted only between lessons within one contiguous instructor
availability window. Disjoint windows do not create a gap between them.

Generation is gated on readiness and persists placements, explicit
unplaced-student reasons, source-data revision, summary statistics, and a new
schedule revision. `GET /api/lesson/schedule` returns typed entries and
verification as well as summary state. `POST /api/lesson/schedule/generate`
generates or regenerates automatic assignments;
`POST /api/lesson/schedule/clear` clears only automatic assignments and
automatic unplaced results. Manual assignments and Manually Unplaced
decisions are preserved.

The schedule summary reports scheduled and unscheduled student counts,
instructor utilization (scheduled assignments divided by total effective
instructor capacity), preferred-tag matches, tentative minutes, and gap-cost
minutes. Unplaced students retain a typed reason: no feasible candidate,
capacity limit, or instructor overlap. Schedule clearing removes the
generated placements and unplaced records, advances `schedule_revision`, and
clears stale state; it does not change source-data revision.

Reports and Export remain pending.

---

# 3. Instructor Roster

Each instructor contains:

- Name
- Email
- Room
- Lesson Length
- Capacity
- Availability Windows
- Tags

## Room

Informational only.

No room-conflict detection is performed.

## Lesson Length

Default:

```text
60 minutes
```

Lesson length belongs to the instructor.

If an instructor offers multiple lesson lengths, create separate instructor records.

## Capacity

Capacity is the maximum number of students assigned to the instructor.

Actual assignments are limited by:

```text
min(capacity, schedulable lessons)
```

where schedulable lessons are determined by availability and lesson length.

---

# 4. Availability

Instructors and students both use the shared Weekly Availability Editor.

Features:

- Monday-Sunday
- 7:00 AM-9:00 PM
- Available
- Tentative
- Unavailable
- Click editing
- Drag painting

Availability is stored as recurring Availability Windows.

Overlapping or adjacent windows with the same status are normalized and merged.

---

# 5. Student Roster

Student Roster uses the same application patterns as Clinical Placement
where practical; its data and tag meanings remain Lesson-owned.

Fields:

- Student ID
- Name
- Email
- Availability Windows
- Required Tags
- Preferred Tags
- Restricted Tags

Student ID is required.

---

# 6. Tags

Tags follow the shared [Tag Architecture](SHARED-TAG-ARCHITECTURE.md) for
identity and infrastructure. Lesson Scheduling owns its own catalog; a tag
is identified by `(lesson-scheduling, tag_uuid)`, not by its name. A matching
name in Clinical or another module is unrelated.

Lesson display and comparison rules are:

- Case-insensitive
- Stored/displayed in uppercase
- Module-owned

Uppercase display is a Lesson policy; shared normalization provides
case-insensitive identity.

Examples:

```text
VOICE
VIOLIN
PIANO
OPERA
BEGINNER
ADVANCED
```

## Tag States

- Neutral
- Required
- Preferred
- Restricted

These are Lesson-owned assignment states and optimizer semantics, not global
tag states. Another module may define different states or none.

## Required

All Required tags must be present.

## Preferred

Each satisfied Preferred tag contributes one preference point.

## Restricted

Any matching Restricted tag disqualifies an instructor during automatic scheduling.

---

# 7. Candidate Generation

A typed Lesson Scheduling input snapshot contains the module-owned students,
instructors, recurring Availability Windows, tag assignments, and the
`source_data_revision` at which they were read. Candidate generation is
database-free and returns a typed candidate set stamped with that revision;
the repository exposes the snapshot and candidate-set operations but does
not optimize or assign lessons.

Each candidate represents one student, instructor, weekday, and lesson
start. It contains:

```text
student_id
instructor_id
day
lesson_start_minute
lesson_end_minute
preferred_tag_match_count
tentative_minutes
issues
is_feasible
```

Candidate starts occur at 15-minute increments. The lesson end is the start
plus the instructor's configured lesson length. A candidate is feasible only
when the entire half-open lesson interval fits within both the student's and
instructor's Available or Tentative Availability Windows and its tag rules
are satisfied. Missing availability is not availability.

All Required student tags must be present on the instructor. Any matching
Restricted student tag makes the candidate infeasible. Preferred tags do not
affect feasibility; each matching Preferred tag contributes one match.
Tentative minutes are the sum of Tentative minutes in the student's and
instructor's Availability Windows for that lesson interval.

The instructor's capacity is exposed as candidate metadata only. Candidate
generation does not consume capacity, prevent overlapping instructor
candidates, assign lessons, or perform readiness analysis. Capacity and
overlap enforcement belong to the future optimizer.

Candidate issue codes identify missing Required tags and matching Restricted
tags. A student/instructor pair with no shared legal Availability Window
produces no time candidate.

## Feasible Starts

Starts occur at:

```text
15-minute increments
```

A start is feasible only when:

- The entire lesson fits within instructor availability.
- The entire lesson fits within student availability.
- Required tags are satisfied.
- Restricted tags are absent.

Example:

```text
Availability:
8:00-12:00

Lesson Length:
60 minutes
```

Legal:

```text
11:00-12:00
```

Illegal:

```text
11:15-12:15
```

---

# 8. Tentative Availability

Availability status:

```text
Available     Preferred
Tentative     Allowed
Unavailable   Disallowed
```

Tentative minutes are cumulative.

Example:

```text
Student tentative:   15
Instructor tentative: 30
```

Result:

```text
45 tentative minutes
```

---

# 9. Scheduling Constraints

## Student Constraint

A student may receive at most one lesson assignment.

## Instructor Constraint

Lessons assigned to the same instructor may not overlap.

Example:

```text
9:00-10:00
9:15-10:15
```

is prohibited.

## Capacity Constraint

Assigned student count may not exceed instructor capacity.

---

# 10. Optimization

The optimizer uses strict lexicographic priorities.

## Priority 1

Maximize students scheduled.

## Priority 2

Maximize preferred-tag matches.

## Priority 3

Minimize tentative minutes.

## Priority 4

Minimize instructor scheduling gaps.

## Priority 5

Deterministic tie-break.

---

# 11. Gap Definition

Gap cost equals unused instructor availability between assigned lessons.

Example:

```text
8:00-9:00
9:15-10:15
11:15-12:15
```

Gap cost:

```text
15 + 60 = 75 minutes
```

## Availability Window Rule

Gaps are measured only inside a single availability window.

Example:

```text
Window 1:
8:00-12:00

Window 2:
2:00-5:00
```

The period:

```text
12:00-2:00
```

does not contribute to gap cost.

Total gap cost is the sum of gap minutes across all instructors.

---

# 12. Readiness Analysis

Readiness consumes a `LessonSchedulingInput` and its matching
`LessonCandidateSet`; mismatched source revisions are rejected. It does not
regenerate candidates or assign lessons. The report contains
`source_data_revision`, student/instructor/candidate counts, grouped
machine-readable warnings, and `is_ready`. Repeated warnings share one code
entry while retaining each affected Student, Instructor, or Tag as a subject.

`is_ready` is true only when there is at least one feasible candidate and no
blocking errors. Empty rosters, zero capacity, invalid instructor
Availability Windows, and no feasible candidates are blocking. Missing
availability, no feasible instructors for a student, Required tags with no
matching instructor, and instructors whose lesson length cannot fit their
availability are warnings.

Warning codes:

```text
NO_STUDENTS

NO_INSTRUCTORS

STUDENT_WITHOUT_AVAILABILITY

INSTRUCTOR_WITHOUT_AVAILABILITY

STUDENT_WITHOUT_CANDIDATES

REQUIRED_TAG_WITHOUT_MATCH

ZERO_CAPACITY

INVALID_INSTRUCTOR_WINDOW

INSTRUCTOR_WITHOUT_SCHEDULABLE_LESSONS
```

An instructor has no schedulable lessons when no 15-minute-aligned start
allows the configured lesson length to fit within that instructor's legal
Available or Tentative Availability Windows. Contiguous Available and
Tentative windows may jointly provide enough time; Unavailable minutes block
the lesson.

The local `GET /api/lesson/readiness` response includes readiness status,
counts, grouped warnings with subject details, and `source_data_revision`.
The Schedule tab compares that revision with the current Lesson module state
and tells the user when to refresh after source-data changes. The Schedule tab
also provides the Phase 6 schedule actions and aggregate-only schedule
summary.

---

# 13. Manual Overrides — Implemented

Follow the Clinical manual override model wherever practical. Coordinators can:

- Assign a student manually.
- Reassign a student to another instructor or lesson time.
- Manually unplace a student.
- Remove an override, returning the student to ordinary Unplaced and making
  them eligible for automatic generation.

The coordinator is authoritative for manual decisions. Manual placements may
violate automatic scheduling rules; allow the action and report all
applicable warnings instead of blocking it. Lesson-specific warnings include:

student or instructor unavailable, Required-tag mismatch, Restricted-tag
conflict, instructor capacity exceeded, or instructor lesson overlap. Manual
lesson times should accept a valid minute-level time, not be restricted to
the automatic optimizer's 15-minute start increments; the lesson end is
derived from the instructor's lesson length. Warnings do not block an
otherwise valid manual override.

---

# 14. Schedule Lifecycle Alignment

Lesson Schedule lifecycle should match Clinical Placement as closely as
possible:

- Generate Schedule replaces/regenerates only automatically generated
  assignments.
- Preserve manual assignments and Manually Unplaced decisions across
  regeneration; exclude Manually Unplaced students and schedule other eligible
  students around preserved manual assignments.
- Manual placements consume instructor capacity. Existing manual overlaps and
  other constraint violations remain visible as warnings and do not get
  silently rewritten by automatic generation.
- Source-data changes preserve the current schedule and mark it stale.
- Clearing the schedule requires confirmation and removes automatic
  assignments and generated unplaced reasons without deleting source data.
  Manual assignments and Manually Unplaced decisions are preserved by the
  ordinary clear-generated-schedule action. If a later product action is
  intended to delete all placement decisions, it must be separately named and
  confirmed, as in Clinical's Delete All Placements.

Lesson now follows the Clinical schedule lifecycle for manual assignments,
Manually Unplaced decisions, stale schedules, verification, and regeneration.
The Schedule view retains assignments after source changes, shows current
verification warnings, and exposes Automatic, Manual, Manually Unplaced,
Verified, Unverified, and Stale status using consistent terminology.

The Schedule tab mirrors the Clinical layout: a toolbar with Generate Schedule
(disabled unless readiness is ready) and Clear Schedule, readiness Warnings,
Placements, Unplaced Students, and the Schedule Summary (with the stale banner)
at the bottom. There is no separate readiness summary. Readiness reloads every time the
tab is selected (there is no Refresh Readiness button). Placed students appear
in a table with Instructor, Day, Start Time, Placement Type, and Status columns;
unplaced students appear in a separate table. The Instructor dropdown ends with
a "Manually unplaced" option, and changing the instructor, day, or start time is
treated as a manual adjustment; there are no Assign or Manually Unplace buttons.
Selecting the blank option in the Unplaced list removes the override.

### Persistence alignment

Lesson schema v14 records automatic, manual, and Manually Unplaced decisions
in one module-owned per-student lifecycle and keeps display snapshots separate
so assignments remain visible after a roster identity is removed. This is a
Lesson persistence mechanism, not a shared Assignment model. It preserves
revision provenance and supports stale-result verification without coupling
Lesson to Clinical tables. Do not add universal cross-module Assignment or
Schedule tables.

---

# 15. Reports — Pending

Lesson reports should follow the Clinical reporting philosophy and use the
same familiar report preview and local export patterns. Provide:

- a student-oriented lesson roster;
- an instructor-oriented lesson roster;
- an unplaced-student report;
- a warnings section, including relevant stale-schedule and manual-override
  warnings; and
- local Markdown and PDF export.

Use Lesson-owned fields and row semantics, but keep section structure,
ordering conventions, warning presentation, and export interaction familiar
to Clinical users. Do not expose internal IDs.

## Instructor Placement Roster

Sorted by:

1. Instructor
2. Day
3. Lesson Start

Displays:

- Instructor
- Room
- Lesson Length
- Assigned Students

## Student Placement Roster

Alphabetical by student.

Displays:

- Student
- Instructor
- Room
- Day
- Start Time

## Unplaced Student Report

Reasons:

```text
No feasible instructor

Capacity limit

Manually unplaced
```

### Three-State Placement Model (Unassigned / Assigned / Manually Unplaced)

Clinical Placement and Lesson Scheduling share the same placement-state model. The API exposes it as the typed `placement_state` field (`unassigned`, `assigned`, `manuallyUnplaced`) on every schedule entry. It is derived, not stored: Unassigned is the absence of a placement record.

| State | Placement | Generate Schedule |
|---|---|---|
| Unassigned | None | Eligible; may be assigned automatically |
| Assigned (Automatic or Manual) | Exists | Automatic is replaceable; Manual is preserved |
| Manually Unplaced | None (explicit coordinator decision) | Excluded; never assigned |

Placement dropdown order is always: `Unassigned`, a separator, the instructors, a separator, `Manually Unplaced`.

Example:

```
Unassigned
────────────
Mina Teacher
Jo Instructor
────────────
Manually Unplaced
```

- Choosing Unassigned removes any placement or Manually Unplaced decision (`DELETE /api/lesson/schedule/placements/{student_id}`) and returns the student to the automatic pool. It does not regenerate.
- Choosing a placement option saves a manual placement; choosing Manually Unplaced saves the exclusion decision.
- The summary reports Unassigned and Manually Unplaced counts separately.
- Both modules use identical terminology; this is the same model in both specifications.

## Stale Schedule Warning

Reports generated from stale schedules must display a warning.

---

# 16. Imports

Lesson imports should follow the same inspect/map/preview/apply interaction
and error/confirmation presentation patterns as Clinical imports wherever
practical. Import targets and replacement/preservation rules remain
Lesson-owned; every imported tag field must be explicitly mapped, and
unmapped columns must never become tags.

Current Lesson roster imports are full replacements. They preserve existing
placement decisions against their original stable student and instructor
UUIDs and mark any existing schedule stale; they do not match or relink those
decisions to replacement records by name, email, or institutional identifier.
Verification retains the old assignment for review and identifies removed
roster identities. Preview and confirmation must make the full-replacement
and stale-schedule behavior clear. Clinical imports may preserve matched
records according to Clinical-owned import rules; Lesson's replacement
semantics remain module-owned.

## Instructor Import

Supported fields:

- Name
- Email
- Room
- Lesson Length
- Capacity
- Availability
- Tags, when explicitly mapped to the Lesson tag field

Mapped tag values use the shared semicolon-separated format. Only columns
explicitly mapped as Lesson tag fields are parsed; unmapped columns never
become tags. Import preview shows matched and newly created Lesson tags.

## Student Import

Supported fields:

- Student ID
- Name
- Email
- Availability

Availability weekday cells accept semicolon-separated time ranges, with
optional `T`/`Tentative` prefixes. When neither time includes AM/PM and the
end time is earlier than the start time, the importer assumes an afternoon
end time:

```text
9:00-3:00  ->  9:00 AM-3:00 PM
8:00-1:00  ->  8:00 AM-1:00 PM
```

Explicit AM/PM values are not adjusted. Existing valid ranges remain valid;
malformed times, zero-length ranges, and conflicting overlaps are still
rejected.

Both imports use:

```text
Inspect
Preview
Apply
```

Student import does not currently include tag assignments. Any future tag
fields must be explicitly mapped to Lesson-owned state fields; unmapped
columns must not be interpreted as tags.

---

# 17. Shared Infrastructure

Reuse shared infrastructure and proven Clinical UI/workflow patterns without
reusing Clinical-owned catalogs or semantics, for:

- Availability editor
- Availability persistence
- Tag identity, normalization, assignment mechanics, import/export parsing,
  validation primitives, filtering, and reporting support
- Generic availability coverage primitives, where useful; Lesson candidate
  feasibility and tag interpretation remain Lesson-owned
- Import architecture
- Readiness framework
- Report framework
- Revision tracking
- Manual override framework
- Stale schedule presentation
- Report preview and Markdown/PDF export interaction

The Lesson tag catalog, assignment-state meanings, validation rules,
candidate-generation rules, optimizer behavior, workflows, and report
semantics remain Lesson-owned.

---

# 18. Dedicated Lesson Optimizer

Lesson Scheduling requires its own optimizer and objective function. Keep its
readiness inputs, revision checks, stale-result lifecycle, generation
workflow, manual-override lifecycle, and user-facing interactions aligned
with Clinical wherever possible. The key scheduling-constraint distinction
is:

- Clinical placements may overlap in time.
- Lesson assignments to the same instructor may not overlap.

This Lesson-specific overlap constraint is enforced for automatic generation
and surfaced as a warning for manual overrides. Module-specific objective
priorities and tag meanings remain owned by each optimizer; workflow
consistency does not imply sharing optimizer code or collapsing domain rules.

---

# 19. Non-Goals

Do not implement:

- Room conflict detection
- Cross-module data sharing
- Multiple assignments per student record
- Cloud synchronization
- Analytics
- Telemetry

---

# 20. Success Criteria

A coordinator can:

1. Import instructors.
2. Import students.
3. Maintain availability.
4. Review readiness and its warnings.
5. Generate an automatic schedule while preserving manual assignments and
   Manually Unplaced decisions.
6. Make manual assignments, reassignments, and manual-unplaced decisions,
   review non-blocking warnings, and remove an override.
7. Regenerate automatic assignments without changing manual decisions.
8. Review student-oriented, instructor-oriented, unplaced, and warning
   reports, and export them locally as Markdown or PDF.