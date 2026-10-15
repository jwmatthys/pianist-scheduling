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

Student Roster behavior is identical to Clinical Placement.

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

Tags are:

- Case-insensitive
- Stored/displayed in uppercase
- Module-owned

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

## Required

All Required tags must be present.

## Preferred

Each satisfied Preferred tag contributes one preference point.

## Restricted

Any matching Restricted tag disqualifies an instructor during automatic scheduling.

---

# 7. Candidate Generation

A candidate consists of:

```text
Student
Instructor
Lesson Start Time
```

Candidate generation:

- Identifies legal possibilities.
- Does not assign lessons.
- Does not enforce capacity.

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

Warnings include:

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

An instructor has no schedulable lessons when at least one lesson cannot fit inside any availability window.

---

# 13. Manual Overrides

Coordinators may:

- Assign manually
- Reassign manually
- Change lesson times
- Mark students Manually Unplaced

Manual changes may violate automatic scheduling rules.

Warnings are displayed but do not block the action.

Potential warnings:

```text
Student unavailable

Instructor unavailable

Required tag mismatch

Restricted tag conflict

Capacity exceeded

Instructor lesson overlap
```

---

# 14. Schedule Lifecycle

Generate Schedule:

- Replaces automatic assignments.
- Preserves manual assignments.
- Preserves Manually Unplaced decisions.

Manual decisions always take precedence.

---

# 15. Reports

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

## Stale Schedule Warning

Reports generated from stale schedules must display a warning.

---

# 16. Imports

## Instructor Import

Supported fields:

- Name
- Email
- Room
- Lesson Length
- Capacity
- Availability
- Tags

## Student Import

Supported fields:

- Student ID
- Name
- Email
- Availability

Both imports use:

```text
Inspect
Preview
Apply
```

---

# 17. Shared Infrastructure

Reuse the existing Clinical infrastructure for:

- Availability editor
- Availability persistence
- Tag registry
- Import architecture
- Readiness framework
- Report framework
- Revision tracking
- Manual override framework

---

# 18. Dedicated Lesson Optimizer

Lesson Scheduling requires its own optimizer.

Clinical permits overlapping placements.

Lesson Scheduling prohibits overlapping lessons.

The optimizer is therefore module-specific.

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
4. Generate schedules.
5. Review readiness.
6. Make manual adjustments.
7. Regenerate while preserving overrides.
8. Produce reports.
9. Export reports as Markdown or PDF.