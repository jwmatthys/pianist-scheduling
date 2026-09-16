# 🎹 Music Jury & Pianist Scheduling Suite

A pair of Python scripts that automate two distinct scheduling problems for music programs:

1. **`generate_pianist_schedule.py`** — Assigns piano accompanists to student lessons based on availability, workload caps, preference, overlap rules, and travel-aware block scheduling.
2. **`generate_jury_schedule.py`** — Generates a jury day schedule by sequencing students across rooms, resolving pianist conflicts across concurrent areas, and inserting breaks.

Both scripts read from a shared Excel workbook and write timestamped `.xlsx` output files.

---

## Table of Contents

- [Requirements](#requirements)
- [Setup](#setup)
- [Workbook Structure](#workbook-structure)
- [Usage](#usage)
  - [Pianist Assignment](#pianist-assignment)
  - [Jury Schedule Generation](#jury-schedule-generation)
- [Output Files](#output-files)
- [Algorithm Details](#algorithm-details)
  - [Pianist Assignment Algorithm](#pianist-assignment-algorithm)
  - [Jury Scheduling Algorithm](#jury-scheduling-algorithm)
- [Flags and Warnings](#flags-and-warnings)

---

## Requirements

- Python 3.9+
- The following PyPI packages:
  - `pandas`
  - `openpyxl`

---

## Setup

### 1. Create a Virtual Environment

```bash
# Create the venv
python3 -m venv .venv

# Activate it (macOS/Linux)
source .venv/bin/activate

# Activate it (Windows)
.venv\Scripts\activate
```

### 2. Install Dependencies

```bash
pip install pandas openpyxl
```

Or, if you have a `requirements.txt`:

```bash
pip install -r requirements.txt
```

**`requirements.txt`** (create this if needed):
```
pandas>=2.0
openpyxl>=3.1
```

### 3. Using the Makefile

A `Makefile` wraps the common commands:

```bash
make install    # pip install -r requirements.txt
make testdata   # generate sample lesson/pianist/jury-info workbooks
make pianist    # run the pianist assignment scheduler
make jury       # run the jury schedule generator (auto-detects the latest assignments file)
make markdown   # generate pianist and instructor Markdown schedules
make clean      # remove generated workbooks and Markdown reports
make distclean  # also remove the generated sample input workbooks
```

Override any input file via variables, e.g. `make pianist LESSONS=my_lessons.xlsx PIANISTS=my_pianists.xlsx`.
For the Markdown report, use `make markdown` or override the master workbook and output path:

```bash
make markdown
make markdown MARKDOWN_INPUT=my_master.xlsx MARKDOWN_OUTPUT=my_schedule.md
```

`make testdata` uses a fixed seed (`SEED=42`) by default so runs are reproducible. Pass `make testdata SEED=random` for a fresh random dataset each time, or `SEED=<int>` for a different fixed dataset.

---

## Workbook Structure

The scripts operate on **three separate Excel workbooks**, each with a single responsibility:

| Workbook | Contains | Used by |
|---|---|---|
| **Lesson information** | `Lessons` sheet | Both scripts |
| **Pianist availability** | `Pianist - [Name]` sheets | Both scripts |
| **Jury information** | `Jury Information` sheet | `generate_jury_schedule.py` |

`generate_jury_schedule.py` additionally reads the timestamped **assignments** workbook produced by `generate_pianist_schedule.py` to know which accompanist was assigned to each student.

### Lesson Information Workbook — `Lessons` Sheet

The lesson schedule. Required columns:

| Column | Description |
|---|---|
| `Lesson Teacher Name` | Teacher's name |
| `Student Name` | Student's name |
| `Lesson Day` | Day of the week (e.g., `Monday`) |
| `Lesson Start Time` | Lesson start time |
| `Lesson End Time` | Lesson end time |
| `Lesson Location` | Room or location |
| `Instrument` | Student's instrument |
| `Need Pianist` | `1` / `TRUE` if an accompanist is required |
| `Required Accompanist` | *(Optional)* Name of a specific required pianist |
| `Jury` | `1` if this student has a jury exam |
| `Area` | The jury area this student belongs to |

### Jury Information Workbook — `Jury Information` Sheet

One row per jury area. Required columns:

| Column | Description |
|---|---|
| `Area` | Area name (must match `Lessons` sheet) |
| `Start Time` | **Earliest** possible start time for the area |
| `Jury Length` | Length of each student's jury slot in minutes |
| `Location` | Room where this area's juries take place |
| `Hourly Break` | `1` / `TRUE` to insert 10-minute breaks each hour |
| `Lunch Break` | `1` / `TRUE` to insert a 30-minute lunch break at noon |

### Pianist Availability Workbook — `Pianist - [Name]` Sheets

One sheet per accompanist, named exactly `Pianist - Firstname Lastname`. Each sheet contains:

- **Cell with label `Name:`** followed by the pianist's full name
- **Cell referencing weekly hour cap** (a cell with "hour" in its label, followed by a number)
- **Day header row** with day names (`Monday`, `Tuesday`, etc.) plus a final `Jury Day Availability` column
- **Availability grid**: rows of 30-minute time slots, one column per weekday plus one `Jury Day Availability` column, with status values:
  - `Available` — fully available (highest preference)
  - `Tentative` — available but deprioritized
  - `Unavailable` — cannot be assigned

> **Note:** `Tentative` slots are treated as **unavailable** in the jury scheduler (which reads the `Jury Day Availability` column), but as a lower-preference option in the pianist assignment scheduler (which reads the weekday columns).

---

## Usage

### Pianist Assignment

Assigns accompanists to all lessons flagged with `Need Pianist = 1`.

```bash
python generate_pianist_schedule.py --lessons <lessons.xlsx> --pianists <pianists.xlsx>
```

**CLI Arguments:**

| Argument | Required | Description |
|---|---|---|
| `--lessons` | ✅ Yes | Path to the lesson information workbook (`Lessons` sheet) |
| `--pianists` | ✅ Yes | Path to the pianist availability workbook (`Pianist - *` sheets) |

**Example:**

```bash
python generate_pianist_schedule.py --lessons lesson_information.xlsx --pianists pianist_availability.xlsx
```

Output is written to the same directory as `--lessons`, with a timestamp appended:
```
lesson_information_2026-02-20_14-35.xlsx
```

---

### Jury Schedule Generation

Builds a jury day timetable from lesson data, jury area rules, and pre-assigned pianist pairings.

```bash
python generate_jury_schedule.py --lessons <lessons.xlsx> --pianists <pianists.xlsx> \
  --assignments <assignments.xlsx> --jury-info <jury_information.xlsx>
```

**CLI Arguments:**

| Argument | Required | Description |
|---|---|---|
| `--lessons` | ✅ Yes | Path to the lesson information workbook (`Lessons` sheet, filtered to `Jury = 1` rows) |
| `--pianists` | ✅ Yes | Path to the pianist availability workbook (reads the `Jury Day Availability` column) |
| `--assignments` | ✅ Yes | Path to the timestamped output of `generate_pianist_schedule.py` (`Schedule - Assignments` sheet) |
| `--jury-info` | ✅ Yes | Path to the jury information workbook (`Jury Information` sheet) |

**Example:**

```bash
python generate_jury_schedule.py \
  --lessons lesson_information.xlsx \
  --pianists pianist_availability.xlsx \
  --assignments lesson_information_2026-02-20_14-35.xlsx \
  --jury-info jury_information.xlsx
```

Output is written to the same directory as `--lessons`, timestamped:
```
jury_schedule_20260223_1213.xlsx
```

---

### Markdown Schedule Report

Generates a Markdown report from a master lesson workbook. The report contains
each pianist's schedule and an instructor-by-instructor list of students with
their assigned pianist.

```bash
make markdown
```

By default, this reads `lesson_final_fa26.xlsx` and writes
`lesson_final_fa26.md`. Override either path with `MARKDOWN_INPUT` and
`MARKDOWN_OUTPUT`.

---

## Output Files

### Pianist Assignment Output (`*_YYYY-MM-DD_HH-MM.xlsx`)

Contains three sheets:

| Sheet | Contents |
|---|---|
| `Schedule - Assignments` | Full row-by-row lesson assignment table with fit quality, notes, and flags |
| `Schedule - Weekly` | Visual weekly grid per accompanist showing their assigned lessons |
| `Schedule - Summary` | Per-accompanist stats: hours assigned, cap utilization, fit quality breakdown, conflicts |

### Jury Schedule Output (`jury_schedule_*.xlsx`)

Contains one summary sheet and one sheet per jury area:

| Sheet | Contents |
|---|---|
| `Summary` | One row per area: location, start/end times, student count, any delay or gap warnings |
| `[Area Name]` (one per area) | Ordered time-slot table with student, instrument, pianist, and any wait-gap annotations |

---

## Algorithm Details

### Pianist Assignment Algorithm

#### 1. Data Loading & Sorting

Lessons are filtered to those requiring a pianist (`Need Pianist = 1`), stripped of invalid rows, then sorted by day and start time. Accompanist availability sheets are parsed into per-day, per-30-minute-slot status maps.

#### 2. Fit Scoring

Each accompanist is scored against each lesson window using a tiered fit system (highest to lowest):

| Level | Score | Condition |
|---|---|---|
| **Full** | 4 | All 30-min slots covering the lesson are `Available` |
| **Partial** | 3 | All slots are `Available` or `Tentative` (at least one `Tentative`) |
| **Near** | 2 | Lesson start falls within 15 minutes before an available/tentative slot boundary |
| **Overlap** | 1 | Lesson overlaps ≤30 minutes with an already-assigned lesson; the same accompanist covers the full combined window |
| **None** | 0 | No usable availability — best-effort assignment, flagged as conflict |

#### 3. Assignment Priority

For each lesson, the algorithm applies the following decision tree:

1. **Required accompanist:** If a `Required Accompanist` is specified, only that pianist is considered. Partial name matching (substring) is used to resolve informal names. If they are unavailable, the lesson is left unassigned and flagged.

2. **Best standard fit:** Among all accompanists, the highest fit tier is identified. All candidates at that tier are ranked by:
   - Not over weekly hour cap (preferred)
   - Non-tentative availability (preferred)
   - Fewer active campus days
   - Fewer separate blocks, then less same-day gap time
   - Lower workload ratio (hours assigned / cap)

3. **Overlap fit:** If no accompanist achieves `Near` fit or better, the algorithm checks whether the lesson can share an accompanist with a recently assigned lesson that overlaps by ≤30 minutes, provided the accompanist covers the combined window.

4. **Conflict fallback:** If no fit exists, the least-loaded accompanist is assigned and the lesson is flagged as a conflict.

#### 4. Travel-Aware Schedule Penalty

To reduce pianist travel to and from campus, each candidate is scored by the schedule they would have *after* receiving the lesson. Tie-breaking prefers, in order:

1. **Fewer active campus days**
2. **Fewer separate blocks across the week**
3. **Less same-day gap time**

This means a pianist keeping two students on the same day with a longer break between them is intentionally preferred over splitting those students across additional days.

#### 5. Workload Balancing

Each accompanist has an optional maximum weekly hours cap. Assignments that would exceed the cap are still made (to avoid leaving lessons unassigned) but are flagged with `⚠ OVER CAP`.

---

### Jury Scheduling Algorithm

#### 1. Room Sequencing

Jury areas are grouped by room. Within each room, areas are sorted by their earliest start time and scheduled sequentially — the next area cannot begin until the previous one ends. This ensures no two areas ever compete for the same physical space.

#### 2. Pianist Availability Parsing

Each `Pianist - *` sheet is parsed into a set of unavailable minutes. Any slot not marked exactly `Available` (including `Tentative`) is treated as unavailable for jury purposes.

#### 3. Pre-Assignment Sorting (Cross-Area Block Scheduling)

Before scheduling an area, students are grouped by their assigned pianist and the groups are reordered. Groups whose pianist finishes their commitments in *other* rooms earliest are placed first. This reduces cross-area conflicts and produces more contiguous pianist blocks across the full day.

#### 4. Gap-Free Start Time Search (`find_gapfree_start`)

Rather than accepting the earliest possible start time naively, the scheduler searches for the earliest start from which the entire area can be scheduled *without any internal wait gaps*. It evaluates the declared earliest start time plus every minute at which any pianist transitions from busy to free, picking the first candidate that results in a fully contiguous schedule via simulation.

#### 5. Greedy In-Area Placement

Students are placed one slot at a time using a greedy scan:

- At each time step, the first student in the sorted list whose pianist is free is placed immediately.
- Students whose pianist is temporarily blocked (due to a concurrent commitment in another room) are skipped and revisited in subsequent iterations.
- When no student can be placed, the scheduler jumps directly to the next minute at which any remaining pianist becomes free — avoiding minute-by-minute scanning.

This combination of pre-sorting and greedy skipping minimizes wait gaps while respecting all cross-room pianist conflicts.

#### 6. Break Insertion

- **10-minute hourly breaks** are inserted after each complete cycle of jury slots (calculated as `(60 - 10) / slot_length` students per cycle), when enabled.
- **30-minute lunch breaks** are inserted when the clock reaches noon, when enabled.
- Both break types reset the hourly jury counter.

---

## Flags and Warnings

### Pianist Assignment Flags

| Flag | Meaning |
|---|---|
| `⚠ CONFLICT` | No available accompanist found; best-effort assignment made |
| `⚠ REQUIRED PIANIST '...' NOT FOUND` | The required pianist name doesn't match any loaded sheet |
| `⚠ REQUIRED PIANIST '...' is UNAVAILABLE` | Required pianist has no usable availability; lesson unassigned |
| `⚠ NEAR FIT` | Accompanist available within 15 minutes of lesson start |
| `⚠ OVER CAP` | Assignment exceeds the accompanist's weekly hour cap |
| `ℹ PARTIAL FIT` | At least one slot is `Tentative` rather than `Available` |
| `ℹ TENTATIVE availability` | Assignment relies on tentative availability |
| `ℹ OVERLAP FIT` | Lesson shares an accompanist with an overlapping lesson |
| `ℹ Required pianist '...' matched to '...'` | Partial name match was used to resolve the required pianist |

### Jury Schedule Flags

| Flag | Meaning |
|---|---|
| `Starts Xm after earliest (HH:MM)` | Area couldn't begin at its declared earliest time |
| `Xm wait gap(s) within area` | One or more students had to wait due to pianist unavailability |
| `+Xm wait` (in area sheets) | Gap before a specific student's slot |
