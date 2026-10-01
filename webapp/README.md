# Music Program Scheduler Desktop App

A local-first desktop application for university music-program scheduling.
Accompanist Scheduling is the first module. The current React workflow,
FastAPI/SQLite service, and scheduling behavior are retained while Tauri 2
replaces Electron as the primary desktop shell. Electron remains available for
side-by-side validation; the two desktop shells use separate local databases.

1. **Import** lesson data from a CSV/XLSX file with a column-mapping wizard
   (or add lessons manually in the schedule grid).
2. **Pianists & Availability** — add pianists with a weekly hour cap, and
   click cells on a Monday–Friday calendar to cycle each half-hour from its
   default Unavailable state through Available and Tentative. Import pianist
   availability from CSV, XLSX, or XLS, review mapped ranges and validation,
   then continue editing the same persisted slots in the grid.
3. **Schedule** — run the same best-fit assignment algorithm as
   `generate_pianist_schedule.py` (fit tiers, required-pianist handling,
   conflict prevention, workload caps, tie-breaks), then manually adjust the
   result in an editable grid. Double-bookings, unassigned students, and
   over-cap pianists are highlighted live, and each pianist's weekly hour
   total recalculates (using the same union-of-time-windows rule as the
   original CLI tool) as you edit.
4. **Reports** — generate the same Markdown schedule report as
   `generate_lesson_markdown.py` (pianist schedules, students by
   instructor, students by name) from the current assignments, and
   download it.

## Architecture

- `backend/` — FastAPI + SQLAlchemy + SQLite. Tauri launches the existing
   Python service as a loopback-only child process and bundles it as a local
   executable for production. SQLite is stored in the OS per-user app-data
   folder; Electron keeps its own user-data path.
   The scheduling algorithm is
   ported from `generate_pianist_schedule.py` into
   `backend/app/services/scheduling.py`, operating on database rows instead
   of an Excel workbook. The existing `Organization` row is a legacy
   Accompanist persistence placeholder. Each active database represents one
   Scheduling Session; this does not introduce account or multi-tenant storage.
- `frontend/` — Vite + React + TypeScript single-page app, packaged by Tauri
   for Windows, macOS, and Linux. Electron remains a comparison build.

## Build Desktop Installers

Install frontend and backend dependencies as described in
[`docs/BUILDING.md`](../docs/BUILDING.md). From `webapp/frontend`, use
`npm run tauri:build` for the Tauri installer. Electron remains available with
`npm run dist:linux`, `npm run dist:mac`, or `npm run dist:win` while parity is
being validated. Build native installers on their target operating systems;
signing credentials are not required for development builds.

Tauri development uses `npm run tauri:dev`; it starts the local Python service
itself. `npm run dev` remains available for browser-based frontend work, with
the backend started separately. `npm run electron:dev` starts the retained
Electron shell.

## Running locally

### Backend

```bash
cd webapp/backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8123
```

The API is served at `http://localhost:8123` (interactive docs at
`/docs`). A SQLite file `pianist_scheduling.db` is created next to
`backend/` on first run.

### Frontend

```bash
cd webapp/frontend
npm install
npm run dev
```

Open the printed local URL (default `http://localhost:5173`). The
frontend talks to the backend at the URL in `webapp/frontend/.env`
(`VITE_API_BASE`, defaults to `http://localhost:8123`).

Use this development workflow while making changes. Vite refreshes frontend
edits, and Uvicorn reloads backend edits. Rebuild an installer only when
validating a release.

## Notes on the assignment algorithm

`backend/app/services/scheduling.py` mirrors the CLI tool's fit-scoring and
assignment-loop logic (see that file's docstring and
`generate_pianist_schedule.py` for the full rules). Manually reassigning a
lesson in the UI marks it as "manually edited"; re-running the algorithm
leaves manually edited lessons untouched but still accounts for them when
assigning everyone else (so it won't double-book a pianist you've already
placed by hand).

## Scheduling Sessions

The active session identifies one institution/program and academic term. Its
metadata and Accompanist data are saved automatically in the local database.
Use **Export Session** to create a portable `.mpsession` copy or **Open Session**
to validate and restore one. Replacing the active session requires confirmation
and first creates a local recovery archive; the application retains the three
most recent recovery archives under its application-data directory. Session
archives are unencrypted and may contain student educational information.

## Import Pianist Availability

Use **Import Availability** beside the existing pianist grid. The importer
accepts CSV, XLSX, and XLS, supports normalized one-window-per-row files and
wide respondent rows with configurable weekday start/end columns, and lets
you choose a workbook sheet and correct suggested column mappings. A normalized
example is available at
[`frontend/public/availability-template.csv`](frontend/public/availability-template.csv).

For a Microsoft Forms workflow, ask respondents for their name and one or more
weekday availability ranges (start and end time). Export responses to Excel,
select the workbook locally, map the respondent name and time columns, review
matches/warnings/errors, then apply the ranges and correct them in the normal
availability grid. Do not send the workbook to a Forms integration; no Forms
or Graph API is used.

Pianists match by exact case-insensitive name. Unknown names and duplicate
matches block apply; the importer does not create or merge pianist records.
Every accepted submission completely replaces availability for each matched
pianist; absent pianists remain unchanged. In a valid complete submission,
blank days/times mean no availability. Incomplete mappings or malformed rows
block the entire import rather than clearing a person's week. A wide Forms
workbook must map every weekday's start/end columns; a respondent with all
mapped windows blank is a valid zero-availability response. Existing manual
edits for matched pianists are replaced, with confirmation. The app stores
Available/Tentative windows sparsely and records weekly completeness separately;
absent slots in a complete week are Unavailable to the Accompanist solver.
Imported ranges must align to the existing 30-minute grid and are not rounded.
