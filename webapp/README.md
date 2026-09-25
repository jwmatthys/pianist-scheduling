# Pianist Scheduling Desktop App

A cross-platform desktop application version of the pianist scheduling tools in
this repo, with a graphical UI. The application bundles the React interface,
local scheduling service, and SQLite database; no server deployment is needed.

1. **Import** lesson data from a CSV/XLSX file with a column-mapping wizard
   (or add lessons manually in the schedule grid).
2. **Pianists & Availability** — add pianists with a weekly hour cap, and
   click cells on a Monday–Friday calendar to cycle each half-hour from its
   default Unavailable state through Available and Tentative.
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

- `backend/` — FastAPI + SQLAlchemy + SQLite bundled as a local executable.
   The database is stored in the operating system's per-user app-data folder.
   The scheduling algorithm is
  ported from `generate_pianist_schedule.py` into
  `backend/app/services/scheduling.py`, operating on database rows instead
  of an Excel workbook. The data model is organization-scoped (every table
  hangs off an `Organization`) so real multi-tenant accounts/auth can be
  layered on later without a schema rewrite; the MVP only ever uses one
  default organization.
- `frontend/` — Vite + React + TypeScript single-page app, packaged by Electron
   as native installers for Windows, macOS, and Linux.

## Build Desktop Installers

Build on each target operating system to produce its native installer:

```bash
cd webapp/backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

cd ../frontend
npm install
npm run dist:linux  # AppImage and .deb on Linux
# npm run dist:mac  # .dmg on macOS
# npm run dist:win  # NSIS installer on Windows
```

The generated installers are placed in `webapp/frontend/dist/`. The first
build downloads Electron and PyInstaller bundles the local API for the current
platform, so releases should be built natively or in an appropriate CI runner.
The Linux AppImage uses Electron Builder's static runtime, avoiding the
`libfuse2` dynamic-library dependency, but AppImage still requires kernel FUSE
support to mount its filesystem. Use the `.deb` installer on Debian-based
systems, or extract the `tar.gz` release and run its executable directly for a
portable distribution that does not use FUSE.
If the backend virtual environment already existed before this project added
desktop packaging, rerun `pip install -r requirements.txt` so it includes
PyInstaller.

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
