# Music Program Scheduler Desktop App

A local-first desktop application for scheduling and resource management in university and private music programs.

The current application consists of a React/Tauri desktop frontend backed by a local FastAPI and SQLite service. All scheduling data remains on the user's device.

## Current Modules

### Lesson Scheduling

- Import instructor and student roster data from CSV/XLSX files.
- Manage instructor availability, teaching loads, lesson lengths, and teaching tags.
- Run automated lesson assignment and schedule optimization.
- Minimize instructor and student gaps while respecting scheduling constraints.
- Manually adjust assignments when needed.
- View readiness warnings, conflicts, and scheduling diagnostics.

### Clinical Placement

- Import students and clinical sites.
- Use tag-based candidate matching and placement optimization.
- Review placement recommendations and placement status.
- Manually override placements while preserving optimization functionality.

### Shared Features

- Tag-based matching and filtering.
- Import and export using CSV and XLSX.
- Schedule and placement reporting.
- Local SQLite data storage.
- Session import/export and backup recovery.

## Architecture

### Backend

Location:

```text
backend/
```

Technology stack:

- FastAPI
- SQLAlchemy
- SQLite
- Pandas
- OpenPyXL

Responsibilities:

- Scheduling and optimization engines
- Clinical placement logic
- Import/export processing
- Session management
- Report generation
- Local database storage

Tauri launches the backend as a loopback-only local service and bundles it as a standalone executable for production builds.

### Frontend

Location:

```text
frontend/
```

Technology stack:

- React
- TypeScript
- Vite
- Tauri

Responsibilities:

- User interface
- Scheduling workflows
- Clinical placement workflows
- Data management
- Reporting and exports

The frontend is packaged as a native desktop application for Windows, macOS, and Linux using Tauri.

## Build Desktop Installers

Install frontend and backend dependencies as described in:

```text
../docs/BUILDING.md
```

From:

```bash
webapp/frontend
```

build the desktop application:

```bash
npm run tauri:build
```

Build native installers on their target operating systems. Code signing is optional for development builds.

## Development Workflow

### Backend

```bash
cd webapp/backend

python3 -m venv .venv
source .venv/bin/activate

pip install -r requirements.txt

uvicorn app.main:app --reload --port 8123
```

The API is available at:

```text
http://localhost:8123
```

Interactive API documentation:

```text
http://localhost:8123/docs
```

### Frontend

```bash
cd webapp/frontend

npm install
npm run dev
```

Open the URL displayed by Vite (typically `http://localhost:5173`).

The frontend communicates with the backend using the URL configured in:

```text
webapp/frontend/.env
```

via:

```text
VITE_API_BASE
```

### Tauri Development

Run the desktop application in development mode:

```bash
npm run tauri:dev
```

Tauri automatically launches the local backend service.

## Scheduling Sessions

A Scheduling Session represents a single institution, program, and academic term.

Session features include:

- Automatic local persistence
- Session export/import
- Recovery archives
- Session replacement protection
- Local-only storage

The application automatically maintains recent recovery archives in the user's application-data directory.

> Session archives are unencrypted and may contain student educational information.

## Reporting

The application can generate reports from current scheduling and placement data, including:

- Lesson schedules
- Instructor schedules
- Student schedules
- Clinical placement reports
- Administrative summaries

Reports may be viewed, exported, or converted to PDF depending on the workflow.

## Documentation

Additional specifications and architecture documents are available in:

```text
../docs/
```

Notable documents include:

- ARCHITECTURE.md
- MODULE-ARCHITECTURE.md
- PRODUCT-MODULE-ARCHITECTURE.md
- LESSON-SCHEDULING-MODULE-SPEC.md
- CLINICAL-PLACEMENT-MODULE-SPEC.md
- SHARED-TAG-ARCHITECTURE.md
- PRIVACY-ARCHITECTURE.md

## Project Status

Music Program Scheduler has evolved from its original prototype command-line scheduling tools into a unified desktop application with:

- React/Tauri frontend
- FastAPI backend
- SQLite persistence
- Modern import/export workflows
- Piano accompanist scheduling module
- Performance jury scheduling module
- Lesson Scheduling module
- Clinical Placement module

All active functionality resides within the current webapp architecture.