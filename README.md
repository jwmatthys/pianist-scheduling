# 🎼 Music Program Scheduler

Music Program Scheduler is a desktop application for managing and optimizing scheduling workflows in collegiate and private music programs.

The application provides:

- Lesson scheduling and optimization
- Clinical placement scheduling
- Instructor and student roster management
- Tag-based matching and filtering
- Schedule reporting and export
- Local-first desktop deployment via Tauri

## Architecture

The project consists of two primary components.

### Frontend

Location:

```text
webapp/frontend
```

Technology stack:

- React
- TypeScript
- Vite
- Tauri

Responsibilities:

- User interface
- Scheduling workflows
- Data management
- Reporting and exports

### Backend

Location:

```text
webapp/backend
```

Technology stack:

- Python
- FastAPI
- SQLAlchemy
- SQLite
- Pandas
- OpenPyXL

Responsibilities:

- Scheduling and optimization engines
- Clinical placement matching
- Lesson scheduling
- Data import/export
- Report generation
- Persistent storage

The backend is packaged as a standalone executable via PyInstaller and bundled into the desktop application during the Tauri build process.

## Repository Structure

```text
docs/
    Project specifications and architectural documentation

sample-data/
    Sample datasets for testing and development

webapp/
    backend/
        FastAPI application
        Scheduling services
        Database models
        Import/export logic
        Automated tests

    frontend/
        React/Vite/Tauri user interface
```

## Development Setup

### Backend

Create and activate a virtual environment:

```bash
cd webapp/backend

python -m venv .venv
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the backend:

```bash
python desktop_server.py
```

### Frontend

Install dependencies:

```bash
cd webapp/frontend

npm install
```

Run the frontend development environment:

```bash
npm run dev
```

### Tauri Development

Run the desktop application:

```bash
npm run tauri:dev
```

## Building

Build the frontend:

```bash
npm run build
```

Build the packaged backend executable:

```bash
npm run build:tauri-api
```

Build the desktop application:

```bash
npm run tauri:build
```

Build artifacts are generated in:

```text
webapp/frontend/src-tauri/target/release/bundle/
```

## Documentation

Project specifications and design documents are located in:

```text
docs/
```

Notable documents include:

- ACCOMPANIST-BEHAVIOR-CHARACTERIZATION.md
- ARCHITECTURE.md
- CLINICAL-PLACEMENT-MODULE-SPEC.md
- JURY-MODULE-DATA-DESIGN.md
- JURY-OPTIMIZER-CONTRACT.md
- JURY-SCHEDULER-CHARACTERIZATION.md
- LESSON-SCHEDULING-MODULE-SPEC.md
- MODULE-ARCHITECTURE.md
- PRIVACY-ARCHITECTURE.md
- PRODUCT-MODULE-ARCHITECTURE.md
- SHARED-TAG-ARCHITECTURE.md

## Testing

Backend tests:

```bash
cd webapp/backend
pytest
```

Frontend tests:

```bash
cd webapp/frontend
npm test
```

## Status

Music Program Scheduler has evolved from its original prototype scheduling scripts into a unified desktop application featuring:

- A React/Tauri frontend
- A FastAPI backend
- Integrated scheduling and optimization engines
- Local SQLite storage
- Modern import/export workflows

All active functionality is implemented within the `webapp` architecture. Development should target the frontend, backend, and documentation contained within that structure.