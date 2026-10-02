# Building

## Prerequisites

- Node.js and npm.
- Python matching the backend requirements. The build scripts use `webapp/backend/.venv` when present; set `PYTHON` to override the interpreter.
- Rust stable with Cargo for Tauri development and builds.
- On Linux, the Tauri system dependencies for the distribution in use, including a C toolchain, `pkg-config`, GTK3, WebKit2GTK 4.1, librsvg, and appindicator development packages. Consult the current Tauri 2 Linux prerequisites for the specific distribution.
- Development builds do not require signing credentials. Build macOS and Windows installers on their native platform for release validation.

Install frontend dependencies and prepare the Python environment from the repository root:

```sh
cd webapp/backend
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
cd ../frontend
npm install
```

On Windows, activate the backend environment with `.venv\\Scripts\\activate` and use `python` where the examples use `python3`.

## Development Commands

From `webapp/frontend`:

- `npm run dev` runs the browser/Vite frontend at `http://127.0.0.1:5173`. For backend functionality, separately run `uvicorn app.main:app --host 127.0.0.1 --port 8123` from `webapp/backend` with its virtual environment active. Override the API URL with `VITE_API_BASE` if needed.
- `npm run electron:dev` builds the packaged Python API executable and starts the retained Electron shell.
- `npm run tauri:dev` starts the Tauri shell plus Vite development server; Rust launches `backend/.venv`'s `desktop_server.py` directly (or the configured Python interpreter if the virtualenv is absent).
- `npm run build` typechecks and builds the React frontend.
- `npm test` runs the Python `unittest` suites, including schema migrations, Accompanist behavior/revision paths, Jury identities/results/input/readiness/generation/API/staleness, and `.mpsession` round trips. The script selects `webapp/backend/.venv` when available or uses `PYTHON`/the platform's Python command.

Tauri starts its development or packaged API process itself on an ephemeral loopback port; do not start Uvicorn separately for `npm run tauri:dev`.

## Tauri Production Build

From `webapp/frontend`:

```sh
npm run tauri:build
```

The Tauri `beforeBuildCommand` builds the frontend and PyInstaller executable, then Tauri bundles both. The build hook produces `pianist-scheduling-api-<Rust-target-triple>` (`.exe` is added on Windows) under `webapp/backend/dist`, matching the `externalBin` source path in `src-tauri/tauri.conf.json`. Tauri strips the target suffix when staging the executable beside the app binary. The build hook reads `TAURI_ENV_TARGET_TRIPLE` when Tauri supplies it, verifies it matches the local Rust host triple, and rejects cross-target PyInstaller builds. Build natively for each target: Windows x86_64 (`x86_64-pc-windows-msvc`), macOS Apple Silicon (`aarch64-apple-darwin`), macOS Intel (`x86_64-apple-darwin`), and Linux (for this host, `x86_64-unknown-linux-gnu`). End users do not need Python, Node, Rust, or a separately installed server. Native targets currently include Linux `.deb`/AppImage, macOS `.dmg`, and Windows NSIS. No signing or notarization is configured.

The Tauri identifier is currently the placeholder `com.example.musicprogramscheduler`. It MUST be replaced with the final organization-owned reverse-domain identifier before public beta. The development version remains `0.0.0`. A provisional Music Program Scheduler icon is generated from `src-tauri/icons/source.svg`; it is separate from the Vite favicon and should receive product-design review before public beta.

## Local Data and Network

Tauri stores `pianist_scheduling.db` under its per-user application-data directory. Electron continues to use Electron's `userData` directory, so test data and database changes are not shared between the two shells. Recovery archives live in the database directory's `recovery/` subdirectory; only the latest three are retained. The temporary backend binds only to `127.0.0.1`; its selected endpoint is passed to React through a Rust command. Tauri's official dialog and filesystem plugins are limited to open/save dialogs and file operations on dialog-selected paths. No external services are configured.

`.mpsession` archives are unencrypted and can include student educational information. Exported archives should be stored and transferred according to institutional policy. Export creates a portable copy; normal application edits are persisted automatically to the active local SQLite database.

The Accompanist availability CSV template is generated locally from synthetic examples. Tauri opens a native Save dialog and writes only to the chosen destination; canceling does not save elsewhere. Browser mode uses a local save picker when available and otherwise a conventional download.

## Schema Migration Tests

Run `npm test` before and after persistence/schema changes. The suite includes synthetic fixtures for the unversioned Accompanist schema, fresh/current database initialization, session metadata migration, `.mpsession` round trips, WAL snapshots, staged migrations, recovery behavior, and hostile/invalid archives. Do not use local production databases as test fixtures. The schema version and procedure for registering the next migration are documented in [ARCHITECTURE.md](ARCHITECTURE.md#sqlite-schema-migrations).
