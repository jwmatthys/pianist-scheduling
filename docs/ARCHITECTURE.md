# Architecture

## Product and Module Boundaries

The product is **Music Program Scheduler**. Accompanist Scheduling is its first module. The current import, pianist availability, assignment, and report views remain inside that module; no future module workflows are introduced here.

Academic terms, small people identities, time primitives, locations, and useful file/import infrastructure may be shared when actual reuse is demonstrated. Lesson requirements, pianist availability and workload, accompanist assignment scoring, jury rules, and clinical placement constraints remain module-owned. The accompanist optimizer is not a universal scheduling engine.

## Transitional Desktop Runtime

```mermaid
flowchart LR
    Tauri[Tauri 2 shell] --> React[Existing React / TypeScript UI]
    React -->|REST over 127.0.0.1 and ephemeral port| API[Packaged FastAPI / uvicorn service]
    API --> SQLAlchemy[Existing SQLAlchemy services]
    SQLAlchemy --> SQLite[(Local SQLite database)]
    Tauri -->|spawn, health-check, terminate| API
```

Tauri currently replaces Electron only as the desktop shell. During development Rust starts `desktop_server.py` with the backend virtualenv (or the configured Python interpreter); production starts the bundled PyInstaller service. Scheduling, imports, reports, and SQLite remain in Python. This is a transitional runtime, not the intended long-term application boundary.

The service binds only to `127.0.0.1`. Tauri reserves an ephemeral port, passes it to the service, waits for `/api/health`, and exposes `http://127.0.0.1:<port>` to React through the narrow `get_api_base` Tauri command. `src/lib/platform.ts` is the only frontend module that detects Tauri or invokes that command. Browser development continues to use `VITE_API_BASE`, then the legacy `apiBase` query parameter, then `http://localhost:8123`.

FastAPI CORS allows only the Vite development origins, the Tauri WebView origins (`tauri://localhost` and `http://tauri.localhost`), and Electron's opaque `null` file origin. Credentials are not enabled; only the API methods and `Content-Type` header used by the app are allowed. CORS does not expose the service beyond its loopback listener.

The service's database path is set to Tauri's per-user application-data directory as `pianist_scheduling.db`. Electron remains available for comparison and continues to use Electron's `userData` path; the two shells therefore have separate local databases and do not concurrently open the same SQLite file. No data is uploaded or synchronized between them.

On Unix the service is launched in a dedicated process group, which is terminated on exit so PyInstaller one-file workers cannot survive their parent. On Windows Tauri invokes the fixed system `taskkill.exe /PID <pid> /T /F` operation to terminate the child tree, with direct-child termination as a fallback. Tauri then waits for the child. Startup failure also terminates the child before returning an error. Unix process-group termination was smoke-tested with the packaged service; Windows tree termination still needs a native Windows validation run.

## Platform Abstraction

React's API client depends on `getApiBase()` in `src/lib/platform.ts`. The Tauri adapter obtains the endpoint from a Rust command. Browser builds retain their existing environment/query configuration. No Tauri APIs are called from React pages or components.

The Tauri capability for the main window grants no plugin permissions. There is no shell, filesystem, opener, updater, or network plugin exposed to the frontend. The custom endpoint command is registered by the Rust application.

## Privacy and Network

Project and scheduling data remain local. The only new app traffic is React-to-FastAPI HTTP on loopback. There are no telemetry, analytics, cloud, licensing, or remote-service integrations. Tauri does not open a LAN listener. Future location/travel scoring and other module-specific algorithms remain out of scope.
