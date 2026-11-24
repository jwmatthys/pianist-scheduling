from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from starlette.concurrency import run_in_threadpool
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import JSONResponse

from . import startup_profile
from .database import init_db
from .services.license_service import service as license_service

startup_profile.mark("Configuration loaded")
from .routers import (
    accompanist_lessons,
    accompanist_results,
    assignments,
    availability_imports,
    clinical,
    imports,
    license,
    jury,
    lesson_scheduling,
    pianists,
    reports,
    session,
    tutor_scheduling,
)

app = FastAPI(title="Scheduling Maestro API")

LOCAL_ORIGINS = [
    "http://127.0.0.1:5173",
    "http://localhost:5173",
    "tauri://localhost",
    "http://tauri.localhost",
    "null",
]


class TrialLockMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        allowed = {
            ("GET", "/api/health"),
            ("GET", "/api/license/status"),
            ("POST", "/api/license/trial-heartbeat"),
            ("POST", "/api/license/activate-file"),
            ("POST", "/api/license/activate-key"),
        }
        if request.url.path.startswith("/api/") and (request.method, request.url.path) not in allowed:
            status = await run_in_threadpool(license_service.trial_heartbeat)
            if status.status != "licensed" and status.trial_seconds_remaining == 0:
                return JSONResponse(
                    status_code=423,
                    content={"detail": {"code": "TRIAL_EXPIRED", "message": "The trial has expired. Activate a license to continue."}},
                )
        return await call_next(request)


app.add_middleware(TrialLockMiddleware)
app.add_middleware(
    CORSMiddleware,
    allow_origins=LOCAL_ORIGINS,
    allow_methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
    allow_headers=["Content-Type"],
)


@app.on_event("startup")
def on_startup():
    startup_profile.mark("Database initialization begin")
    startup_profile.emit_event("database-initialization-begin")
    try:
        init_db()
    except Exception:
        startup_profile.emit_event("database-initialization-failed")
        raise
    license_service.initialize()
    startup_profile.mark("Database initialization complete")
    startup_profile.mark("FastAPI startup complete")
    startup_profile.log_summary()


for _name, _router in (
    ("pianists", pianists),
    ("accompanist_lessons", accompanist_lessons),
    ("imports", imports),
    ("license", license),
    ("assignments", assignments),
    ("reports", reports),
    ("session", session),
    ("availability_imports", availability_imports),
    ("accompanist_results", accompanist_results),
    ("jury", jury),
    ("clinical", clinical),
    ("lesson_scheduling", lesson_scheduling),
    ("tutor_scheduling", tutor_scheduling),
):
    with startup_profile.measure(f"Router registration: {_name}"):
        app.include_router(_router.router)

startup_profile.mark("Router registration complete")


@app.get("/api/health")
def health():
    startup_profile.mark("Health endpoint first served")
    return {"status": "ok"}
