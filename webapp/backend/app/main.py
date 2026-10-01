from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from .database import init_db
from .routers import assignments, imports, lessons, pianists, reports

app = FastAPI(title="Music Program Scheduler API")

LOCAL_ORIGINS = [
    "http://127.0.0.1:5173",
    "http://localhost:5173",
    "tauri://localhost",
    "http://tauri.localhost",
    "null",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=LOCAL_ORIGINS,
    allow_methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
    allow_headers=["Content-Type"],
)


@app.on_event("startup")
def on_startup():
    init_db()


app.include_router(pianists.router)
app.include_router(lessons.router)
app.include_router(imports.router)
app.include_router(assignments.router)
app.include_router(reports.router)


@app.get("/api/health")
def health():
    return {"status": "ok"}
