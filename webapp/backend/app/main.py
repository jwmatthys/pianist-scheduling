from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from .database import init_db
from .routers import assignments, imports, lessons, pianists, reports

app = FastAPI(title="Pianist Scheduling API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # local MVP; tighten when deployed multi-tenant
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
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
