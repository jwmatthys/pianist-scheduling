"""Database engine/session setup.

Uses SQLite for the local/MVP deployment. The schema is deliberately
multi-tenant-friendly (every top-level table hangs off an Organization)
so a future move to a shared server with per-customer accounts does not
require a data-model rewrite -- only adding real auth and scoping requests
by the authenticated organization.
"""

import os
from pathlib import Path

from sqlalchemy import create_engine, text
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker

DB_PATH = Path(
    os.environ.get("PIANIST_SCHEDULING_DB_PATH", Path(__file__).resolve().parent.parent / "pianist_scheduling.db")
).expanduser().resolve()
DB_PATH.parent.mkdir(parents=True, exist_ok=True)
DATABASE_URL = f"sqlite:///{DB_PATH}"

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


class Base(DeclarativeBase):
    pass


def get_db():
    db: Session = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def init_db():
    # Import models so they are registered on Base.metadata before create_all.
    from . import models  # noqa: F401

    Base.metadata.create_all(bind=engine)

    # Lightweight migration for installations created before instructor email
    # became an importable lesson field.
    with engine.begin() as connection:
        lesson_columns = {
            row[1] for row in connection.execute(text("PRAGMA table_info(lessons)"))
        }
        if "teacher_email" not in lesson_columns:
            connection.execute(text("ALTER TABLE lessons ADD COLUMN teacher_email VARCHAR(200) NOT NULL DEFAULT ''"))
        if "student_id" not in lesson_columns:
            connection.execute(text("ALTER TABLE lessons ADD COLUMN student_id VARCHAR(100) NOT NULL DEFAULT ''"))

    # Ensure a default organization exists (single-tenant MVP).
    from .models import Organization

    with SessionLocal() as db:
        if not db.query(Organization).first():
            db.add(Organization(name="Default Organization"))
            db.commit()
