"""Database engine/session setup.

Uses SQLite for the local/MVP deployment. The schema is deliberately
multi-tenant-friendly (every top-level table hangs off an Organization)
so a future move to a shared server with per-customer accounts does not
require a data-model rewrite -- only adding real auth and scoping requests
by the authenticated organization.
"""

from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker

DB_PATH = Path(__file__).resolve().parent.parent / "pianist_scheduling.db"
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

    # Ensure a default organization exists (single-tenant MVP).
    from .models import Organization

    with SessionLocal() as db:
        if not db.query(Organization).first():
            db.add(Organization(name="Default Organization"))
            db.commit()
