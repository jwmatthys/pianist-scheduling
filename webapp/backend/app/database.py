"""Database engine and migration setup for the active local Scheduling Session."""

import os
from threading import RLock
from dataclasses import dataclass
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Callable
from uuid import uuid4

from sqlalchemy import create_engine, text
from sqlalchemy.engine import Connection, Engine
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker

DB_PATH = Path(
    os.environ.get("PIANIST_SCHEDULING_DB_PATH", Path(__file__).resolve().parent.parent / "pianist_scheduling.db")
).expanduser().resolve()
DB_PATH.parent.mkdir(parents=True, exist_ok=True)
DATABASE_URL = f"sqlite:///{DB_PATH}"

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
DATABASE_LOCK = RLock()


class Base(DeclarativeBase):
    pass


def get_db():
    with DATABASE_LOCK:
        db: Session = SessionLocal()
        try:
            yield db
        finally:
            db.close()


def init_db():
    return migrate_database(engine)


class DatabaseState(str, Enum):
    FRESH = "fresh"
    LEGACY = "legacy"
    CURRENT = "current"
    VERSIONED_OLD = "versioned_old"


class DatabaseMigrationError(RuntimeError):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


@dataclass(frozen=True)
class SchemaMigration:
    version: int
    name: str
    upgrade: Callable[[Connection], None]


@dataclass(frozen=True)
class MigrationReport:
    initial_state: DatabaseState
    initial_version: int
    current_version: int
    applied_migrations: tuple[str, ...]


def _table_names(connection: Connection) -> set[str]:
    return set(connection.exec_driver_sql(
        "SELECT name FROM sqlite_master "
        "WHERE type = 'table' AND name NOT LIKE 'sqlite_%'"
    ).scalars())


def _column_names(connection: Connection, table_name: str) -> set[str]:
    escaped_name = table_name.replace('"', '""')
    return {
        row[1]
        for row in connection.exec_driver_sql(f'PRAGMA table_info("{escaped_name}")')
    }


def _expected_columns() -> dict[str, set[str]]:
    from . import models  # noqa: F401

    return {
        table_name: set(table.columns.keys())
        for table_name, table in Base.metadata.tables.items()
    }


def _validate_schema_shape(
    connection: Connection,
    *,
    allow_legacy_lesson_columns: bool,
) -> None:
    expected = _expected_columns()
    actual_tables = _table_names(connection)
    expected_tables = set(expected)
    unknown_tables = actual_tables - expected_tables
    if unknown_tables:
        raise DatabaseMigrationError(
            "UNRECOGNIZED_DATABASE",
            f"Database contains unrecognized tables: {', '.join(sorted(unknown_tables))}.",
        )

    missing_tables = expected_tables - actual_tables
    if missing_tables and not allow_legacy_lesson_columns:
        raise DatabaseMigrationError(
            "INVALID_DATABASE_SCHEMA",
            f"Database is missing required tables: {', '.join(sorted(missing_tables))}.",
        )
    if allow_legacy_lesson_columns and "lessons" not in actual_tables:
        raise DatabaseMigrationError(
            "UNRECOGNIZED_DATABASE",
            "Database has no lessons table and does not match a supported legacy schema.",
        )

    for table_name in sorted(actual_tables):
        expected_columns = expected[table_name]
        actual_columns = _column_names(connection, table_name)
        missing_columns = expected_columns - actual_columns
        if allow_legacy_lesson_columns and table_name == "lessons":
            missing_columns -= {"teacher_email", "student_id"}
        extra_columns = actual_columns - expected_columns
        if missing_columns or extra_columns:
            details = []
            if missing_columns:
                details.append(f"missing {', '.join(sorted(missing_columns))}")
            if extra_columns:
                details.append(f"unexpected {', '.join(sorted(extra_columns))}")
            raise DatabaseMigrationError(
                "INVALID_DATABASE_SCHEMA",
                f"Table '{table_name}' has an unsupported schema ({'; '.join(details)}).",
            )


def _classify_unversioned_database(connection: Connection) -> DatabaseState:
    tables = _table_names(connection)
    if not tables:
        return DatabaseState.FRESH
    _validate_schema_shape(connection, allow_legacy_lesson_columns=True)
    return DatabaseState.LEGACY


def _validate_current_schema(connection: Connection) -> None:
    _validate_schema_shape(connection, allow_legacy_lesson_columns=False)


def _upgrade_to_v1(connection: Connection) -> None:
    """Create the current Accompanist schema and absorb the supported v0 columns."""
    from . import models

    baseline_tables = [
        table for table in Base.metadata.sorted_tables
        if table.name != "scheduling_sessions"
    ]
    Base.metadata.create_all(bind=connection, tables=baseline_tables)

    lesson_columns = _column_names(connection, "lessons")
    legacy_columns = {
        "teacher_email": "VARCHAR(200) NOT NULL DEFAULT ''",
        "student_id": "VARCHAR(100) NOT NULL DEFAULT ''",
    }
    for column_name, column_definition in legacy_columns.items():
        if column_name not in lesson_columns:
            connection.exec_driver_sql(
                f'ALTER TABLE lessons ADD COLUMN "{column_name}" {column_definition}'
            )

    has_organization = connection.exec_driver_sql(
        "SELECT 1 FROM organizations LIMIT 1"
    ).first()
    if has_organization is None:
        connection.execute(
            models.Organization.__table__.insert().values(
                id=1,
                name="Default Organization",
            )
        )


def _upgrade_to_v2(connection: Connection) -> None:
    """Add session metadata without changing the released Accompanist baseline."""
    from . import models

    models.SchedulingSession.__table__.create(bind=connection)
    now = datetime.now(timezone.utc).replace(tzinfo=None)
    connection.execute(
        models.SchedulingSession.__table__.insert().values(
            session_uuid=str(uuid4()),
            institution_name="Legacy Session",
            program_name="Music Program",
            term_label="Term not set",
            year=None,
            start_date=None,
            end_date=None,
            created_at=now,
            modified_at=now,
        )
    )


MIGRATIONS = (
    SchemaMigration(
        version=1,
        name="create_current_schema_and_upgrade_legacy_lesson_columns",
        upgrade=_upgrade_to_v1,
    ),
    SchemaMigration(
        version=2,
        name="add_scheduling_session_metadata",
        upgrade=_upgrade_to_v2,
    ),
)
LATEST_SCHEMA_VERSION = MIGRATIONS[-1].version


def migrate_database(
    target_engine: Engine | None = None,
    migrations: tuple[SchemaMigration, ...] = MIGRATIONS,
) -> MigrationReport:
    """Upgrade an arbitrary SQLite database forward, including staged session copies."""
    migration_engine = target_engine or engine
    versions = [migration.version for migration in migrations]
    if versions != list(range(1, len(versions) + 1)):
        raise DatabaseMigrationError(
            "INVALID_MIGRATION_REGISTRY",
            "Registered migrations must be ordered and numbered consecutively from version 1.",
        )
    latest_version = versions[-1] if versions else 0

    try:
        with migration_engine.connect() as connection:
            try:
                connection.exec_driver_sql("BEGIN IMMEDIATE")
                initial_version = connection.exec_driver_sql("PRAGMA user_version").scalar_one()
                if initial_version > latest_version:
                    raise DatabaseMigrationError(
                        "UNSUPPORTED_SCHEMA_VERSION",
                        f"Database schema version {initial_version} is newer than the supported "
                        f"version {latest_version}; use a newer application version.",
                    )

                if initial_version == 0:
                    initial_state = _classify_unversioned_database(connection)
                elif initial_version == latest_version:
                    _validate_current_schema(connection)
                    initial_state = DatabaseState.CURRENT
                else:
                    initial_state = DatabaseState.VERSIONED_OLD

                current_version = initial_version
                applied_migrations = []
                migration_by_version = {migration.version: migration for migration in migrations}
                while current_version < latest_version:
                    next_version = current_version + 1
                    migration = migration_by_version.get(next_version)
                    if migration is None:
                        raise DatabaseMigrationError(
                            "INVALID_MIGRATION_REGISTRY",
                            f"No migration is registered for schema version {next_version}.",
                        )
                    try:
                        migration.upgrade(connection)
                    except Exception as error:
                        raise DatabaseMigrationError(
                            "MIGRATION_FAILED",
                            f"Migration {migration.version} ({migration.name}) failed: {error}",
                        ) from error
                    connection.exec_driver_sql(f"PRAGMA user_version = {migration.version}")
                    current_version = migration.version
                    applied_migrations.append(migration.name)

                _validate_current_schema(connection)
                connection.commit()
                return MigrationReport(
                    initial_state=initial_state,
                    initial_version=initial_version,
                    current_version=current_version,
                    applied_migrations=tuple(applied_migrations),
                )
            except Exception:
                connection.rollback()
                raise
    except DatabaseMigrationError:
        raise
    except Exception as error:
        raise DatabaseMigrationError(
            "INVALID_OR_INACCESSIBLE_DATABASE",
            f"Could not inspect or migrate the SQLite database: {error}",
        ) from error
