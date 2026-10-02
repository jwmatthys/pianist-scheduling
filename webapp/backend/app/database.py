"""Database engine and migration setup for the active local Scheduling Session."""

import os
from contextlib import contextmanager
from threading import Condition, RLock
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
DATABASE_QUIESCENCE = Condition(DATABASE_LOCK)
ACTIVE_DATABASE_SESSIONS = 0


class Base(DeclarativeBase):
    pass


def get_db():
    global ACTIVE_DATABASE_SESSIONS
    with DATABASE_QUIESCENCE:
        ACTIVE_DATABASE_SESSIONS += 1
    try:
        db: Session = SessionLocal()
        try:
            yield db
        finally:
            db.close()
    finally:
        with DATABASE_QUIESCENCE:
            ACTIVE_DATABASE_SESSIONS -= 1
            DATABASE_QUIESCENCE.notify_all()


@contextmanager
def database_operation_lock():
    """Exclude session-file operations until all request-scoped DB sessions close."""
    with DATABASE_QUIESCENCE:
        while ACTIVE_DATABASE_SESSIONS:
            DATABASE_QUIESCENCE.wait()
        yield


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
    from . import module_models  # noqa: F401

    expected = {
        table_name: set(table.columns.keys())
        for table_name, table in Base.metadata.tables.items()
    }
    expected.update({
        table_name: set(table.columns.keys())
        for table_name, table in module_models.ModuleBase.metadata.tables.items()
    })
    return expected


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
        if allow_legacy_lesson_columns and table_name == "pianists":
            missing_columns -= {"pianist_code"}
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
        if table.name not in {"scheduling_sessions", "pianist_availability_states"}
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


def _upgrade_to_v3(connection: Connection) -> None:
    """Track Accompanist weekly completeness separately from availability status."""
    from . import models

    models.PianistAvailabilityState.__table__.create(bind=connection)


def _create_module_tables(connection: Connection, table_names: set[str]) -> None:
    from . import module_models

    tables = [
        table for table in module_models.ModuleBase.metadata.sorted_tables
        if table.name in table_names
    ]
    for table in tables:
        table.create(bind=connection)


def _upgrade_to_v4(connection: Connection) -> None:
    """Add UUID identities, Accompanist source mappings, and source revision state."""
    from . import module_models

    _create_module_tables(connection, {
        "person_identities",
        "accompanist_student_profiles",
        "accompanist_pianist_identities",
        "accompanist_lesson_identities",
        "module_revisions",
    })
    session_rows = connection.exec_driver_sql(
        "SELECT session_uuid FROM scheduling_sessions"
    ).all()
    if len(session_rows) != 1:
        raise DatabaseMigrationError(
            "INVALID_SESSION_METADATA",
            "Identity migration requires exactly one active Scheduling Session.",
        )
    session_uuid = session_rows[0][0]
    now = datetime.now(timezone.utc).replace(tzinfo=None)

    for row in connection.exec_driver_sql("SELECT id, name FROM pianists ORDER BY id").all():
        person_uuid = str(uuid4())
        connection.execute(module_models.PersonIdentity.__table__.insert().values(
            person_uuid=person_uuid,
            display_name=row[1] or "",
            created_at=now,
        ))
        connection.execute(module_models.AccompanistPianistIdentity.__table__.insert().values(
            pianist_id=row[0],
            person_uuid=person_uuid,
        ))

    lesson_rows = connection.exec_driver_sql(
        "SELECT id, student_id, student FROM lessons ORDER BY id"
    ).all()
    grouped_lessons: dict[str, list[tuple[int, str | None, str | None]]] = {}
    for lesson_id, raw_student_id, student_name in lesson_rows:
        normalized_id = module_models.normalize_student_id(raw_student_id)
        if normalized_id:
            grouped_lessons.setdefault(normalized_id, []).append(
                (lesson_id, raw_student_id, student_name)
            )

    identities_by_student_id: dict[str, str] = {}
    for student_id, rows in grouped_lessons.items():
        normalized_names = {
            module_models.normalize_student_name(row[2])
            for row in rows
            if module_models.normalize_student_name(row[2])
        }
        person_uuid = str(uuid4())
        display_name = next(
            (row[2] or "" for row in rows if (row[2] or "").strip()),
            "",
        )
        connection.execute(module_models.PersonIdentity.__table__.insert().values(
            person_uuid=person_uuid,
            display_name=display_name,
            created_at=now,
        ))
        connection.execute(module_models.AccompanistStudentProfile.__table__.insert().values(
            person_uuid=person_uuid,
            session_uuid=session_uuid,
            student_id=student_id,
            normalized_student_id=student_id,
            seen_names_json=module_models.normalized_names_json(normalized_names),
            identity_conflict=len(normalized_names) > 1,
        ))
        identities_by_student_id[student_id] = person_uuid

    for lesson_id, raw_student_id, student_name in lesson_rows:
        student_id = module_models.normalize_student_id(raw_student_id)
        if student_id:
            person_uuid = identities_by_student_id[student_id]
        else:
            person_uuid = str(uuid4())
            normalized_name = module_models.normalize_student_name(student_name)
            connection.execute(module_models.PersonIdentity.__table__.insert().values(
                person_uuid=person_uuid,
                display_name=student_name or "",
                created_at=now,
            ))
            connection.execute(module_models.AccompanistStudentProfile.__table__.insert().values(
                person_uuid=person_uuid,
                session_uuid=session_uuid,
                student_id=None,
                normalized_student_id=None,
                seen_names_json=module_models.normalized_names_json(
                    {normalized_name} if normalized_name else set()
                ),
                identity_conflict=False,
            ))
        connection.execute(module_models.AccompanistLessonIdentity.__table__.insert().values(
            lesson_id=lesson_id,
            lesson_uuid=str(uuid4()),
            student_person_uuid=person_uuid,
        ))

    connection.execute(module_models.ModuleRevision.__table__.insert().values(
        session_uuid=session_uuid,
        module_id="accompanists",
        source_revision=0,
        current_result_uuid=None,
        modified_at=now,
    ))


def _upgrade_to_v5(connection: Connection) -> None:
    """Add typed result envelopes and dependency provenance."""
    _create_module_tables(connection, {
        "module_results",
        "module_result_dependencies",
    })


def _upgrade_to_v6(connection: Connection) -> None:
    """Add Jury configuration, lesson entries, panels, and date-scoped availability."""
    from . import module_models

    _create_module_tables(connection, {
        "jury_configurations",
        "jury_panels",
        "jury_lesson_entries",
        "jury_pianist_availability_declarations",
        "jury_pianist_available_windows",
    })
    sessions = connection.exec_driver_sql(
        "SELECT session_uuid FROM scheduling_sessions"
    ).all()
    now = datetime.now(timezone.utc).replace(tzinfo=None)
    for (session_uuid,) in sessions:
        connection.execute(module_models.JuryConfiguration.__table__.insert().values(
            session_uuid=session_uuid,
            jury_date=None,
            input_revision=0,
            roster_source_result_uuid=None,
            modified_at=now,
        ))


def _upgrade_to_v7(connection: Connection) -> None:
    """Add Accompanist-owned Jury Required facts with a conservative false default."""
    from . import module_models

    _create_module_tables(connection, {"accompanist_lesson_jury_requirements"})
    connection.exec_driver_sql(
        "INSERT INTO accompanist_lesson_jury_requirements (lesson_id, jury_required) "
        "SELECT id, 0 FROM lessons"
    )
    connection.exec_driver_sql(
        "UPDATE module_results SET state = 'superseded' "
        "WHERE result_uuid IN ("
        "SELECT current_result_uuid FROM module_revisions "
        "WHERE module_id = 'accompanists' AND current_result_uuid IS NOT NULL"
        ") AND contract_id = 'accompanist.assignment-result' AND contract_version = 1"
    )
    connection.exec_driver_sql(
        "UPDATE module_revisions SET current_result_uuid = NULL "
        "WHERE module_id = 'accompanists' AND current_result_uuid IN ("
        "SELECT result_uuid FROM module_results "
        "WHERE contract_id = 'accompanist.assignment-result' AND contract_version = 1"
        ")"
    )


def _upgrade_to_v8(connection: Connection) -> None:
    """Add Panel-scoped Jury Dates and preserve the legacy session date where set."""
    from . import module_models

    module_models.JuryPanelDate.__table__.create(bind=connection)
    connection.exec_driver_sql(
        "INSERT INTO jury_panel_dates (panel_uuid, session_uuid, jury_date) "
        "SELECT jury_panels.panel_uuid, jury_panels.session_uuid, jury_configurations.jury_date "
        "FROM jury_panels JOIN jury_configurations "
        "ON jury_configurations.session_uuid = jury_panels.session_uuid "
        "WHERE jury_configurations.jury_date IS NOT NULL"
    )


def _upgrade_to_v9(connection: Connection) -> None:
    """Retain the legacy Pianist code column for compatibility."""
    if "pianist_code" not in _column_names(connection, "pianists"):
        connection.exec_driver_sql("ALTER TABLE pianists ADD COLUMN pianist_code VARCHAR(100)")
    connection.exec_driver_sql(
        "UPDATE pianists SET pianist_code = CAST(id AS TEXT) WHERE pianist_code IS NULL"
    )
    connection.exec_driver_sql(
        "CREATE UNIQUE INDEX IF NOT EXISTS uq_pianists_organization_code_ci "
        "ON pianists(organization_id, lower(pianist_code))"
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
    SchemaMigration(
        version=3,
        name="add_pianist_availability_completeness_state",
        upgrade=_upgrade_to_v3,
    ),
    SchemaMigration(
        version=4,
        name="add_shared_identities_and_accompanist_source_revision",
        upgrade=_upgrade_to_v4,
    ),
    SchemaMigration(
        version=5,
        name="add_typed_module_results_and_dependencies",
        upgrade=_upgrade_to_v5,
    ),
    SchemaMigration(
        version=6,
        name="add_jury_configuration_and_inputs",
        upgrade=_upgrade_to_v6,
    ),
    SchemaMigration(
        version=7,
        name="add_accompanist_lesson_jury_requirement",
        upgrade=_upgrade_to_v7,
    ),
    SchemaMigration(
        version=8,
        name="scope_jury_date_to_panels",
        upgrade=_upgrade_to_v8,
    ),
    SchemaMigration(
        version=9,
        name="add_editable_accompanist_pianist_ids",
        upgrade=_upgrade_to_v9,
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
