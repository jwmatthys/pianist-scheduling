"""Portable Scheduling Session archives and safe active-database replacement."""

import hashlib
import hmac
import json
import os
import sqlite3
import tempfile
import zipfile
from contextlib import closing
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path, PurePosixPath
from uuid import uuid4

from sqlalchemy import Engine
from sqlalchemy.orm import Session

from .. import models, schemas
from ..database import DATABASE_LOCK, LATEST_SCHEMA_VERSION, migrate_database

ARCHIVE_FORMAT = "music-program-scheduler-session"
ARCHIVE_VERSION = 1
DATABASE_ENTRY = "data/session.sqlite3"
MAX_ARCHIVE_BYTES = 512 * 1024 * 1024
MAX_DATABASE_BYTES = 512 * 1024 * 1024
MAX_MANIFEST_BYTES = 1024 * 1024
RECOVERY_RETENTION = 3


class SessionFileError(RuntimeError):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


def _database_path(target_engine: Engine) -> Path:
    raw_path = target_engine.url.database
    if not raw_path or raw_path == ":memory:":
        raise SessionFileError("UNSUPPORTED_DATABASE", "Session files require a file-backed SQLite database.")
    return Path(raw_path).resolve()


def _session_metadata(target_engine: Engine) -> models.SchedulingSession:
    with Session(target_engine) as db:
        rows = db.query(models.SchedulingSession).all()
        if len(rows) != 1:
            raise SessionFileError("INVALID_SESSION_METADATA", "The active database must contain exactly one session.")
        return rows[0]


def get_active_session(target_engine: Engine) -> schemas.SessionMetadataOut:
    with DATABASE_LOCK:
        return schemas.SessionMetadataOut.model_validate(_session_metadata(target_engine), from_attributes=True)


def update_session_metadata(
    target_engine: Engine,
    values: schemas.SessionMetadataIn,
) -> schemas.SessionMetadataOut:
    with DATABASE_LOCK:
        metadata = _metadata_input(values.model_dump())
        with Session(target_engine) as db:
            rows = db.query(models.SchedulingSession).all()
            if len(rows) != 1:
                raise SessionFileError("INVALID_SESSION_METADATA", "The active database must contain exactly one session.")
            row = rows[0]
            row.institution_name = metadata.institution_name
            row.program_name = metadata.program_name
            row.term_label = metadata.term_label
            row.year = metadata.year
            row.start_date = metadata.start_date
            row.end_date = metadata.end_date
            row.modified_at = datetime.now(timezone.utc).replace(tzinfo=None)
            db.commit()
            db.refresh(row)
            return schemas.SessionMetadataOut.model_validate(row, from_attributes=True)


def _metadata_input(values: dict) -> schemas.SessionMetadataIn:
    try:
        return schemas.SessionMetadataIn.model_validate(values)
    except Exception as error:
        raise SessionFileError("INVALID_SESSION_METADATA", str(error)) from error


def _replace_metadata(target_engine: Engine, values: schemas.SessionMetadataIn, session_uuid: str | None = None) -> None:
    from .. import module_models

    with Session(target_engine) as db:
        rows = db.query(models.SchedulingSession).all()
        if len(rows) != 1:
            raise SessionFileError("INVALID_SESSION_METADATA", "The staged database must contain exactly one session.")
        row = rows[0]
        previous_uuid = row.session_uuid
        row.session_uuid = session_uuid or str(uuid4())
        if row.session_uuid != previous_uuid:
            for table in module_models.ModuleBase.metadata.tables.values():
                if "session_uuid" in table.c:
                    db.execute(
                        table.update()
                        .where(table.c.session_uuid == previous_uuid)
                        .values(session_uuid=row.session_uuid)
                    )
        row.institution_name = values.institution_name.strip()
        row.program_name = values.program_name.strip()
        row.term_label = values.term_label.strip()
        row.year = values.year
        row.start_date = values.start_date
        row.end_date = values.end_date
        row.modified_at = datetime.now(timezone.utc).replace(tzinfo=None)
        db.commit()


def _manifest(target_engine: Engine, database_bytes: bytes) -> schemas.SessionArchiveManifest:
    session = _session_metadata(target_engine)
    with target_engine.connect() as connection:
        schema_version = connection.exec_driver_sql("PRAGMA user_version").scalar_one()
    return schemas.SessionArchiveManifest(
        format=ARCHIVE_FORMAT,
        formatVersion=ARCHIVE_VERSION,
        sessionId=session.session_uuid,
        institution=session.institution_name,
        program=session.program_name,
        term={
            "label": session.term_label,
            "year": session.year,
            "startDate": session.start_date,
            "endDate": session.end_date,
        },
        applicationVersion="0.0.0",
        databaseSchemaVersion=schema_version,
        exportedAt=datetime.now(timezone.utc),
        payload={"path": DATABASE_ENTRY, "sha256": hashlib.sha256(database_bytes).hexdigest()},
    )


def _snapshot_database(target_engine: Engine) -> bytes:
    source_path = _database_path(target_engine)
    if not source_path.is_file():
        raise SessionFileError("DATABASE_NOT_FOUND", "The active session database is unavailable.")
    with tempfile.TemporaryDirectory(dir=source_path.parent) as temporary_directory:
        snapshot_path = Path(temporary_directory) / "snapshot.sqlite3"
        try:
            with closing(sqlite3.connect(source_path)) as source, closing(sqlite3.connect(snapshot_path)) as destination:
                source.backup(destination)
                destination.execute("PRAGMA journal_mode=DELETE")
                destination.commit()
            snapshot = snapshot_path.read_bytes()
        except sqlite3.Error as error:
            raise SessionFileError("SNAPSHOT_FAILED", f"Could not snapshot the active database: {error}") from error
    if len(snapshot) > MAX_DATABASE_BYTES:
        raise SessionFileError("SESSION_TOO_LARGE", "The session database exceeds the supported archive size.")
    return snapshot


def export_session_archive(target_engine: Engine) -> bytes:
    with DATABASE_LOCK:
        database_bytes = _snapshot_database(target_engine)
        manifest = _manifest(target_engine, database_bytes)
        archive = BytesIO()
        with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED) as output:
            output.writestr("manifest.json", manifest.model_dump_json(by_alias=True))
            output.writestr(DATABASE_ENTRY, database_bytes)
        result = archive.getvalue()
        if len(result) > MAX_ARCHIVE_BYTES:
            raise SessionFileError("SESSION_TOO_LARGE", "The exported session archive exceeds the supported size.")
        return result


def _write_atomic(path: Path, contents: bytes) -> None:
    temporary_path: Path | None = None
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=path.parent, prefix=f".{path.name}.", delete=False) as target:
            temporary_path = Path(target.name)
            target.write(contents)
            target.flush()
            os.fsync(target.fileno())
        os.replace(temporary_path, path)
    except OSError as error:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
        raise SessionFileError("FILE_WRITE_FAILED", f"Could not safely write the session file: {error}") from error


def _recovery_directory(target_engine: Engine) -> Path:
    return _database_path(target_engine).parent / "recovery"


def create_recovery_snapshot(target_engine: Engine, directory: Path | None = None) -> Path:
    with DATABASE_LOCK:
        archive_bytes = export_session_archive(target_engine)
        recovery_directory = directory or _recovery_directory(target_engine)
        timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        path = recovery_directory / f"recovery-{timestamp}-{uuid4().hex[:8]}.mpsession"
        _write_atomic(path, archive_bytes)
        try:
            manifest, database_bytes = _validate_archive(path.read_bytes())
            _check_database_payload(database_bytes, manifest)
            _validate_session_matches(target_engine, manifest)
        except Exception as error:
            path.unlink(missing_ok=True)
            if isinstance(error, SessionFileError):
                raise SessionFileError(
                    "RECOVERY_VERIFICATION_FAILED",
                    f"The recovery snapshot could not be verified: {error}",
                ) from error
            raise SessionFileError(
                "RECOVERY_VERIFICATION_FAILED",
                f"The recovery snapshot could not be verified: {error}",
            ) from error
        snapshots = sorted(
            recovery_directory.glob("recovery-*.mpsession"),
            key=lambda item: item.stat().st_mtime_ns,
            reverse=True,
        )
        for expired in snapshots[RECOVERY_RETENTION:]:
            expired.unlink(missing_ok=True)
        return path


def _validate_archive(archive_bytes: bytes) -> tuple[schemas.SessionArchiveManifest, bytes]:
    if len(archive_bytes) > MAX_ARCHIVE_BYTES:
        raise SessionFileError("SESSION_TOO_LARGE", "The selected session archive exceeds the supported size.")
    try:
        archive = zipfile.ZipFile(BytesIO(archive_bytes))
    except (zipfile.BadZipFile, OSError) as error:
        raise SessionFileError("INVALID_ARCHIVE", "The selected file is not a valid session archive.") from error

    with archive:
        entries = archive.infolist()
        names = [entry.filename for entry in entries]
        for name in names:
            path = PurePosixPath(name)
            if "\\" in name or path.is_absolute() or ".." in path.parts:
                raise SessionFileError("UNSAFE_ARCHIVE_PATH", "The archive contains an unsafe file path.")
        if len(names) != len(set(names)):
            raise SessionFileError("DUPLICATE_ARCHIVE_ENTRY", "The archive contains duplicate entries.")
        if set(names) != {"manifest.json", DATABASE_ENTRY}:
            raise SessionFileError("INVALID_ARCHIVE_STRUCTURE", "The archive must contain only its manifest and database payload.")
        by_name = {entry.filename: entry for entry in entries}
        if any(entry.flag_bits & 1 for entry in entries):
            raise SessionFileError("ENCRYPTED_ARCHIVE_UNSUPPORTED", "Encrypted session archives are not supported.")
        if by_name["manifest.json"].file_size > MAX_MANIFEST_BYTES:
            raise SessionFileError("SESSION_TOO_LARGE", "The archive manifest exceeds the supported size.")
        if by_name[DATABASE_ENTRY].file_size > MAX_DATABASE_BYTES:
            raise SessionFileError("SESSION_TOO_LARGE", "The session database exceeds the supported size.")
        try:
            manifest_raw = _read_zip_entry(archive, "manifest.json", MAX_MANIFEST_BYTES)
            manifest_json = json.loads(manifest_raw)
            if manifest_json.get("format") != ARCHIVE_FORMAT:
                raise SessionFileError("UNSUPPORTED_ARCHIVE_FORMAT", "This file is not a supported Music Program Scheduler session archive.")
            if manifest_json.get("formatVersion") != ARCHIVE_VERSION:
                raise SessionFileError("UNSUPPORTED_ARCHIVE_VERSION", f"Archive format version {manifest_json.get('formatVersion')} is not supported.")
            manifest = schemas.SessionArchiveManifest.model_validate(manifest_json)
            database_bytes = _read_zip_entry(archive, DATABASE_ENTRY, MAX_DATABASE_BYTES)
        except SessionFileError:
            raise
        except Exception as error:
            raise SessionFileError("INVALID_MANIFEST", f"The session manifest is invalid: {error}") from error
        if not hmac.compare_digest(hashlib.sha256(database_bytes).hexdigest(), manifest.payload.sha256):
            raise SessionFileError("PAYLOAD_CHECKSUM_MISMATCH", "The session database payload does not match its manifest checksum.")
        return manifest, database_bytes


def _read_zip_entry(archive: zipfile.ZipFile, name: str, maximum: int) -> bytes:
    contents = bytearray()
    with archive.open(name) as source:
        while chunk := source.read(min(1024 * 1024, maximum + 1 - len(contents))):
            contents.extend(chunk)
            if len(contents) > maximum:
                raise SessionFileError("SESSION_TOO_LARGE", f"Archive entry '{name}' exceeds the supported size.")
    return bytes(contents)


def _check_database_payload(database_bytes: bytes, manifest: schemas.SessionArchiveManifest) -> int:
    if not database_bytes:
        raise SessionFileError("INVALID_DATABASE", "The session database payload is empty.")
    with tempfile.NamedTemporaryFile(suffix=".sqlite3") as staged_file:
        staged_file.write(database_bytes)
        staged_file.flush()
        try:
            with closing(sqlite3.connect(f"file:{staged_file.name}?mode=ro", uri=True)) as connection:
                integrity = connection.execute("PRAGMA integrity_check").fetchone()
                version = connection.execute("PRAGMA user_version").fetchone()[0]
        except sqlite3.Error as error:
            raise SessionFileError("INVALID_DATABASE", "The archive does not contain a readable SQLite database.") from error
    if integrity != ("ok",):
        raise SessionFileError("INVALID_DATABASE", "The session database failed SQLite integrity validation.")
    if version != manifest.database_schema_version:
        raise SessionFileError("SCHEMA_VERSION_MISMATCH", "The database schema version does not match the archive manifest.")
    if version > LATEST_SCHEMA_VERSION:
        raise SessionFileError("UNSUPPORTED_SCHEMA_VERSION", f"Database schema version {version} is newer than supported version {LATEST_SCHEMA_VERSION}.")
    return version


def _apply_manifest_metadata(target_engine: Engine, manifest: schemas.SessionArchiveManifest) -> None:
    try:
        metadata = schemas.SessionMetadataIn(
            institution_name=manifest.institution,
            program_name=manifest.program,
            term_label=manifest.term.label,
            year=manifest.term.year,
            start_date=manifest.term.startDate,
            end_date=manifest.term.endDate,
        )
        _replace_metadata(target_engine, metadata, str(manifest.session_id))
    except Exception as error:
        raise SessionFileError("INVALID_SESSION_METADATA", f"Session metadata is invalid: {error}") from error


def _validate_session_matches(target_engine: Engine, manifest: schemas.SessionArchiveManifest) -> None:
    session = _session_metadata(target_engine)
    expected = (
        str(manifest.session_id), manifest.institution, manifest.program, manifest.term.label,
        manifest.term.year, manifest.term.startDate, manifest.term.endDate,
    )
    actual = (
        session.session_uuid, session.institution_name, session.program_name, session.term_label,
        session.year, session.start_date, session.end_date,
    )
    if actual != expected:
        raise SessionFileError("SESSION_METADATA_MISMATCH", "Database session metadata does not match the archive manifest.")


def _activate_staged_database(target_engine: Engine, staged_path: Path) -> None:
    active_path = _database_path(target_engine)
    target_engine.dispose()
    try:
        with closing(sqlite3.connect(active_path)) as active_database:
            checkpoint = active_database.execute("PRAGMA wal_checkpoint(TRUNCATE)").fetchone()
            if checkpoint and checkpoint[0] != 0:
                raise SessionFileError("DATABASE_BUSY", "The active database could not be safely closed for replacement.")
        Path(f"{active_path}-wal").unlink(missing_ok=True)
        Path(f"{active_path}-shm").unlink(missing_ok=True)
        os.replace(staged_path, active_path)
    except SessionFileError:
        raise
    except OSError as error:
        raise SessionFileError("SESSION_REPLACEMENT_FAILED", f"Could not activate the staged session: {error}") from error


def create_new_session(
    target_engine: Engine,
    values: schemas.SessionMetadataIn,
    recovery_directory: Path | None = None,
) -> schemas.SessionMetadataOut:
    with DATABASE_LOCK:
        metadata = _metadata_input(values.model_dump())
        create_recovery_snapshot(target_engine, recovery_directory)
        active_path = _database_path(target_engine)
        active_path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=active_path.parent) as temporary_directory:
            staged_path = Path(temporary_directory) / "new-session.sqlite3"
            from sqlalchemy import create_engine

            staged_engine = create_engine(f"sqlite:///{staged_path}", connect_args={"check_same_thread": False})
            try:
                migrate_database(staged_engine)
                _replace_metadata(staged_engine, metadata)
            except Exception as error:
                staged_engine.dispose()
                if isinstance(error, SessionFileError):
                    raise
                raise SessionFileError("SESSION_INITIALIZATION_FAILED", f"Could not initialize the new session: {error}") from error
            staged_engine.dispose()
            _activate_staged_database(target_engine, staged_path)
        return schemas.SessionMetadataOut.model_validate(_session_metadata(target_engine), from_attributes=True)


def restore_session(
    target_engine: Engine,
    archive_bytes: bytes,
    recovery_directory: Path | None = None,
) -> schemas.SessionMetadataOut:
    with DATABASE_LOCK:
        manifest, database_bytes = _validate_archive(archive_bytes)
        initial_schema_version = _check_database_payload(database_bytes, manifest)
        active_path = _database_path(target_engine)
        active_path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=active_path.parent) as temporary_directory:
            staged_path = Path(temporary_directory) / "restored-session.sqlite3"
            staged_path.write_bytes(database_bytes)
            from sqlalchemy import create_engine

            staged_engine = create_engine(f"sqlite:///{staged_path}", connect_args={"check_same_thread": False})
            try:
                migrate_database(staged_engine)
                if initial_schema_version < 2:
                    _apply_manifest_metadata(staged_engine, manifest)
                _validate_session_matches(staged_engine, manifest)
            except Exception as error:
                staged_engine.dispose()
                if isinstance(error, SessionFileError):
                    raise
                code = getattr(error, "code", "STAGED_MIGRATION_FAILED")
                raise SessionFileError(code, f"Could not validate or migrate the staged session: {error}") from error
            staged_engine.dispose()
            create_recovery_snapshot(target_engine, recovery_directory)
            _activate_staged_database(target_engine, staged_path)
        return schemas.SessionMetadataOut.model_validate(_session_metadata(target_engine), from_attributes=True)