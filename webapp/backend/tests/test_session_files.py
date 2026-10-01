import hashlib
import json
import sqlite3
import tempfile
import unittest
import warnings
import zipfile
from datetime import date
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import models, schemas
from app.database import LATEST_SCHEMA_VERSION, MIGRATIONS, migrate_database
from app.services.session_files import (
    DATABASE_ENTRY,
    RECOVERY_RETENTION,
    SessionFileError,
    create_new_session,
    create_recovery_snapshot,
    export_session_archive,
    get_active_session,
    restore_session,
    update_session_metadata,
)


class SessionFileTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.directory = Path(self.temporary_directory.name)
        self.database_path = self.directory / "active.sqlite3"
        self.engine = create_engine(f"sqlite:///{self.database_path}", connect_args={"check_same_thread": False})
        migrate_database(self.engine)
        self.recovery_directory = self.directory / "recovery"

    def tearDown(self):
        self.engine.dispose()
        self.temporary_directory.cleanup()

    def metadata(self, institution="Synthetic University", program="Music", term="Fall 2026"):
        return schemas.SessionMetadataIn(
            institution_name=institution,
            program_name=program,
            term_label=term,
            year=2026,
            start_date=date(2026, 8, 20),
            end_date=date(2026, 12, 20),
        )

    def populate(self, name="Synthetic Pianist"):
        with Session(self.engine) as db:
            db.add(models.Pianist(name=name, email="pianist@example.invalid"))
            db.add(models.Lesson(
                teacher="Synthetic Instructor",
                student="Synthetic Student",
                day="Tuesday",
                start_minute=600,
                end_minute=650,
            ))
            db.commit()

    def archive_entries(self, archive_bytes):
        with zipfile.ZipFile(BytesIO(archive_bytes)) as archive:
            return {item.filename: archive.read(item.filename) for item in archive.infolist()}

    def make_archive(self, database_bytes, schema_version=3, manifest_overrides=None, filenames=None):
        manifest = {
            "format": "music-program-scheduler-session",
            "formatVersion": 1,
            "sessionId": "46e0ca39-11d0-4f07-9326-26c5a1ee349f",
            "institution": "Synthetic University",
            "program": "Music",
            "term": {
                "label": "Fall 2026",
                "year": 2026,
                "startDate": "2026-08-20",
                "endDate": "2026-12-20",
            },
            "applicationVersion": "0.0.0",
            "databaseSchemaVersion": schema_version,
            "exportedAt": "2026-06-01T00:00:00Z",
            "payload": {
                "path": DATABASE_ENTRY,
                "sha256": hashlib.sha256(database_bytes).hexdigest(),
            },
        }
        manifest.update(manifest_overrides or {})
        output = BytesIO()
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, contents in filenames or [("manifest.json", json.dumps(manifest).encode()), (DATABASE_ENTRY, database_bytes)]:
                archive.writestr(name, contents)
        return output.getvalue()

    def test_new_session_has_uuid_metadata_and_no_prior_accompanist_data(self):
        self.populate()

        created = create_new_session(self.engine, self.metadata(), self.recovery_directory)

        self.assertEqual(created.institution_name, "Synthetic University")
        self.assertEqual(created.program_name, "Music")
        self.assertEqual(created.term_label, "Fall 2026")
        self.assertEqual(created.year, 2026)
        self.assertIsNotNone(created.session_uuid)
        self.assertEqual(created, get_active_session(self.engine))
        with Session(self.engine) as db:
            self.assertEqual(db.query(models.Pianist).count(), 0)
            self.assertEqual(db.query(models.Lesson).count(), 0)
            self.assertEqual(db.query(models.SchedulingSession).count(), 1)
        self.assertEqual(len(list(self.recovery_directory.glob("*.mpsession"))), 1)

    def test_optional_metadata_is_persisted_as_null(self):
        metadata = schemas.SessionMetadataIn(
            institution_name="Synthetic University",
            program_name="Music",
            term_label="Custom Term",
        )

        created = create_new_session(self.engine, metadata, self.recovery_directory)

        self.assertIsNone(created.year)
        self.assertIsNone(created.start_date)
        self.assertIsNone(created.end_date)

    def test_existing_session_metadata_can_be_edited_without_replacing_data_or_identity(self):
        created = create_new_session(self.engine, self.metadata(), self.recovery_directory)
        self.populate()
        edited = update_session_metadata(
            self.engine,
            schemas.SessionMetadataIn(
                institution_name="Edited Synthetic University",
                program_name="Piano Studies",
                term_label="Fall 2026",
                year=2026,
            ),
        )

        self.assertEqual(edited.session_uuid, created.session_uuid)
        self.assertGreaterEqual(edited.modified_at, created.modified_at)
        self.assertEqual(edited.institution_name, "Edited Synthetic University")
        with Session(self.engine) as db:
            self.assertEqual(db.query(models.Pianist).count(), 1)
            self.assertEqual(db.query(models.Lesson).count(), 1)

    def test_recovery_snapshot_failure_keeps_active_database_unchanged(self):
        self.populate()
        with Session(self.engine) as db:
            original_uuid = db.query(models.SchedulingSession).one().session_uuid
            original_count = db.query(models.Pianist).count()
        blocker = self.directory / "not-a-directory"
        blocker.write_text("synthetic", encoding="utf-8")

        with self.assertRaises(SessionFileError):
            create_new_session(self.engine, self.metadata(), blocker / "recovery")

        with Session(self.engine) as db:
            self.assertEqual(db.query(models.SchedulingSession).one().session_uuid, original_uuid)
            self.assertEqual(db.query(models.Pianist).count(), original_count)

    def test_export_manifest_and_payload_include_data_without_modifying_active_session(self):
        self.populate()
        before = get_active_session(self.engine)

        archive_bytes = export_session_archive(self.engine)
        entries = self.archive_entries(archive_bytes)
        manifest = json.loads(entries["manifest.json"])

        self.assertEqual(set(entries), {"manifest.json", DATABASE_ENTRY})
        self.assertEqual(manifest["formatVersion"], 1)
        self.assertEqual(manifest["databaseSchemaVersion"], LATEST_SCHEMA_VERSION)
        self.assertEqual(manifest["sessionId"], str(before.session_uuid))
        self.assertEqual(manifest["payload"]["sha256"], hashlib.sha256(entries[DATABASE_ENTRY]).hexdigest())
        with sqlite3.connect(":memory:") as connection:
            connection.deserialize(entries[DATABASE_ENTRY])
            self.assertEqual(connection.execute("SELECT name FROM pianists").fetchone()[0], "Synthetic Pianist")
        self.assertEqual(get_active_session(self.engine), before)

    def test_export_captures_wal_enabled_database_state(self):
        self.populate()
        with self.engine.connect() as connection:
            connection.exec_driver_sql("PRAGMA journal_mode=WAL")
        with Session(self.engine) as db:
            db.add(models.Pianist(name="WAL Synthetic Pianist"))
            db.commit()

        entries = self.archive_entries(export_session_archive(self.engine))

        with sqlite3.connect(":memory:") as snapshot:
            snapshot.deserialize(entries[DATABASE_ENTRY])
            names = {row[0] for row in snapshot.execute("SELECT name FROM pianists")}
        self.assertEqual(names, {"Synthetic Pianist", "WAL Synthetic Pianist"})

    def test_restore_round_trip_recovers_metadata_and_accompanist_data(self):
        original = create_new_session(self.engine, self.metadata(), self.recovery_directory)
        self.populate()
        exported = export_session_archive(self.engine)
        create_new_session(self.engine, self.metadata("Other University", "Voice", "Spring 2027"), self.recovery_directory)

        restored = restore_session(self.engine, exported, self.recovery_directory)

        self.assertEqual(restored.session_uuid, original.session_uuid)
        self.assertEqual(restored.institution_name, "Synthetic University")
        with Session(self.engine) as db:
            self.assertEqual([row.name for row in db.query(models.Pianist)], ["Synthetic Pianist"])
            self.assertEqual(db.query(models.Lesson).one().student, "Synthetic Student")
        self.assertGreaterEqual(len(list(self.recovery_directory.glob("*.mpsession"))), 2)

    def test_schema_one_archive_migrates_only_in_staging(self):
        schema_one_path = self.directory / "schema-one.sqlite3"
        staged_engine = create_engine(f"sqlite:///{schema_one_path}")
        with staged_engine.begin() as connection:
            MIGRATIONS[0].upgrade(connection)
            connection.exec_driver_sql("PRAGMA user_version = 1")
        staged_engine.dispose()
        old_archive = self.make_archive(schema_one_path.read_bytes(), schema_version=1)
        active_before = get_active_session(self.engine)

        restored = restore_session(self.engine, old_archive, self.recovery_directory)

        self.assertEqual(str(restored.session_uuid), "46e0ca39-11d0-4f07-9326-26c5a1ee349f")
        self.assertEqual(get_active_session(self.engine).institution_name, "Synthetic University")
        self.assertNotEqual(active_before.session_uuid, restored.session_uuid)
        with self.engine.connect() as connection:
            self.assertEqual(connection.exec_driver_sql("PRAGMA user_version").scalar_one(), LATEST_SCHEMA_VERSION)

    def test_restore_failure_leaves_active_session_unchanged(self):
        self.populate()
        active_before = get_active_session(self.engine)
        with Session(self.engine) as db:
            pianist_count = db.query(models.Pianist).count()

        with self.assertRaises(SessionFileError):
            restore_session(self.engine, b"not a zip", self.recovery_directory)

        self.assertEqual(get_active_session(self.engine), active_before)
        with Session(self.engine) as db:
            self.assertEqual(db.query(models.Pianist).count(), pianist_count)
        self.assertFalse(self.recovery_directory.exists())

    def test_recovery_retention_keeps_only_three_latest_archives(self):
        for _ in range(RECOVERY_RETENTION + 2):
            create_recovery_snapshot(self.engine, self.recovery_directory)

        self.assertEqual(len(list(self.recovery_directory.glob("recovery-*.mpsession"))), RECOVERY_RETENTION)

    def test_malformed_and_missing_archive_entries_are_rejected(self):
        valid = self.archive_entries(export_session_archive(self.engine))
        cases = [
            b"not a zip",
            self.make_archive(valid[DATABASE_ENTRY], filenames=[(DATABASE_ENTRY, valid[DATABASE_ENTRY])]),
            self.make_archive(valid[DATABASE_ENTRY], filenames=[("manifest.json", valid["manifest.json"])]),
        ]
        for archive_bytes in cases:
            with self.subTest(archive_size=len(archive_bytes)), self.assertRaises(SessionFileError):
                restore_session(self.engine, archive_bytes, self.recovery_directory)

    def test_archive_size_limit_rejects_oversized_input(self):
        archive_bytes = export_session_archive(self.engine)

        with patch("app.services.session_files.MAX_ARCHIVE_BYTES", len(archive_bytes) - 1):
            with self.assertRaisesRegex(SessionFileError, "exceeds the supported size"):
                restore_session(self.engine, archive_bytes, self.recovery_directory)

    def test_unsupported_format_and_newer_schema_are_rejected(self):
        payload = self.archive_entries(export_session_archive(self.engine))[DATABASE_ENTRY]
        unsupported_format = self.make_archive(payload, manifest_overrides={"formatVersion": 77})
        with self.assertRaisesRegex(SessionFileError, "not supported"):
            restore_session(self.engine, unsupported_format, self.recovery_directory)

        newer_path = self.directory / "newer.sqlite3"
        with sqlite3.connect(newer_path) as connection:
            connection.executescript("CREATE TABLE marker (id INTEGER); PRAGMA user_version = 77;")
        newer_archive = self.make_archive(newer_path.read_bytes(), schema_version=77)
        with self.assertRaisesRegex(SessionFileError, "newer than supported"):
            restore_session(self.engine, newer_archive, self.recovery_directory)

    def test_corrupt_database_and_unsafe_paths_are_rejected(self):
        corrupt = self.make_archive(b"not sqlite")
        with self.assertRaisesRegex(SessionFileError, "SQLite database"):
            restore_session(self.engine, corrupt, self.recovery_directory)

        traversal = self.make_archive(
            b"not used",
            filenames=[("../manifest.json", b"{}"), (DATABASE_ENTRY, b"not used")],
        )
        with self.assertRaisesRegex(SessionFileError, "unsafe file path"):
            restore_session(self.engine, traversal, self.recovery_directory)

    def test_duplicate_unexpected_and_checksum_mismatch_archives_are_rejected(self):
        entries = self.archive_entries(export_session_archive(self.engine))
        manifest = entries["manifest.json"]
        database = entries[DATABASE_ENTRY]
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            duplicate = self.make_archive(
                database,
                filenames=[("manifest.json", manifest), (DATABASE_ENTRY, database), (DATABASE_ENTRY, database)],
            )
        unexpected = self.make_archive(
            database,
            filenames=[("manifest.json", manifest), (DATABASE_ENTRY, database), ("notes.txt", b"unexpected")],
        )
        altered_payload = self.make_archive(
            database + b"changed",
            filenames=[("manifest.json", manifest), (DATABASE_ENTRY, database + b"changed")],
        )

        for archive_bytes in (duplicate, unexpected, altered_payload):
            with self.subTest(size=len(archive_bytes)), self.assertRaises(SessionFileError):
                restore_session(self.engine, archive_bytes, self.recovery_directory)

    def test_invalid_metadata_and_database_schema_claim_are_rejected(self):
        entries = self.archive_entries(export_session_archive(self.engine))
        payload = entries[DATABASE_ENTRY]
        invalid_metadata = self.make_archive(payload, manifest_overrides={"institution": "   "})
        schema_mismatch = self.make_archive(payload, schema_version=1)

        with self.assertRaises(SessionFileError):
            restore_session(self.engine, invalid_metadata, self.recovery_directory)
        with self.assertRaisesRegex(SessionFileError, "does not match"):
            restore_session(self.engine, schema_mismatch, self.recovery_directory)

    def test_staged_migration_failure_leaves_active_session_unchanged(self):
        schema_one_path = self.directory / "unrecognized-schema-one.sqlite3"
        staged_engine = create_engine(f"sqlite:///{schema_one_path}")
        with staged_engine.begin() as connection:
            MIGRATIONS[0].upgrade(connection)
            connection.exec_driver_sql("CREATE TABLE unrecognized_data (id INTEGER PRIMARY KEY)")
            connection.exec_driver_sql("PRAGMA user_version = 1")
        staged_engine.dispose()
        archive = self.make_archive(schema_one_path.read_bytes(), schema_version=1)
        original = get_active_session(self.engine)

        with self.assertRaises(SessionFileError):
            restore_session(self.engine, archive, self.recovery_directory)

        self.assertEqual(get_active_session(self.engine), original)
        self.assertFalse(self.recovery_directory.exists())


if __name__ == "__main__":
    unittest.main()