"""Archive recency warnings: hidden UTC save timestamp and read-only comparison (synthetic data only)."""

import json
import re
import tempfile
import unittest
import zipfile
from datetime import date, datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import models, schemas
from app.database import MIGRATIONS, migrate_database
from app.services.session_files import (
    SessionFileError,
    _set_last_activity,
    create_new_session,
    create_recovery_snapshot,
    export_session_archive,
    inspect_session_archive,
    restore_session,
    update_session_metadata,
)

ISO_UTC = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$")
T0 = datetime(2026, 11, 2, 21, 14, 33, tzinfo=timezone.utc)


def metadata(institution="Synthetic University", term="Fall 2026"):
    return schemas.SessionMetadataIn(
        institution_name=institution,
        program_name="Music",
        term_label=term,
        year=2026,
        start_date=date(2026, 8, 20),
        end_date=date(2026, 12, 20),
    )


class RecencyTestCase(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.directory = Path(self.temporary_directory.name)
        self.recovery_directory = self.directory / "recovery"
        self.engines = []
        self.active = self.new_engine("active")

    def tearDown(self):
        for engine in self.engines:
            engine.dispose()
        self.temporary_directory.cleanup()

    def new_engine(self, name):
        engine = create_engine(
            f"sqlite:///{self.directory / (name + '.sqlite3')}", connect_args={"check_same_thread": False}
        )
        migrate_database(engine)
        self.engines.append(engine)
        return engine

    def archive_with_timestamp(self, timestamp, source=None):
        """Export a real archive, then rewrite only its manifest timestamp (None removes the field)."""
        archive_bytes = export_session_archive(source or self.new_engine("source"))
        with zipfile.ZipFile(BytesIO(archive_bytes)) as archive:
            entries = {item.filename: archive.read(item.filename) for item in archive.infolist()}
        manifest = json.loads(entries["manifest.json"])
        if timestamp is None:
            manifest.pop("exportedAt", None)
        else:
            manifest["exportedAt"] = timestamp.strftime("%Y-%m-%dT%H:%M:%SZ")
        entries["manifest.json"] = json.dumps(manifest).encode()
        output = BytesIO()
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, contents in entries.items():
                archive.writestr(name, contents)
        return output.getvalue()

    def last_activity(self, engine=None):
        with Session(engine or self.active) as db:
            return db.query(models.SchedulingSession).one().last_activity_utc

    def edit(self, engine=None, name="Synthetic Pianist"):
        with Session(engine or self.active) as db:
            db.add(models.Pianist(name=name, email="pianist@example.invalid"))
            db.commit()

    def set_current(self, moment):
        _set_last_activity(self.active, moment)


class ExportTimestampTests(RecencyTestCase):
    def test_export_writes_canonical_utc_timestamp_to_manifest(self):
        before = datetime.now(timezone.utc).replace(microsecond=0)
        archive_bytes = export_session_archive(self.active)
        after = datetime.now(timezone.utc)

        with zipfile.ZipFile(BytesIO(archive_bytes)) as archive:
            manifest = json.loads(archive.read("manifest.json"))
        self.assertRegex(manifest["exportedAt"], ISO_UTC)
        stamped = datetime.strptime(manifest["exportedAt"], "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
        self.assertTrue(before <= stamped <= after)
        self.assertEqual(set(manifest) & {"exportUtcTimestamp", "savedUtcTimestamp"}, set())

    def test_export_and_recovery_snapshot_are_not_session_activity(self):
        self.set_current(T0)
        export_session_archive(self.active)
        create_recovery_snapshot(self.active, self.recovery_directory)
        self.assertEqual(self.last_activity(), T0.replace(tzinfo=None))


class ActivityTrackingTests(RecencyTestCase):
    def assert_advances(self, action):
        self.set_current(datetime(2020, 1, 1, tzinfo=timezone.utc))
        action()
        self.assertGreater(self.last_activity(), datetime(2020, 1, 1))
        self.assertLessEqual(self.last_activity(), datetime.now(timezone.utc).replace(tzinfo=None))
        self.assertEqual(self.last_activity().microsecond, 0)

    def test_creating_records_advances_activity(self):
        self.assert_advances(self.edit)

    def test_updating_records_advances_activity(self):
        self.edit()
        def update():
            with Session(self.active) as db:
                db.query(models.Pianist).one().name = "Renamed Synthetic Pianist"
                db.commit()
        self.assert_advances(update)

    def test_deleting_records_advances_activity(self):
        self.edit()
        def delete():
            with Session(self.active) as db:
                db.delete(db.query(models.Pianist).one())
                db.commit()
        self.assert_advances(delete)

    def test_session_metadata_edit_advances_activity(self):
        self.assert_advances(lambda: update_session_metadata(self.active, metadata("Edited University")))

    def test_module_data_changes_advance_activity(self):
        from app.services.clinical_repository import ClinicalRepository

        def change():
            with Session(self.active) as db:
                ClinicalRepository(db).create_tag("Synthetic Tag")
                db.commit()
        self.assert_advances(change)

    def test_read_only_access_does_not_advance_activity(self):
        self.set_current(T0)
        with Session(self.active) as db:
            db.query(models.Pianist).all()
            db.commit()
        self.assertEqual(self.last_activity(), T0.replace(tzinfo=None))

    def test_new_session_starts_with_creation_time_activity(self):
        create_new_session(self.active, metadata("New College"), self.recovery_directory)
        self.assertIsNotNone(self.last_activity())

    def test_activity_is_unknown_until_first_change_on_migrated_database(self):
        self.assertIsNone(self.last_activity())


class InspectTests(RecencyTestCase):
    def inspect(self, archive_time, current_time):
        archive = self.archive_with_timestamp(archive_time)
        if current_time is not None:
            self.set_current(current_time)
        return inspect_session_archive(self.active, archive)

    def test_newer_archive_does_not_warn(self):
        result = self.inspect(T0 + timedelta(hours=1), T0)
        self.assertFalse(result.is_older)

    def test_older_archive_warns_with_both_timestamps(self):
        result = self.inspect(T0 - timedelta(hours=1), T0)
        self.assertTrue(result.is_older)
        self.assertEqual(result.archive_exported_utc, T0 - timedelta(hours=1))
        self.assertEqual(result.current_activity_utc, T0)

    def test_equal_timestamps_do_not_warn(self):
        self.assertFalse(self.inspect(T0, T0).is_older)

    def test_missing_archive_timestamp_skips_warning(self):
        result = self.inspect(None, T0)
        self.assertFalse(result.is_older)
        self.assertIsNone(result.archive_exported_utc)

    def test_missing_current_timestamp_skips_warning(self):
        result = self.inspect(T0, None)
        self.assertFalse(result.is_older)
        self.assertIsNone(result.current_activity_utc)

    def test_comparison_ignores_institution_program_term_and_session_identity(self):
        other = self.new_engine("other")
        create_new_session(other, metadata("Other College", "Spring 2031"), self.recovery_directory)
        archive = self.archive_with_timestamp(T0 - timedelta(days=30), source=other)
        self.set_current(T0)
        self.assertTrue(inspect_session_archive(self.active, archive).is_older)

    def test_comparison_uses_instants_not_display_strings(self):
        archive = self.archive_with_timestamp(T0)
        with zipfile.ZipFile(BytesIO(archive)) as source:
            entries = {item.filename: source.read(item.filename) for item in source.infolist()}
        manifest = json.loads(entries["manifest.json"])
        manifest["exportedAt"] = "2026-11-02T15:14:33-06:00"  # same instant as T0 in UTC
        entries["manifest.json"] = json.dumps(manifest).encode()
        output = BytesIO()
        with zipfile.ZipFile(output, "w") as target:
            for name, contents in entries.items():
                target.writestr(name, contents)
        self.set_current(T0)
        self.assertFalse(inspect_session_archive(self.active, output.getvalue()).is_older)

    def test_inspect_changes_nothing(self):
        self.set_current(T0)
        with Session(self.active) as db:
            db.add(models.Pianist(name="Synthetic Pianist", email="pianist@example.invalid"))
            db.commit()
        archive = self.archive_with_timestamp(T0 - timedelta(days=1))
        before = (self.last_activity(), (self.directory / "active.sqlite3").exists())
        inspect_session_archive(self.active, archive)
        self.assertEqual((self.last_activity(), True), before)
        with Session(self.active) as db:
            self.assertEqual(db.query(models.Pianist).count(), 1)
        self.assertFalse(self.recovery_directory.exists())

    def test_inspect_still_rejects_invalid_archives(self):
        archive = self.archive_with_timestamp(T0)
        with zipfile.ZipFile(BytesIO(archive)) as source:
            entries = {item.filename: source.read(item.filename) for item in source.infolist()}
        entries["data/session.sqlite3"] = entries["data/session.sqlite3"] + b"tampered"
        output = BytesIO()
        with zipfile.ZipFile(output, "w") as target:
            for name, contents in entries.items():
                target.writestr(name, contents)
        with self.assertRaises(SessionFileError) as raised:
            inspect_session_archive(self.active, output.getvalue())
        self.assertEqual(raised.exception.code, "PAYLOAD_CHECKSUM_MISMATCH")
        with self.assertRaises(SessionFileError):
            inspect_session_archive(self.active, b"not an archive")


class RestoreAfterWarningTests(RecencyTestCase):
    def source_with_activity(self, moment, name="source"):
        source = self.new_engine(name)
        _set_last_activity(source, moment)
        return source

    def test_restore_preserves_the_archives_own_activity_timestamp(self):
        activity = T0 - timedelta(days=3)
        archive = self.archive_with_timestamp(T0 - timedelta(days=2), source=self.source_with_activity(activity))
        self.set_current(T0)
        self.assertTrue(inspect_session_archive(self.active, archive).is_older)

        restore_session(self.active, archive, self.recovery_directory)

        self.assertEqual(self.last_activity(), activity.replace(tzinfo=None))
        self.assertTrue(list(self.recovery_directory.glob("*.mpsession")))

    def test_opening_an_archive_is_not_activity_so_the_same_archive_opens_cleanly_again(self):
        activity = T0 - timedelta(days=3)
        archive = self.archive_with_timestamp(T0 - timedelta(days=2), source=self.source_with_activity(activity))
        self.set_current(T0)
        restore_session(self.active, archive, self.recovery_directory)
        self.assertFalse(inspect_session_archive(self.active, archive).is_older)

    def test_legacy_archive_without_activity_restores_with_unknown_activity(self):
        archive = self.archive_with_timestamp(None)
        self.set_current(T0)
        restore_session(self.active, archive, self.recovery_directory)
        self.assertIsNone(self.last_activity())

    def test_unsaved_work_after_last_export_makes_that_export_older(self):
        exported = export_session_archive(self.active)
        self.edit()
        self.set_current(datetime.now(timezone.utc) + timedelta(seconds=5))
        self.assertTrue(inspect_session_archive(self.active, exported).is_older)

    def test_export_of_unchanged_session_is_not_older_than_that_session(self):
        self.edit()
        exported = export_session_archive(self.active)
        self.assertFalse(inspect_session_archive(self.active, exported).is_older)

    def test_restore_still_enforces_validation_after_a_warning(self):
        self.set_current(T0)
        archive = self.archive_with_timestamp(T0 - timedelta(days=1))
        with zipfile.ZipFile(BytesIO(archive)) as source:
            entries = {item.filename: source.read(item.filename) for item in source.infolist()}
        manifest = json.loads(entries["manifest.json"])
        manifest["databaseSchemaVersion"] = manifest["databaseSchemaVersion"] - 1
        entries["manifest.json"] = json.dumps(manifest).encode()
        output = BytesIO()
        with zipfile.ZipFile(output, "w") as target:
            for name, contents in entries.items():
                target.writestr(name, contents)
        with self.assertRaises(SessionFileError) as raised:
            restore_session(self.active, output.getvalue(), self.recovery_directory)
        self.assertEqual(raised.exception.code, "SCHEMA_VERSION_MISMATCH")
        self.assertEqual(self.last_activity(), T0.replace(tzinfo=None))


class LastActivityMigrationTests(unittest.TestCase):
    def test_v18_replaces_last_saved_with_unknown_last_activity_and_preserves_session(self):
        with tempfile.TemporaryDirectory() as directory:
            engine = create_engine(f"sqlite:///{Path(directory) / 'old.sqlite3'}")
            migrate_database(engine)
            with engine.begin() as connection:
                connection.exec_driver_sql("ALTER TABLE scheduling_sessions DROP COLUMN last_activity_utc")
                connection.exec_driver_sql("ALTER TABLE scheduling_sessions ADD COLUMN last_saved_utc DATETIME")
                connection.exec_driver_sql("PRAGMA user_version = 17")
            with Session(engine) as db:
                uuid_before = db.query(models.SchedulingSession.session_uuid).scalar()

            report = migrate_database(engine)

            self.assertEqual(report.initial_version, 17)
            self.assertEqual(report.current_version, MIGRATIONS[-1].version)
            with engine.connect() as connection:
                columns = {row[1] for row in connection.exec_driver_sql("PRAGMA table_info(scheduling_sessions)")}
            self.assertIn("last_activity_utc", columns)
            self.assertNotIn("last_saved_utc", columns)
            with Session(engine) as db:
                row = db.query(models.SchedulingSession).one()
                self.assertEqual(row.session_uuid, uuid_before)
                self.assertIsNone(row.last_activity_utc)
            engine.dispose()


if __name__ == "__main__":
    unittest.main()
