from pathlib import Path
import sqlite3
import tempfile
import unittest

from sqlalchemy import create_engine, text
from sqlalchemy.orm import Session

from app import database, models
from app.database import (
    DatabaseMigrationError,
    DatabaseState,
    SchemaMigration,
    migrate_database,
)

LEGACY_FIXTURE = Path(__file__).parent / "fixtures" / "legacy_accompanist_v0.sql"


class DatabaseMigrationTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.database_path = Path(self.temporary_directory.name) / "session.sqlite3"
        self.engine = create_engine(f"sqlite:///{self.database_path}")

    def tearDown(self):
        self.engine.dispose()
        self.temporary_directory.cleanup()

    def test_fresh_database_initializes_current_schema(self):
        report = migrate_database(self.engine)

        self.assertEqual(report.initial_state, DatabaseState.FRESH)
        self.assertEqual(report.initial_version, 0)
        self.assertEqual(report.current_version, 2)
        self.assertEqual(len(report.applied_migrations), 2)
        with self.engine.connect() as connection:
            self.assertEqual(connection.exec_driver_sql("PRAGMA user_version").scalar_one(), 2)
            self.assertEqual(
                set(connection.execute(text(
                    "SELECT name FROM sqlite_master WHERE type='table' "
                    "AND name NOT LIKE 'sqlite_%'"
                )).scalars()),
                set(models.Base.metadata.tables),
            )
            self.assertEqual(
                connection.execute(text("SELECT name FROM organizations WHERE id = 1")).scalar_one(),
                "Default Organization",
            )
            session = connection.execute(text("SELECT * FROM scheduling_sessions")).mappings().one()
            self.assertEqual(session["term_label"], "Term not set")

    def test_current_database_reopens_without_changing_file_or_domain_data(self):
        migrate_database(self.engine)
        with Session(self.engine) as db:
            db.add(models.Pianist(name="Synthetic Current Pianist"))
            db.flush()
            db.add(models.Lesson(
                teacher="Synthetic Current Teacher",
                student="Synthetic Current Student",
                day="Tuesday",
                start_minute=600,
                end_minute=650,
            ))
            db.commit()

        self.engine.dispose()
        before = self.database_path.read_bytes()
        self.engine = create_engine(f"sqlite:///{self.database_path}")
        report = migrate_database(self.engine)
        self.engine.dispose()
        after = self.database_path.read_bytes()

        self.assertEqual(report.initial_state, DatabaseState.CURRENT)
        self.assertEqual(report.initial_version, 2)
        self.assertEqual(report.current_version, 2)
        self.assertEqual(report.applied_migrations, ())
        self.assertEqual(after, before)

    def test_legacy_fixture_migrates_forward_without_losing_existing_records(self):
        self.engine.dispose()
        with sqlite3.connect(self.database_path) as connection:
            connection.executescript(LEGACY_FIXTURE.read_text(encoding="utf-8"))
        self.engine = create_engine(f"sqlite:///{self.database_path}")

        report = migrate_database(self.engine)

        self.assertEqual(report.initial_state, DatabaseState.LEGACY)
        self.assertEqual(report.initial_version, 0)
        self.assertEqual(report.current_version, 2)
        with Session(self.engine) as db:
            organization = db.get(models.Organization, 1)
            pianist = db.get(models.Pianist, 7)
            lesson = db.get(models.Lesson, 11)
            profile = db.get(models.ImportProfile, 13)
            self.assertEqual(organization.name, "Synthetic Music Program")
            self.assertEqual(pianist.name, "Synthetic Pianist")
            self.assertEqual(pianist.email, "pianist@example.invalid")
            self.assertEqual(pianist.availability[0].status, "Available")
            self.assertEqual(lesson.student, "Synthetic Student")
            self.assertEqual(lesson.teacher, "Synthetic Instructor")
            self.assertEqual(lesson.assigned_pianist_id, 7)
            self.assertEqual(lesson.teacher_email, "")
            self.assertEqual(lesson.student_id, "")
            self.assertEqual(profile.name, "Synthetic Lesson Map")
            self.assertEqual(profile.mapping_json, '{"student":"Student Name"}')
            self.assertEqual(db.query(models.Lesson).count(), 1)

    def test_unversioned_current_shape_is_recognized_as_legacy(self):
        Base = models.Base
        baseline_tables = [
            table for table in Base.metadata.sorted_tables
            if table.name != "scheduling_sessions"
        ]
        Base.metadata.create_all(self.engine, tables=baseline_tables)
        with Session(self.engine) as db:
            db.add(models.Organization(id=1, name="Synthetic Existing Program"))
            db.add(models.Lesson(
                organization_id=1,
                teacher="Synthetic Teacher",
                teacher_email="teacher@example.invalid",
                student="Synthetic Student",
                student_id="SYN-01",
                day="Wednesday",
                start_minute=600,
                end_minute=650,
            ))
            db.commit()

        report = migrate_database(self.engine)

        self.assertEqual(report.initial_state, DatabaseState.LEGACY)
        self.assertEqual(report.current_version, 2)
        with Session(self.engine) as db:
            lesson = db.query(models.Lesson).one()
            self.assertEqual(lesson.teacher_email, "teacher@example.invalid")
            self.assertEqual(lesson.student_id, "SYN-01")
            self.assertEqual(lesson.student, "Synthetic Student")

    def test_failed_migration_rolls_back_ddl_and_keeps_version_unadvanced(self):
        def fail_after_ddl(connection):
            database._upgrade_to_v1(connection)
            connection.exec_driver_sql("CREATE TABLE partial_migration (id INTEGER PRIMARY KEY)")
            raise RuntimeError("synthetic migration failure")

        migration = SchemaMigration(1, "synthetic_failure", fail_after_ddl)
        with self.assertRaises(DatabaseMigrationError) as raised:
            migrate_database(self.engine, migrations=(migration,))

        self.assertEqual(raised.exception.code, "MIGRATION_FAILED")
        with self.engine.connect() as connection:
            self.assertEqual(connection.exec_driver_sql("PRAGMA user_version").scalar_one(), 0)
            self.assertEqual(
                connection.execute(text(
                    "SELECT name FROM sqlite_master WHERE type='table' "
                    "AND name NOT LIKE 'sqlite_%'"
                )).all(),
                [],
            )

    def test_newer_schema_version_is_rejected_without_downgrade(self):
        migrate_database(self.engine)
        with self.engine.begin() as connection:
            connection.exec_driver_sql("PRAGMA user_version = 3")

        with self.assertRaises(DatabaseMigrationError) as raised:
            migrate_database(self.engine)

        self.assertEqual(raised.exception.code, "UNSUPPORTED_SCHEMA_VERSION")
        with self.engine.connect() as connection:
            self.assertEqual(connection.exec_driver_sql("PRAGMA user_version").scalar_one(), 3)
            self.assertEqual(
                connection.execute(text("SELECT name FROM organizations WHERE id = 1")).scalar_one(),
                "Default Organization",
            )

    def test_unrecognized_database_fails_without_modification(self):
        with self.engine.begin() as connection:
            connection.exec_driver_sql("CREATE TABLE unrelated_data (id INTEGER PRIMARY KEY)")
            connection.exec_driver_sql("INSERT INTO unrelated_data VALUES (42)")
        self.engine.dispose()
        before = self.database_path.read_bytes()
        self.engine = create_engine(f"sqlite:///{self.database_path}")

        with self.assertRaises(DatabaseMigrationError) as raised:
            migrate_database(self.engine)

        self.assertEqual(raised.exception.code, "UNRECOGNIZED_DATABASE")
        self.engine.dispose()
        self.assertEqual(self.database_path.read_bytes(), before)

    def test_malformed_known_schema_fails_clearly(self):
        with self.engine.begin() as connection:
            connection.exec_driver_sql("CREATE TABLE lessons (id INTEGER PRIMARY KEY, student TEXT)")

        with self.assertRaises(DatabaseMigrationError) as raised:
            migrate_database(self.engine)

        self.assertEqual(raised.exception.code, "INVALID_DATABASE_SCHEMA")
        self.assertIn("lessons", str(raised.exception))

    def test_migration_registry_rejects_skipped_versions(self):
        skipped_migration = SchemaMigration(2, "skipped_v1", lambda connection: None)

        with self.assertRaises(DatabaseMigrationError) as raised:
            migrate_database(self.engine, migrations=(skipped_migration,))

        self.assertEqual(raised.exception.code, "INVALID_MIGRATION_REGISTRY")
        self.assertFalse(self.database_path.exists())


if __name__ == "__main__":
    unittest.main()