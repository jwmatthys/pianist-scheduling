"""Lesson Scheduling foundation: persistence, repository, migration, and module boundaries.

All data is synthetic.
"""

import tempfile
import unittest
from pathlib import Path

from sqlalchemy import create_engine, inspect
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from app import clinical_models, lesson_models, models
from app.database import LATEST_SCHEMA_VERSION, migrate_database
from app.lesson_models import (
    LessonInstructor,
    LessonModuleState,
    LessonPlacementType,
    LessonTagState,
)
from app.module_models import ModuleBase
from app.services.availability import AvailabilityWindow
from app.services.clinical_repository import ClinicalRepository
from app.services.lesson_repository import LessonRepository, LessonValidationError

LESSON_TABLES = {
    "lesson_tags",
    "lesson_instructors",
    "lesson_instructor_tags",
    "lesson_instructor_availability_windows",
    "lesson_students",
    "lesson_student_availability_windows",
    "lesson_student_tag_states",
    "lesson_placements",
    "lesson_module_states",
}


def window(day="Monday", start=9 * 60, end=12 * 60, status="Available"):
    return AvailabilityWindow(day=day, start_minute=start, end_minute=end, status=status)


class LessonTestCase(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.engine = create_engine(
            f"sqlite:///{Path(self.temporary_directory.name) / 'session.sqlite3'}",
            connect_args={"check_same_thread": False},
        )
        migrate_database(self.engine)
        self.db = Session(self.engine)
        self.repo = LessonRepository(self.db)

    def tearDown(self):
        self.db.close()
        self.engine.dispose()
        self.temporary_directory.cleanup()

    def revision(self):
        return self.repo.get_state().source_data_revision


class LessonPersistenceTests(LessonTestCase):
    def test_instructor_defaults_and_fields(self):
        instructor = self.repo.create_instructor(" Synthetic Teacher ")
        self.assertEqual(instructor.name, "Synthetic Teacher")
        self.assertEqual(instructor.lesson_length_minutes, 60)
        self.assertEqual(instructor.capacity, 1)
        self.assertEqual((instructor.email, instructor.room), ("", ""))
        self.assertIsNotNone(instructor.created_at)
        self.assertIsNotNone(instructor.updated_at)

        updated = self.repo.update_instructor(
            instructor.id, "Synthetic Teacher", "t@example.test", "Room 1", 30, 4
        )
        self.assertEqual(
            (updated.email, updated.room, updated.lesson_length_minutes, updated.capacity),
            ("t@example.test", "Room 1", 30, 4),
        )

    def test_instructor_validation(self):
        with self.assertRaises(LessonValidationError):
            self.repo.create_instructor("  ")
        with self.assertRaises(LessonValidationError):
            self.repo.create_instructor("A", lesson_length_minutes=0)
        with self.assertRaises(LessonValidationError):
            self.repo.create_instructor("A", capacity=0)

    def test_student_creation_requires_unique_case_insensitive_student_id(self):
        student = self.repo.create_student(" S-1 ", "Pat Test", "pat@example.test")
        self.assertEqual(student.student_id, "S-1")
        with self.assertRaises(LessonValidationError) as missing:
            self.repo.create_student("  ")
        self.assertEqual(missing.exception.code, "INVALID_STUDENT_ID")
        with self.assertRaises(LessonValidationError) as duplicate:
            self.repo.create_student("s-1")
        self.assertEqual(duplicate.exception.code, "DUPLICATE_STUDENT_ID")

    def test_tags_are_uppercase_and_case_insensitively_unique(self):
        tag = self.repo.create_tag("  voice ")
        self.assertEqual((tag.name, tag.normalized_name), ("VOICE", "voice"))
        with self.assertRaises(LessonValidationError):
            self.repo.create_tag("Voice")
        self.assertEqual(self.repo.rename_tag(tag.id, "opera").name, "OPERA")
        self.repo.delete_tag(tag.id)
        self.assertEqual(self.repo.list_tags(), [])

    def test_tag_states_use_absence_for_neutral(self):
        student = self.repo.create_student("S-1", "Pat")
        tag = self.repo.create_tag("violin")
        self.assertEqual(self.repo.get_tag_state(student.id, tag.id), LessonTagState.NEUTRAL)
        for state in (LessonTagState.REQUIRED, LessonTagState.PREFERRED, LessonTagState.RESTRICTED):
            self.repo.set_tag_state(student.id, tag.id, state)
            self.assertEqual(self.repo.get_tag_state(student.id, tag.id), state)
        self.repo.set_tag_state(student.id, tag.id, LessonTagState.NEUTRAL)
        self.assertEqual(self.repo.student_tag_states(), {})
        with self.assertRaises(LessonValidationError):
            self.repo.set_tag_state(student.id, tag.id, "bogus")

    def test_instructor_tags_and_deleting_tag_cleans_associations(self):
        tag = self.repo.create_tag("piano")
        instructor = self.repo.create_instructor("T", tag_ids=[tag.id])
        self.assertEqual([t.name for t in instructor.tags], ["PIANO"])
        student = self.repo.create_student("S-1")
        self.repo.set_tag_state(student.id, tag.id, LessonTagState.REQUIRED)
        self.repo.delete_tag(tag.id)
        self.db.refresh(instructor)
        self.assertEqual(list(instructor.tags), [])
        self.assertEqual(self.repo.student_tag_states(), {})

    def test_availability_persists_separately_and_merges_same_status(self):
        instructor = self.repo.create_instructor("T")
        student = self.repo.create_student("S-1")
        self.repo.set_instructor_availability(
            instructor.id,
            [window(start=540, end=600), window(start=600, end=660), window(day="Tuesday", status="Tentative")],
        )
        self.repo.set_student_availability(student.id, [window(day="Wednesday")])
        stored = self.repo.instructor_availability(instructor.id)
        self.assertEqual(
            [(w.day, w.start_minute, w.end_minute, w.status) for w in stored],
            [("Monday", 540, 660, "Available"), ("Tuesday", 540, 720, "Tentative")],
        )
        self.assertEqual([w.day for w in self.repo.student_availability(student.id)], ["Wednesday"])

    def test_conflicting_or_invalid_availability_is_rejected_without_partial_save(self):
        instructor = self.repo.create_instructor("T", availability=[window()])
        with self.assertRaises(LessonValidationError):
            self.repo.set_instructor_availability(
                instructor.id, [window(start=540, end=660), window(start=600, end=720, status="Unavailable")]
            )
        with self.assertRaises(LessonValidationError):
            self.repo.set_instructor_availability(instructor.id, [window(start=700, end=600)])
        self.assertEqual(len(self.repo.instructor_availability(instructor.id)), 1)

    def test_placement_types_and_single_placement_per_student(self):
        instructor = self.repo.create_instructor("T")
        student = self.repo.create_student("S-1")
        self.repo.set_placement(student.id, LessonPlacementType.AUTOMATIC, instructor.id, 540)
        self.repo.set_placement(student.id, LessonPlacementType.MANUAL, instructor.id, 600)
        placements = self.repo.list_placements()
        self.assertEqual(len(placements), 1)
        self.assertEqual((placements[0].placement_type, placements[0].start_minute), ("manual", 600))
        self.repo.set_placement(student.id, LessonPlacementType.MANUALLY_UNPLACED)
        placement = self.repo.get_placement(student.id)
        self.assertEqual(placement.placement_type, "manuallyUnplaced")
        self.assertIsNone(placement.lesson_instructor_id)
        self.assertIsNone(placement.start_minute)

    def test_placement_shape_validation(self):
        instructor = self.repo.create_instructor("T")
        student = self.repo.create_student("S-1")
        with self.assertRaises(LessonValidationError):
            self.repo.set_placement(student.id, LessonPlacementType.MANUAL)
        with self.assertRaises(LessonValidationError):
            self.repo.set_placement(student.id, LessonPlacementType.MANUALLY_UNPLACED, instructor.id, 540)
        with self.assertRaises(LessonValidationError):
            self.repo.set_placement(student.id, LessonPlacementType.MANUAL, instructor.id, 24 * 60)
        with self.assertRaises(LessonValidationError):
            self.repo.set_placement(student.id, LessonPlacementType.MANUAL, "missing", 540)

    def test_database_rejects_malformed_placement_rows(self):
        student = self.repo.create_student("S-1")
        self.db.add(lesson_models.LessonPlacement(
            lesson_student_id=student.id, placement_type="manuallyUnplaced", start_minute=540
        ))
        with self.assertRaises(IntegrityError):
            self.db.flush()
        self.db.rollback()

    def test_deleting_student_or_instructor_removes_dependent_placements(self):
        instructor = self.repo.create_instructor("T")
        first, second = self.repo.create_student("S-1"), self.repo.create_student("S-2")
        self.repo.set_placement(first.id, LessonPlacementType.MANUAL, instructor.id, 540)
        self.repo.set_placement(second.id, LessonPlacementType.MANUAL, instructor.id, 600)
        self.repo.delete_student(first.id)
        self.assertEqual(len(self.repo.list_placements()), 1)
        self.repo.delete_instructor(instructor.id)
        self.assertEqual(self.repo.list_placements(), [])

    def test_module_state_persists_revisions_and_staleness(self):
        state = self.repo.get_state()
        self.assertEqual(
            (state.schedule_revision, state.source_data_revision, state.is_schedule_stale), (0, 0, False)
        )
        instructor = self.repo.create_instructor("T")
        student = self.repo.create_student("S-1")
        self.assertEqual(self.revision(), 2)
        self.assertFalse(self.repo.get_state().is_schedule_stale)
        self.repo.set_placement(student.id, LessonPlacementType.AUTOMATIC, instructor.id, 540)
        self.assertEqual(self.repo.get_state().schedule_revision, 1)
        self.repo.update_instructor(instructor.id, "T2")
        self.db.commit()
        with Session(self.engine) as other:
            stored = other.query(LessonModuleState).one()
            self.assertTrue(stored.is_schedule_stale)
            self.assertEqual(stored.source_data_revision, 3)
        self.assertEqual(self.repo.clear_all_placements(), 1)
        state = self.repo.get_state()
        self.assertEqual((state.schedule_revision, state.is_schedule_stale), (0, False))

    def test_every_source_mutation_bumps_revision(self):
        seen = [self.revision()]
        tag = self.repo.create_tag("t"); seen.append(self.revision())
        self.repo.rename_tag(tag.id, "t2"); seen.append(self.revision())
        instructor = self.repo.create_instructor("T"); seen.append(self.revision())
        self.repo.update_instructor(instructor.id, "T2"); seen.append(self.revision())
        self.repo.set_instructor_tags(instructor.id, [tag.id]); seen.append(self.revision())
        self.repo.set_instructor_availability(instructor.id, [window()]); seen.append(self.revision())
        student = self.repo.create_student("S-1"); seen.append(self.revision())
        self.repo.update_student(student.id, "S-1", "Pat"); seen.append(self.revision())
        self.repo.set_student_availability(student.id, [window()]); seen.append(self.revision())
        self.repo.set_tag_state(student.id, tag.id, LessonTagState.PREFERRED); seen.append(self.revision())
        self.repo.delete_student(student.id); seen.append(self.revision())
        self.repo.delete_instructor(instructor.id); seen.append(self.revision())
        self.repo.delete_tag(tag.id); seen.append(self.revision())
        self.assertEqual(seen, list(range(seen[0], seen[0] + len(seen))))


class LessonMigrationTests(LessonTestCase):
    def test_migration_creates_lesson_tables_and_module_state_for_existing_session(self):
        self.assertGreaterEqual(LATEST_SCHEMA_VERSION, 12)
        self.assertTrue(LESSON_TABLES <= set(inspect(self.engine).get_table_names()))
        session_uuid = self.db.query(models.SchedulingSession.session_uuid).scalar()
        self.assertEqual(self.db.query(LessonModuleState).one().session_uuid, session_uuid)

    def test_v11_database_upgrades_preserving_existing_data(self):
        self.db.add(models.Pianist(id=7, name="Synthetic Pianist"))
        self.db.commit()
        self.db.close()
        with self.engine.begin() as connection:
            for table in sorted(LESSON_TABLES - {"lesson_module_states"}, reverse=True):
                connection.exec_driver_sql(f"DROP TABLE {table}")
            connection.exec_driver_sql("DROP TABLE lesson_module_states")
            connection.exec_driver_sql("PRAGMA user_version = 11")

        report = migrate_database(self.engine)

        self.assertEqual(report.initial_version, 11)
        self.assertEqual(report.current_version, LATEST_SCHEMA_VERSION)
        self.assertEqual(report.applied_migrations, ("add_lesson_scheduling_module_tables",))
        with Session(self.engine) as db:
            self.assertEqual(db.get(models.Pianist, 7).name, "Synthetic Pianist")
            self.assertEqual(db.query(LessonModuleState).count(), 1)
            self.assertEqual(LessonRepository(db).list_instructors(), [])

    def test_current_database_is_not_migrated_again(self):
        self.db.close()
        self.assertEqual(migrate_database(self.engine).applied_migrations, ())

    def test_module_state_is_keyed_by_session_uuid_for_session_rekeying(self):
        self.assertIn("session_uuid", lesson_models.LessonModuleState.__table__.c)


class LessonBoundaryTests(LessonTestCase):
    def test_lesson_tables_are_registered_with_the_module_metadata(self):
        self.assertTrue(LESSON_TABLES <= set(ModuleBase.metadata.tables))

    def test_lesson_schema_has_no_foreign_keys_outside_lesson_tables(self):
        for name in LESSON_TABLES:
            for foreign_key in ModuleBase.metadata.tables[name].foreign_keys:
                self.assertTrue(foreign_key.column.table.name.startswith("lesson_"), name)

    def test_no_other_module_references_lesson_tables(self):
        for name, table in ModuleBase.metadata.tables.items():
            if name in LESSON_TABLES:
                continue
            for foreign_key in table.foreign_keys:
                self.assertFalse(foreign_key.column.table.name.startswith("lesson_"), name)

    def test_lesson_and_clinical_tags_are_independent_registries(self):
        self.assertNotEqual(lesson_models.LessonTag.__tablename__, clinical_models.ClinicalTag.__tablename__)
        clinical_repo = ClinicalRepository(self.db)
        clinical_repo.create_tag("voice")
        self.assertEqual(self.repo.list_tags(), [])
        lesson_tag = self.repo.create_tag("voice")
        self.assertEqual(len(clinical_repo.list_tags()), 1)
        self.repo.delete_tag(lesson_tag.id)
        self.assertEqual(len(clinical_repo.list_tags()), 1)

    def test_lesson_and_clinical_students_with_matching_identity_are_unrelated(self):
        clinical_repo = ClinicalRepository(self.db)
        clinical_student = clinical_repo.create_student("S-1", "Pat Test", "pat@example.test")
        lesson_student = self.repo.create_student("S-1", "Pat Test", "pat@example.test")
        self.assertNotEqual(clinical_student.id, lesson_student.id)
        self.repo.delete_student(lesson_student.id)
        self.assertEqual(len(clinical_repo.list_students()), 1)

    def test_lesson_changes_do_not_touch_clinical_state(self):
        clinical_state = ClinicalRepository(self.db).get_state()
        before = clinical_state.source_data_revision
        self.repo.create_instructor("T")
        self.repo.create_student("S-1")
        self.assertEqual(clinical_state.source_data_revision, before)
        self.assertEqual(self.db.query(LessonInstructor).count(), 1)


if __name__ == "__main__":
    unittest.main()
