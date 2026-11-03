"""Tutor Scheduling foundation: persistence, repository, migration, and module boundaries.

All data is synthetic.
"""

import tempfile
import unittest
from pathlib import Path

from sqlalchemy import create_engine, inspect
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from app import clinical_models, tutor_models, models
from app.database import LATEST_SCHEMA_VERSION, migrate_database
from app.tutor_models import (
    TutorTutor,
    TutorModuleState,
    TutorPlacementType,
    TutorTagState,
)
from app.module_models import ModuleBase
from app.services.availability import AvailabilityWindow
from app.services.clinical_repository import ClinicalRepository
from app.services.tutor_repository import StudentInput, TutorRepository, TutorValidationError

TUTOR_TABLES = {
    "tutor_tags",
    "tutor_tutors",
    "tutor_tutor_tags",
    "tutor_tutor_availability_windows",
    "tutor_students",
    "tutor_student_availability_windows",
    "tutor_student_tag_states",
    "tutor_placements",
    "tutor_module_states",
    "tutor_schedule_placements",
    "tutor_schedule_results",
    "tutor_unplaced_students",
}


def window(day="Monday", start=9 * 60, end=12 * 60, status="Available"):
    return AvailabilityWindow(day=day, start_minute=start, end_minute=end, status=status)


class TutorTestCase(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.engine = create_engine(
            f"sqlite:///{Path(self.temporary_directory.name) / 'session.sqlite3'}",
            connect_args={"check_same_thread": False},
        )
        migrate_database(self.engine)
        self.db = Session(self.engine)
        self.repo = TutorRepository(self.db)

    def tearDown(self):
        self.db.close()
        self.engine.dispose()
        self.temporary_directory.cleanup()

    def revision(self):
        return self.repo.get_state().source_data_revision


class TutorPersistenceTests(TutorTestCase):
    def test_tutor_defaults_and_fields(self):
        tutor = self.repo.create_tutor(" Synthetic Teacher ")
        self.assertEqual(tutor.name, "Synthetic Teacher")
        self.assertEqual(tutor.lesson_length_minutes, 60)
        self.assertEqual(tutor.capacity, 1)
        self.assertEqual((tutor.email, tutor.room), ("", ""))
        self.assertIsNotNone(tutor.created_at)
        self.assertIsNotNone(tutor.updated_at)

        updated = self.repo.update_tutor(
            tutor.id, "Synthetic Teacher", "t@example.test", "Room 1", 30, 4
        )
        self.assertEqual(
            (updated.email, updated.room, updated.lesson_length_minutes, updated.capacity),
            ("t@example.test", "Room 1", 30, 4),
        )

    def test_tutor_validation(self):
        with self.assertRaises(TutorValidationError):
            self.repo.create_tutor("  ")
        with self.assertRaises(TutorValidationError):
            self.repo.create_tutor("A", lesson_length_minutes=0)
        with self.assertRaises(TutorValidationError):
            self.repo.create_tutor("A", capacity=0)

    def test_student_id_may_repeat_only_for_the_same_name(self):
        student = self.repo.create_student(" S-1 ", "Pat Test", "pat@example.test")
        self.assertEqual(student.student_id, "S-1")
        with self.assertRaises(TutorValidationError) as missing:
            self.repo.create_student("  ")
        self.assertEqual(missing.exception.code, "INVALID_STUDENT_ID")
        with self.assertRaises(TutorValidationError) as conflict:
            self.repo.create_student("s-1", "Another Student")
        self.assertEqual(conflict.exception.code, "STUDENT_ID_NAME_CONFLICT")
        second = self.repo.create_student("S-1", "  pat   TEST ")
        self.assertEqual(len([item for item in self.repo.list_students() if item.student_id.casefold() == "s-1"]), 2)
        # A rename of an unchanged Student ID applies to every record of that person.
        self.repo.update_student(second.id, "S-1", "Pat Renamed")
        self.assertEqual({item.name for item in self.repo.list_students()}, {"Pat Renamed"})

    def test_repeated_student_id_may_differ_in_availability_and_tags(self):
        tag = self.repo.create_tag("Voice")
        piano = self.repo.create_student("S-1", "Pat", availability=[window("Tuesday")])
        voice = self.repo.create_student(
            "s-1", "Pat", availability=[window("Thursday")], tag_states={tag.id: TutorTagState.REQUIRED}
        )
        self.assertEqual([item.day for item in self.repo.student_availability(piano.id)], ["Tuesday"])
        self.assertEqual([item.day for item in self.repo.student_availability(voice.id)], ["Thursday"])
        self.assertEqual(self.repo.student_tag_states().get(piano.id, {}), {})
        self.assertEqual(self.repo.student_tag_states()[voice.id], {tag.id: TutorTagState.REQUIRED})

    def test_replace_students_accepts_repeated_ids_but_blocks_name_conflicts(self):
        def entry(student_id, name, day, tags=()):
            return (StudentInput(student_id, name, ""), [window(day)], list(tags))

        result = self.repo.replace_students([
            entry("S-1", "Pat", "Tuesday"),
            entry("s-1", " pat ", "Thursday", [("Voice", "required")]),
        ])
        self.assertEqual(result.students_created, 2)
        with self.assertRaises(TutorValidationError) as conflict:
            self.repo.replace_students([entry("S-1", "Pat", "Monday"), entry("S-1", "Other", "Monday")])
        self.assertEqual(conflict.exception.code, "STUDENT_ID_NAME_CONFLICT")
        self.assertEqual(len(self.repo.list_students()), 2)

    def test_manual_placements_for_one_student_id_must_not_overlap(self):
        tutor_a = self.repo.create_tutor("A", capacity=4, availability=[window()])
        tutor_b = self.repo.create_tutor("B", capacity=4, availability=[window()])
        first = self.repo.create_student("S-1", "Pat", availability=[window()])
        second = self.repo.create_student("S-1", "Pat", availability=[window()])
        other = self.repo.create_student("S-2", "Sam", availability=[window()])
        self.repo.set_manual_placement(first.id, tutor_a.id, "Monday", 540)
        self.repo.set_manual_placement(second.id, tutor_b.id, "Monday", 570)
        self.repo.set_manual_placement(other.id, tutor_b.id, "Monday", 660)

        entries = {item.student_id: item for item in self.repo._schedule_review()[3]}

        for record in (first, second):
            codes = {warning.code.value for warning in entries[record.id].warnings}
            self.assertIn("STUDENT_OVERLAP", codes)
            self.assertFalse(entries[record.id].verified)
        self.assertNotIn("STUDENT_OVERLAP", {warning.code.value for warning in entries[other.id].warnings})
        self.repo.set_manual_placement(second.id, tutor_b.id, "Monday", 600)
        entries = {item.student_id: item for item in self.repo._schedule_review()[3]}
        self.assertNotIn("STUDENT_OVERLAP", {warning.code.value for warning in entries[first.id].warnings})

    def test_student_creation_and_update_require_a_name(self):
        with self.assertRaises(TutorValidationError) as missing:
            self.repo.create_student("S-1", "  ")
        self.assertEqual(missing.exception.code, "INVALID_STUDENT_NAME")

        student = self.repo.create_student("S-1", "Pat")
        with self.assertRaises(TutorValidationError) as blank:
            self.repo.update_student(student.id, "S-1", "  ")
        self.assertEqual(blank.exception.code, "INVALID_STUDENT_NAME")

    def test_tags_are_uppercase_and_case_insensitively_unique(self):
        tag = self.repo.create_tag("  voice ")
        self.assertEqual((tag.name, tag.normalized_name), ("VOICE", "voice"))
        with self.assertRaises(TutorValidationError):
            self.repo.create_tag("Voice")
        self.assertEqual(self.repo.rename_tag(tag.id, "opera").name, "OPERA")
        self.repo.delete_tag(tag.id)
        self.assertEqual(self.repo.list_tags(), [])

    def test_tag_states_use_absence_for_neutral(self):
        student = self.repo.create_student("S-1", "Pat")
        tag = self.repo.create_tag("violin")
        self.assertEqual(self.repo.get_tag_state(student.id, tag.id), TutorTagState.NEUTRAL)
        for state in (TutorTagState.REQUIRED, TutorTagState.PREFERRED, TutorTagState.RESTRICTED):
            self.repo.set_tag_state(student.id, tag.id, state)
            self.assertEqual(self.repo.get_tag_state(student.id, tag.id), state)
        self.repo.set_tag_state(student.id, tag.id, TutorTagState.NEUTRAL)
        self.assertEqual(self.repo.student_tag_states(), {})
        with self.assertRaises(TutorValidationError):
            self.repo.set_tag_state(student.id, tag.id, "bogus")

    def test_tutor_tags_and_deleting_tag_cleans_associations(self):
        tag = self.repo.create_tag("piano")
        tutor = self.repo.create_tutor("T", tag_ids=[tag.id])
        self.assertEqual([t.name for t in tutor.tags], ["PIANO"])
        student = self.repo.create_student("S-1", "Pat Test")
        self.repo.set_tag_state(student.id, tag.id, TutorTagState.REQUIRED)
        self.repo.delete_tag(tag.id)
        self.db.refresh(tutor)
        self.assertEqual(list(tutor.tags), [])
        self.assertEqual(self.repo.student_tag_states(), {})

    def test_availability_persists_separately_and_merges_same_status(self):
        tutor = self.repo.create_tutor("T")
        student = self.repo.create_student("S-1", "Pat Test")
        self.repo.set_tutor_availability(
            tutor.id,
            [window(start=540, end=600), window(start=600, end=660), window(day="Tuesday", status="Tentative")],
        )
        self.repo.set_student_availability(student.id, [window(day="Wednesday")])
        stored = self.repo.tutor_availability(tutor.id)
        self.assertEqual(
            [(w.day, w.start_minute, w.end_minute, w.status) for w in stored],
            [("Monday", 540, 660, "Available"), ("Tuesday", 540, 720, "Tentative")],
        )
        self.assertEqual([w.day for w in self.repo.student_availability(student.id)], ["Wednesday"])

    def test_conflicting_or_invalid_availability_is_rejected_without_partial_save(self):
        tutor = self.repo.create_tutor("T", availability=[window()])
        with self.assertRaises(TutorValidationError):
            self.repo.set_tutor_availability(
                tutor.id, [window(start=540, end=660), window(start=600, end=720, status="Unavailable")]
            )
        with self.assertRaises(TutorValidationError):
            self.repo.set_tutor_availability(tutor.id, [window(start=700, end=600)])
        self.assertEqual(len(self.repo.tutor_availability(tutor.id)), 1)

    def test_placement_types_and_single_placement_per_student(self):
        tutor = self.repo.create_tutor("T")
        student = self.repo.create_student("S-1", "Pat Test")
        self.repo.set_placement(student.id, TutorPlacementType.AUTOMATIC, tutor.id, 540)
        self.repo.set_placement(student.id, TutorPlacementType.MANUAL, tutor.id, 600)
        placements = self.repo.list_placements()
        self.assertEqual(len(placements), 1)
        self.assertEqual((placements[0].placement_type, placements[0].start_minute), ("manual", 600))
        self.repo.set_placement(student.id, TutorPlacementType.MANUALLY_UNPLACED)
        placement = self.repo.get_placement(student.id)
        self.assertEqual(placement.placement_type, "manuallyUnplaced")
        self.assertIsNone(placement.tutor_tutor_id)
        self.assertIsNone(placement.start_minute)

    def test_placement_shape_validation(self):
        tutor = self.repo.create_tutor("T")
        student = self.repo.create_student("S-1", "Pat Test")
        with self.assertRaises(TutorValidationError):
            self.repo.set_placement(student.id, TutorPlacementType.MANUAL)
        with self.assertRaises(TutorValidationError):
            self.repo.set_placement(student.id, TutorPlacementType.MANUALLY_UNPLACED, tutor.id, 540)
        with self.assertRaises(TutorValidationError):
            self.repo.set_placement(student.id, TutorPlacementType.MANUAL, tutor.id, 24 * 60)
        with self.assertRaises(TutorValidationError):
            self.repo.set_placement(student.id, TutorPlacementType.MANUAL, "missing", 540)

    def test_database_rejects_malformed_placement_rows(self):
        student = self.repo.create_student("S-1", "Pat Test")
        self.db.add(tutor_models.TutorPlacement(
            tutor_student_id=student.id, placement_type="manuallyUnplaced", start_minute=540
        ))
        with self.assertRaises(IntegrityError):
            self.db.flush()
        self.db.rollback()

    def test_deleting_student_or_tutor_preserves_dependent_placements(self):
        tutor = self.repo.create_tutor("T")
        first, second = self.repo.create_student("S-1", "Pat"), self.repo.create_student("S-2", "Sam")
        self.repo.set_placement(first.id, TutorPlacementType.MANUAL, tutor.id, 540)
        self.repo.set_placement(second.id, TutorPlacementType.MANUAL, tutor.id, 600)
        self.repo.delete_student(first.id)
        self.assertEqual(len(self.repo.list_placements()), 2)
        self.repo.delete_tutor(tutor.id)
        self.assertEqual(len(self.repo.list_placements()), 2)

    def test_module_state_persists_revisions_and_staleness(self):
        state = self.repo.get_state()
        self.assertEqual(
            (state.schedule_revision, state.source_data_revision, state.is_schedule_stale), (0, 0, False)
        )
        tutor = self.repo.create_tutor("T")
        student = self.repo.create_student("S-1", "Pat Test")
        self.assertEqual(self.revision(), 2)
        self.assertFalse(self.repo.get_state().is_schedule_stale)
        self.repo.set_placement(student.id, TutorPlacementType.AUTOMATIC, tutor.id, 540)
        self.assertEqual(self.repo.get_state().schedule_revision, 1)
        self.repo.update_tutor(tutor.id, "T2")
        self.db.commit()
        with Session(self.engine) as other:
            stored = other.query(TutorModuleState).one()
            self.assertTrue(stored.is_schedule_stale)
            self.assertEqual(stored.source_data_revision, 3)
        self.assertEqual(self.repo.clear_all_placements(), 1)
        state = self.repo.get_state()
        self.assertEqual((state.schedule_revision, state.is_schedule_stale), (0, False))

    def test_every_source_mutation_bumps_revision(self):
        seen = [self.revision()]
        tag = self.repo.create_tag("t"); seen.append(self.revision())
        self.repo.rename_tag(tag.id, "t2"); seen.append(self.revision())
        tutor = self.repo.create_tutor("T"); seen.append(self.revision())
        self.repo.update_tutor(tutor.id, "T2"); seen.append(self.revision())
        self.repo.set_tutor_tags(tutor.id, [tag.id]); seen.append(self.revision())
        self.repo.set_tutor_availability(tutor.id, [window()]); seen.append(self.revision())
        student = self.repo.create_student("S-1", "Pat"); seen.append(self.revision())
        self.repo.update_student(student.id, "S-1", "Pat"); seen.append(self.revision())
        self.repo.set_student_availability(student.id, [window()]); seen.append(self.revision())
        self.repo.set_tag_state(student.id, tag.id, TutorTagState.PREFERRED); seen.append(self.revision())
        self.repo.delete_student(student.id); seen.append(self.revision())
        self.repo.delete_tutor(tutor.id); seen.append(self.revision())
        self.repo.delete_tag(tag.id); seen.append(self.revision())
        self.assertEqual(seen, list(range(seen[0], seen[0] + len(seen))))


class TutorMigrationTests(TutorTestCase):
    def test_migration_creates_tutor_tables_and_module_state_for_existing_session(self):
        self.assertGreaterEqual(LATEST_SCHEMA_VERSION, 15)
        self.assertTrue(TUTOR_TABLES <= set(inspect(self.engine).get_table_names()))
        session_uuid = self.db.query(models.SchedulingSession.session_uuid).scalar()
        self.assertEqual(self.db.query(TutorModuleState).one().session_uuid, session_uuid)

    def test_v14_database_upgrades_preserving_existing_data(self):
        self.db.add(models.Pianist(id=7, name="Synthetic Pianist"))
        self.db.commit()
        self.db.close()
        with self.engine.begin() as connection:
            for table in sorted(TUTOR_TABLES - {"tutor_module_states"}, reverse=True):
                connection.exec_driver_sql(f"DROP TABLE {table}")
            connection.exec_driver_sql("DROP TABLE tutor_module_states")
            connection.exec_driver_sql("PRAGMA user_version = 14")

        report = migrate_database(self.engine)

        self.assertEqual(report.initial_version, 14)
        self.assertEqual(report.current_version, LATEST_SCHEMA_VERSION)
        self.assertEqual(report.applied_migrations, (
            "add_tutor_scheduling_module_tables",
            "allow_repeated_student_ids_across_placement_records",
            "add_session_last_saved_utc",
            "replace_last_saved_with_last_activity_utc",
        ))
        with Session(self.engine) as db:
            self.assertEqual(db.get(models.Pianist, 7).name, "Synthetic Pianist")
            self.assertEqual(db.query(TutorModuleState).count(), 1)
            self.assertEqual(TutorRepository(db).list_tutors(), [])

    def test_v14_upgrade_leaves_lesson_scheduling_data_untouched(self):
        from app import lesson_models
        from app.services.lesson_repository import LessonRepository

        lesson_repo = LessonRepository(self.db)
        lesson_repo.create_instructor("Synthetic Lesson Instructor")
        self.db.commit()
        self.db.close()
        with self.engine.begin() as connection:
            for table in sorted(TUTOR_TABLES - {"tutor_module_states"}, reverse=True):
                connection.exec_driver_sql(f"DROP TABLE {table}")
            connection.exec_driver_sql("DROP TABLE tutor_module_states")
            connection.exec_driver_sql("PRAGMA user_version = 14")

        migrate_database(self.engine)

        with Session(self.engine) as db:
            self.assertEqual(db.query(lesson_models.LessonInstructor).count(), 1)
            self.assertEqual(db.query(TutorTutor).count(), 0)

    def test_current_database_is_not_migrated_again(self):
        self.db.close()
        self.assertEqual(migrate_database(self.engine).applied_migrations, ())

    def test_module_state_is_keyed_by_session_uuid_for_session_rekeying(self):
        self.assertIn("session_uuid", tutor_models.TutorModuleState.__table__.c)


class TutorBoundaryTests(TutorTestCase):
    def test_tutor_tables_are_registered_with_the_module_metadata(self):
        self.assertTrue(TUTOR_TABLES <= set(ModuleBase.metadata.tables))

    def test_tutor_schema_has_no_foreign_keys_outside_tutor_tables(self):
        for name in TUTOR_TABLES:
            for foreign_key in ModuleBase.metadata.tables[name].foreign_keys:
                self.assertTrue(foreign_key.column.table.name.startswith("tutor_"), name)

    def test_no_other_module_references_tutor_tables(self):
        for name, table in ModuleBase.metadata.tables.items():
            if name in TUTOR_TABLES:
                continue
            for foreign_key in table.foreign_keys:
                self.assertFalse(foreign_key.column.table.name.startswith("tutor_"), name)

    def test_lesson_and_tutor_data_in_one_session_are_independent(self):
        from app.services.lesson_repository import LessonRepository

        lesson_repo = LessonRepository(self.db)
        lesson_repo.create_tag("voice")
        lesson_repo.create_instructor("Synthetic Lesson Instructor")
        self.assertEqual(self.repo.list_tags(), [])
        self.assertEqual(self.repo.list_tutors(), [])

        self.repo.create_tag("voice")
        self.repo.create_tutor("Synthetic Tutor")
        self.assertEqual(len(lesson_repo.list_tags()), 1)
        self.assertEqual([item.name for item in lesson_repo.list_instructors()], ["Synthetic Lesson Instructor"])
        self.assertEqual(lesson_repo.get_state().source_data_revision, 2)
        self.assertEqual(self.repo.get_state().source_data_revision, 2)

    def test_tutor_and_clinical_tags_are_independent_registries(self):
        self.assertNotEqual(tutor_models.TutorTag.__tablename__, clinical_models.ClinicalTag.__tablename__)
        clinical_repo = ClinicalRepository(self.db)
        clinical_repo.create_tag("voice")
        self.assertEqual(self.repo.list_tags(), [])
        tutor_tag = self.repo.create_tag("voice")
        self.assertEqual(len(clinical_repo.list_tags()), 1)
        self.repo.delete_tag(tutor_tag.id)
        self.assertEqual(len(clinical_repo.list_tags()), 1)

    def test_tutor_and_clinical_students_with_matching_identity_are_unrelated(self):
        clinical_repo = ClinicalRepository(self.db)
        clinical_student = clinical_repo.create_student("S-1", "Pat Test", "pat@example.test")
        tutor_student = self.repo.create_student("S-1", "Pat Test", "pat@example.test")
        self.assertNotEqual(clinical_student.id, tutor_student.id)
        self.repo.delete_student(tutor_student.id)
        self.assertEqual(len(clinical_repo.list_students()), 1)

    def test_tutor_changes_do_not_touch_clinical_state(self):
        clinical_state = ClinicalRepository(self.db).get_state()
        before = clinical_state.source_data_revision
        self.repo.create_tutor("T")
        self.repo.create_student("S-1", "Pat Test")
        self.assertEqual(clinical_state.source_data_revision, before)
        self.assertEqual(self.db.query(TutorTutor).count(), 1)


if __name__ == "__main__":
    unittest.main()
