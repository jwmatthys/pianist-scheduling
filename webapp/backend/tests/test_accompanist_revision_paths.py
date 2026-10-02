import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import models, module_models, schemas
from app.database import migrate_database
from app.routers import assignments, imports, lessons, pianists
from app.services.accompanist_results import finalize_accompanist_result, get_current_accompanist_result
from app.services.module_lifecycle import module_revision


class AccompanistRevisionPathTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.engine = create_engine(
            f"sqlite:///{Path(self.temporary_directory.name) / 'session.sqlite3'}"
        )
        migrate_database(self.engine)

    def tearDown(self):
        self.engine.dispose()
        self.temporary_directory.cleanup()

    def revision(self, db: Session) -> int:
        return module_revision(db).source_revision

    def test_lesson_pianist_assignment_and_validation_paths(self):
        with Session(self.engine) as db:
            lesson = lessons.create_lesson(
                schemas.LessonCreate(
                    teacher="Synthetic Teacher",
                    student="Synthetic Student",
                    student_id="S-REV-1",
                    day="Monday",
                    start_minute=540,
                    end_minute=590,
                ),
                db,
            )
            self.assertEqual(self.revision(db), 1)

            pianist = pianists.create_pianist(schemas.PianistCreate(name="Synthetic Pianist"), db)
            self.assertEqual(self.revision(db), 1)

            lessons.update_lesson(
                lesson.id,
                schemas.LessonUpdate(assigned_pianist_id=pianist.id),
                db,
            )
            self.assertEqual(self.revision(db), 2)

            pianists.update_pianist(
                pianist.id,
                schemas.PianistUpdate(name="Updated Synthetic Pianist"),
                db,
            )
            self.assertEqual(self.revision(db), 3)

            pianists.update_pianist(
                pianist.id,
                schemas.PianistUpdate(email="updated@example.invalid", max_hours_per_week=10),
                db,
            )
            self.assertEqual(self.revision(db), 3)

            assignments.validate(db)
            self.assertEqual(self.revision(db), 3)

            assignments.clear_assignments(db)
            self.assertEqual(self.revision(db), 4)

            def assign_synthetic(lessons_arg, pianists_arg, locked_ids=None):
                lessons_arg[0].assigned_pianist_id = pianist.id
                return {}, []

            with patch.object(assignments.scheduling, "assign_lessons", side_effect=assign_synthetic):
                assignments.run_assignment(db)
            self.assertEqual(self.revision(db), 5)

            result_row = db.query(module_models.ModuleResult).first()
            self.assertIsNone(result_row)

            lessons.update_lesson(
                lesson.id,
                schemas.LessonUpdate(instrument="Cello"),
                db,
            )
            self.assertEqual(self.revision(db), 6)

            lessons.delete_lesson(lesson.id, db)
            self.assertEqual(self.revision(db), 7)
            self.assertIsNone(db.get(module_models.AccompanistLessonIdentity, lesson.id))

    def test_lesson_requirement_field_edits_are_revisioned_and_independent_of_jury_required(self):
        with Session(self.engine) as db:
            lesson = lessons.create_lesson(schemas.LessonCreate(
                teacher="Synthetic Teacher",
                student="Requirement Field Student",
                day="Monday",
                start_minute=540,
                end_minute=590,
                need_pianist=False,
                jury_required=True,
            ), db)
            self.assertEqual(self.revision(db), 1)

            first_result = finalize_accompanist_result(db, expected_source_revision=1)
            first_result_row = db.get(module_models.ModuleResult, str(first_result.result_uuid))
            specific_edit = lessons.update_lesson(
                lesson.id,
                schemas.LessonUpdate(required_pianist_name="Susan Roberts"),
                db,
            )

            self.assertEqual(self.revision(db), 2)
            self.assertEqual(first_result_row.state, "superseded")
            self.assertFalse(specific_edit.need_pianist)
            self.assertEqual(specific_edit.required_pianist_name, "Susan Roberts")
            self.assertTrue(specific_edit.jury_required)

            second_result = finalize_accompanist_result(db, expected_source_revision=2)
            second_result_row = db.get(module_models.ModuleResult, str(second_result.result_uuid))
            needs_pianist_edit = lessons.update_lesson(
                lesson.id,
                schemas.LessonUpdate(need_pianist=True),
                db,
            )

            self.assertEqual(self.revision(db), 3)
            self.assertEqual(second_result_row.state, "superseded")
            self.assertTrue(needs_pianist_edit.need_pianist)
            self.assertEqual(needs_pianist_edit.required_pianist_name, "Susan Roberts")
            self.assertTrue(needs_pianist_edit.jury_required)

    def test_pianist_required_toggles_solver_eligibility_without_erasing_assignments(self):
        with Session(self.engine) as db:
            pianist = pianists.create_pianist(schemas.PianistCreate(name="Synthetic Available Pianist"), db)
            pianists.set_availability(pianist.id, schemas.AvailabilityBulkIn(slots=[
                schemas.AvailabilitySlotIn(day="Monday", slot_start_minute=540, status="Available"),
                schemas.AvailabilitySlotIn(day="Monday", slot_start_minute=570, status="Available"),
            ]), db)
            lesson = lessons.create_lesson(schemas.LessonCreate(
                student="Synthetic Required Toggle",
                day="Monday",
                start_minute=540,
                end_minute=590,
                need_pianist=False,
            ), db)

            not_required = assignments.run_assignment(db)
            self.assertIsNone(not_required.lessons[0].assigned_pianist_id)
            self.assertEqual(not_required.unassigned_count, 0)

            lessons.update_lesson(lesson.id, schemas.LessonUpdate(need_pianist=True), db)
            required = assignments.run_assignment(db)
            self.assertEqual(required.lessons[0].assigned_pianist_id, pianist.id)
            self.assertEqual(required.unassigned_count, 0)

            lessons.update_lesson(lesson.id, schemas.LessonUpdate(need_pianist=False), db)
            no_longer_required = assignments.run_assignment(db)
            self.assertEqual(no_longer_required.lessons[0].assigned_pianist_id, pianist.id)
            self.assertEqual(no_longer_required.unassigned_count, 0)

    def test_import_replacement_preserves_nonblank_student_identity_and_bumps_once(self):
        with Session(self.engine) as db:
            existing = lessons.create_lesson(
                schemas.LessonCreate(
                    student="Synthetic Student",
                    student_id=" S-IMPORT-1 ",
                    day="Monday",
                    start_minute=540,
                    end_minute=590,
                ),
                db,
            )
            existing_identity = db.get(module_models.AccompanistLessonIdentity, existing.id)
            original_person_uuid = existing_identity.student_person_uuid
            revision_before = self.revision(db)

            imported = [{
                "teacher": "New Teacher",
                "teacher_email": "",
                "student": "Synthetic Student",
                "student_id": "S-IMPORT-1",
                "day": "Tuesday",
                "start_minute": 600,
                "end_minute": 650,
                "room": "",
                "instrument": "Voice",
                "required_pianist_name": "",
                "need_pianist": False,
            }]
            with patch.object(imports.importer, "commit_upload", return_value=(imported, [])):
                imports.commit(schemas.ImportCommit(upload_token="synthetic", mapping={}), db)

            current_identity = db.query(module_models.AccompanistLessonIdentity).one()
            self.assertEqual(current_identity.student_person_uuid, original_person_uuid)
            self.assertNotEqual(current_identity.lesson_uuid, existing_identity.lesson_uuid)
            self.assertEqual(self.revision(db), revision_before + 1)

    def test_jury_required_edit_is_source_revisioned_and_published_in_contract_v2(self):
        with Session(self.engine) as db:
            lesson = lessons.create_lesson(
                schemas.LessonCreate(
                    teacher="Synthetic Teacher",
                    student="Synthetic Jury Student",
                    student_id="S-JURY-REV",
                    day="Monday",
                    start_minute=540,
                    end_minute=590,
                    need_pianist=False,
                    jury_required=False,
                ),
                db,
            )
            self.assertFalse(lesson.jury_required)
            self.assertEqual(self.revision(db), 1)

            updated = lessons.update_lesson(
                lesson.id,
                schemas.LessonUpdate(jury_required=True),
                db,
            )
            self.assertTrue(updated.jury_required)
            self.assertEqual(self.revision(db), 2)

            finalized = finalize_accompanist_result(db, expected_source_revision=2)
            self.assertEqual(finalized.contract_version, 2)
            self.assertTrue(finalized.payload.entries[0].jury_required)
            db.commit()

            unchanged = lessons.update_lesson(
                lesson.id,
                schemas.LessonUpdate(jury_required=True),
                db,
            )
            self.assertTrue(unchanged.jury_required)
            self.assertEqual(self.revision(db), 2)

            no_longer_required = lessons.update_lesson(
                lesson.id,
                schemas.LessonUpdate(jury_required=False),
                db,
            )
            self.assertFalse(no_longer_required.jury_required)
            self.assertEqual(self.revision(db), 3)
            self.assertIsNone(get_current_accompanist_result(db))

            next_result = finalize_accompanist_result(db, expected_source_revision=3)
            self.assertFalse(next_result.payload.entries[0].jury_required)


if __name__ == "__main__":
    unittest.main()