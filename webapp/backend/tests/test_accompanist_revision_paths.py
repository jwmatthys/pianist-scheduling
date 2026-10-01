import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import models, module_models, schemas
from app.database import migrate_database
from app.routers import assignments, imports, lessons, pianists
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


if __name__ == "__main__":
    unittest.main()