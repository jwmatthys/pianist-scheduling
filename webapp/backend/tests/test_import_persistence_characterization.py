from datetime import time
from io import BytesIO
import json
import unittest
from unittest.mock import patch

from openpyxl import Workbook
from sqlalchemy import create_engine, text
from sqlalchemy.orm import Session, sessionmaker
from sqlalchemy.pool import StaticPool

from app import models, schemas
from app import database
from app.database import Base
from app.routers import assignments, imports, lessons, pianists
from app.services import importer


class ImportCharacterizationTests(unittest.TestCase):
    def setUp(self):
        importer._UPLOAD_CACHE.clear()

    def tearDown(self):
        importer._UPLOAD_CACHE.clear()

    def test_csv_mapping_normalizes_days_defaults_times_and_skips_invalid_rows(self):
        content = (
            "Teacher,Student,Day,Start,End,Room,Instrument,Needs,Required\n"
            "Prof One,Student One,R,9:15 AM,,Room A,Violin,,Ari Lee\n"
            "Prof Two,Student Two,FRI,1:30 PM,2:15 PM,Room B,Cello,0,\n"
            "Prof Three,Student Three,Funday,9:00 AM,9:50 AM,Room C,Voice,yes,\n"
            "Prof Four,Student Four,Monday,,10:00 AM,Room D,Piano,1,\n"
        ).encode()
        token, columns, preview = importer.stage_upload("synthetic.csv", content)
        mapping = {
            "teacher": "Teacher",
            "student": "Student",
            "day": "Day",
            "start_time": "Start",
            "end_time": "End",
            "room": "Room",
            "instrument": "Instrument",
            "need_pianist": "Needs",
            "required_pianist_name": "Required",
        }

        rows, warnings = importer.commit_upload(token, mapping)

        self.assertEqual(columns[0], "Teacher")
        self.assertEqual(len(preview), 4)
        self.assertEqual(len(rows), 2)
        self.assertEqual(len(warnings), 2)
        self.assertEqual(rows[0]["day"], "Thursday")
        self.assertEqual(rows[0]["start_minute"], 555)
        self.assertEqual(rows[0]["end_minute"], 605)
        self.assertTrue(rows[0]["need_pianist"])
        self.assertEqual(rows[0]["required_pianist_name"], "Ari Lee")
        self.assertEqual(rows[1]["day"], "Friday")
        self.assertEqual(rows[1]["start_minute"], 810)
        self.assertEqual(rows[1]["end_minute"], 855)
        self.assertFalse(rows[1]["need_pianist"])

    def test_xlsx_mapping_reads_excel_time_cells(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Day", "Start", "End", "Student"])
        sheet.append(["THU", time(10, 0), time(10, 45), "Student Example"])
        content = BytesIO()
        workbook.save(content)

        token, columns, _ = importer.stage_upload("synthetic.xlsx", content.getvalue())
        rows, warnings = importer.commit_upload(
            token,
            {"day": "Day", "start_time": "Start", "end_time": "End", "student": "Student"},
        )

        self.assertEqual(columns, ["Day", "Start", "End", "Student"])
        self.assertEqual(warnings, [])
        self.assertEqual(rows[0]["day"], "Thursday")
        self.assertEqual(rows[0]["start_minute"], 600)
        self.assertEqual(rows[0]["end_minute"], 645)


class PersistenceCharacterizationTests(unittest.TestCase):
    def setUp(self):
        importer._UPLOAD_CACHE.clear()
        self.engine = create_engine(
            "sqlite://",
            connect_args={"check_same_thread": False},
            poolclass=StaticPool,
        )
        Base.metadata.create_all(self.engine)
        self.db = Session(self.engine)
        self.db.add(models.Organization(id=1, name="Synthetic Test Program"))
        self.db.commit()

    def tearDown(self):
        importer._UPLOAD_CACHE.clear()
        self.db.close()
        self.engine.dispose()

    def test_import_replaces_lessons_and_persists_mapping_profile(self):
        self.db.add_all([
            models.Pianist(name="Ari Lee"),
            models.Lesson(
                teacher="Old Teacher",
                student="Old Student",
                day="Monday",
                start_minute=540,
                end_minute=590,
            ),
        ])
        self.db.commit()
        content = b"Student,Day,Start\nStudent New,T,11:00 AM\n"
        token, _, _ = importer.stage_upload("replacement.csv", content)
        mapping = {"student": "Student", "day": "Day", "start_time": "Start"}

        result = imports.commit(
            schemas.ImportCommit(
                upload_token=token,
                mapping=mapping,
                save_profile_name="Synthetic Lesson Columns",
            ),
            self.db,
        )

        self.assertEqual(result.created, 1)
        self.assertEqual(result.skipped, 0)
        stored = self.db.query(models.Lesson).all()
        self.assertEqual(len(stored), 1)
        self.assertEqual(stored[0].student, "Student New")
        self.assertEqual(stored[0].day, "Tuesday")
        self.assertEqual(stored[0].start_minute, 660)
        self.assertEqual(self.db.query(models.Pianist).count(), 1)
        profile = self.db.query(models.ImportProfile).one()
        self.assertEqual(profile.name, "Synthetic Lesson Columns")
        self.assertEqual(json.loads(profile.mapping_json), mapping)

        self.db.close()
        self.db = Session(self.engine)
        self.assertEqual(self.db.query(models.Lesson).one().student, "Student New")
        self.assertEqual(self.db.query(models.ImportProfile).one().name, "Synthetic Lesson Columns")

    def test_availability_bulk_update_replaces_and_persists_slots(self):
        pianist = models.Pianist(name="Bea Lin")
        self.db.add(pianist)
        self.db.commit()
        self.db.refresh(pianist)

        pianists.set_availability(
            pianist.id,
            schemas.AvailabilityBulkIn(slots=[
                schemas.AvailabilitySlotIn(
                    day="Monday", slot_start_minute=540, status="Available"
                ),
                schemas.AvailabilitySlotIn(
                    day="Monday", slot_start_minute=570, status="Tentative"
                ),
            ]),
            self.db,
        )
        pianists.set_availability(
            pianist.id,
            schemas.AvailabilityBulkIn(slots=[
                schemas.AvailabilitySlotIn(
                    day="Tuesday", slot_start_minute=600, status="Available"
                ),
            ]),
            self.db,
        )

        self.db.close()
        self.db = Session(self.engine)
        slots = self.db.query(models.AvailabilitySlot).all()
        self.assertEqual(len(slots), 1)
        self.assertEqual((slots[0].day, slots[0].slot_start_minute, slots[0].status),
                         ("Tuesday", 600, "Available"))
        self.assertTrue(self.db.query(models.Pianist).filter_by(name="Bea Lin").one().availability_complete)

    def test_database_startup_adds_legacy_columns_without_losing_rows(self):
        legacy_engine = create_engine(
            "sqlite://",
            connect_args={"check_same_thread": False},
            poolclass=StaticPool,
        )
        with legacy_engine.begin() as connection:
            connection.execute(text(
                "CREATE TABLE lessons ("
                "id INTEGER PRIMARY KEY, organization_id INTEGER NOT NULL DEFAULT 1, "
                "teacher VARCHAR(200) NOT NULL DEFAULT '', student VARCHAR(200) NOT NULL DEFAULT '', "
                "day VARCHAR(20) NOT NULL, start_minute INTEGER NOT NULL, end_minute INTEGER NOT NULL, "
                "room VARCHAR(200) NOT NULL DEFAULT '', instrument VARCHAR(200) NOT NULL DEFAULT '', "
                "required_pianist_name VARCHAR(200) NOT NULL DEFAULT '', need_pianist BOOLEAN NOT NULL DEFAULT 1, "
                "assigned_pianist_id INTEGER, fit_quality VARCHAR(20) NOT NULL DEFAULT '', "
                "notes TEXT NOT NULL DEFAULT '', hours FLOAT NOT NULL DEFAULT 0.0, "
                "manually_edited BOOLEAN NOT NULL DEFAULT 0)"
            ))
            connection.execute(text(
                "INSERT INTO lessons (id, teacher, student, day, start_minute, end_minute) "
                "VALUES (1, 'Teacher Example', 'Student Example', 'Monday', 540, 590)"
            ))

        legacy_sessions = sessionmaker(autocommit=False, autoflush=False, bind=legacy_engine)
        with patch.object(database, "engine", legacy_engine), patch.object(
            database, "SessionLocal", legacy_sessions
        ):
            database.init_db()

        with legacy_sessions() as migrated_db:
            migrated = migrated_db.get(models.Lesson, 1)
            self.assertIsNotNone(migrated)
            self.assertEqual(migrated.student, "Student Example")
            self.assertEqual(migrated.teacher_email, "")
            self.assertEqual(migrated.student_id, "")
            self.assertEqual(migrated_db.query(models.Organization).count(), 1)
        legacy_engine.dispose()

    def test_manual_assignments_survive_run_and_validation_persists_derived_state(self):
        pianist = models.Pianist(name="Ari Lee", max_hours_per_week=0.5)
        self.db.add(pianist)
        self.db.flush()
        self.db.add_all([
            models.AvailabilitySlot(
                pianist_id=pianist.id,
                day="Monday",
                slot_start_minute=540,
                status="Available",
            ),
            models.AvailabilitySlot(
                pianist_id=pianist.id,
                day="Monday",
                slot_start_minute=570,
                status="Available",
            ),
        ])
        first = models.Lesson(
            student="Student One", day="Monday", start_minute=540, end_minute=600
        )
        second = models.Lesson(
            student="Student Two", day="Monday", start_minute=540, end_minute=600
        )
        self.db.add_all([first, second])
        self.db.commit()
        self.db.refresh(pianist)
        self.db.refresh(first)
        self.db.refresh(second)

        lessons.update_lesson(
            first.id,
            schemas.LessonUpdate(assigned_pianist_id=pianist.id),
            self.db,
        )
        lessons.update_lesson(
            second.id,
            schemas.LessonUpdate(assigned_pianist_id=pianist.id),
            self.db,
        )
        validated = assignments.validate(self.db)

        self.assertEqual(len(validated.conflicts), 1)
        self.assertEqual(validated.hours_by_pianist["Ari Lee"], 1.0)
        self.assertEqual(sum(item.hours for item in validated.lessons), 1.0)
        self.assertTrue(all(item.manually_edited for item in validated.lessons))
        self.assertTrue(all("OVER CAP" in item.notes for item in validated.lessons))

        rerun = assignments.run_assignment(self.db)
        self.assertEqual([item.assigned_pianist_id for item in rerun.lessons],
                         [pianist.id, pianist.id])
        self.assertEqual(len(rerun.conflicts), 1)
        self.assertTrue(all(item.manually_edited for item in rerun.lessons))
        self.assertTrue(all("OVER CAP" in item.notes for item in rerun.lessons))

        self.db.close()
        self.db = Session(self.engine)
        persisted = self.db.query(models.Lesson).order_by(models.Lesson.id).all()
        self.assertEqual([item.assigned_pianist_id for item in persisted],
                         [pianist.id, pianist.id])
        self.assertTrue(all(item.manually_edited for item in persisted))

        assignments.clear_assignments(self.db)
        cleared = self.db.query(models.Lesson).order_by(models.Lesson.id).all()
        self.assertTrue(all(item.assigned_pianist_id is None for item in cleared))
        self.assertTrue(all(not item.manually_edited for item in cleared))

    def test_manual_edit_reclassifies_solver_overlap_as_conflict(self):
        pianist = models.Pianist(name="Ari Lee")
        self.db.add(pianist)
        self.db.flush()
        first = models.Lesson(
            student="Student One",
            day="Monday",
            start_minute=540,
            end_minute=600,
            assigned_pianist_id=pianist.id,
            fit_quality="Overlap",
        )
        second = models.Lesson(
            student="Student Two",
            day="Monday",
            start_minute=570,
            end_minute=630,
            assigned_pianist_id=pianist.id,
            fit_quality="Overlap",
        )
        self.db.add_all([first, second])
        self.db.commit()
        self.db.refresh(first)

        lessons.update_lesson(
            first.id,
            schemas.LessonUpdate(assigned_pianist_id=pianist.id),
            self.db,
        )
        validated = assignments.validate(self.db)

        self.assertEqual(first.fit_quality, "Manual")
        self.assertTrue(first.manually_edited)
        self.assertEqual(len(validated.conflicts), 1)
        self.assertIn("CONFLICT", first.notes)
        self.assertIn("CONFLICT", second.notes)


if __name__ == "__main__":
    unittest.main()
