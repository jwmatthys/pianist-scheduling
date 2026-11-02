"""Tutor Scheduling data must survive .mpsession export and restore. All data is synthetic."""

import tempfile
import unittest
from datetime import date
from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import module_models, schemas
from app.database import migrate_database
from app.routers import tutor_scheduling as tutor_api, reports
from app.services.tutor_repository import TutorRepository
from app.services.session_files import create_new_session, export_session_archive, restore_session
from tests.test_tutor_foundation import window


def metadata(term="Fall 2026"):
    return schemas.SessionMetadataIn(
        institution_name="Synthetic University",
        program_name="Music",
        term_label=term,
        year=2026,
        start_date=date(2026, 8, 20),
        end_date=date(2026, 12, 20),
    )


def snapshot(db: Session) -> dict:
    repo = TutorRepository(db)
    schedule = tutor_api.get_schedule(db)
    state = repo.get_state()
    return {
        "tags": sorted(tag.name for tag in repo.list_tags()),
        "tutors": [
            (
                item.name,
                item.room,
                item.lesson_length_minutes,
                item.capacity,
                sorted((w.day, w.start_minute, w.end_minute, w.status) for w in repo.tutor_availability(item.id)),
                sorted(tag.name for tag in item.tags),
            )
            for item in repo.list_tutors()
        ],
        "students": [
            (
                item.student_id,
                item.name,
                sorted((w.day, w.start_minute, w.end_minute, w.status) for w in repo.student_availability(item.id)),
                sorted((tag_id, str(state)) for tag_id, state in repo.student_tag_states().get(item.id, {}).items()),
            )
            for item in repo.list_students()
        ],
        "entries": sorted(
            (e.student_id, e.placement_type, e.placement_state, e.tutor_name, e.day,
             e.start_minute, e.end_minute, e.unplaced_reason)
            for e in schedule.entries
        ),
        "revisions": (state.schedule_revision, state.source_data_revision, state.is_schedule_stale),
        "counts": (schedule.automatic_count, schedule.manual_count, schedule.manually_unplaced_count),
    }


class TutorSessionArchiveTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.directory = Path(self.temporary_directory.name)
        self.engine = create_engine(
            f"sqlite:///{self.directory / 'active.sqlite3'}", connect_args={"check_same_thread": False}
        )
        migrate_database(self.engine)
        self.recovery = self.directory / "recovery"

    def tearDown(self):
        self.engine.dispose()
        self.temporary_directory.cleanup()

    def test_tutor_data_schedule_state_and_saved_report_round_trip(self):
        original = create_new_session(self.engine, metadata(), self.recovery)
        with Session(self.engine) as db:
            repo = TutorRepository(db)
            tag = repo.create_tag("Synthetic Tag")
            repo.create_tutor(
                "Synthetic Tutor", room="Studio 1", lesson_length_minutes=60, capacity=1,
                availability=[window(start=540, end=660)], tag_ids=[tag.id],
            )
            for number in range(3):
                repo.create_student(
                    f"S-{number}", f"Synthetic Student {number}",
                    availability=[window(start=540, end=660, status="Tentative" if number else "Available")],
                    tag_states={tag.id: "preferred"},
                )
            schedule = tutor_api.generate_schedule(db)
            unplaced = next(e.student_id for e in schedule.entries if e.day is None)
            tutor_api.manually_unplace_student(unplaced, db)
            reports.put_saved_report(
                schemas.SavedReportIn(
                    markdown="# Tutor Schedule",
                    page_break_headings=["Unplaced Students"],
                    table=schemas.SavedReportTable(
                        name="Tutor Schedule", columns=["Student"], rows=[["Synthetic Student 0"]]
                    ),
                    source_stamp="1:2",
                ),
                "tutor",
                db,
            )
            before = snapshot(db)
        self.assertTrue(before["tutors"] and before["students"] and before["entries"])
        self.assertEqual(before["counts"][2], 1)

        exported = export_session_archive(self.engine)

        with Session(self.engine) as db:
            repo = TutorRepository(db)
            repo.create_student("S-99", "Synthetic Late Student", availability=[window()])
            for student in repo.list_students():
                if student.student_id == "S-0":
                    repo.delete_student(student.id)
        create_new_session(self.engine, metadata("Spring 2027"), self.recovery)

        restored = restore_session(self.engine, exported, self.recovery)

        self.assertEqual(restored.session_uuid, original.session_uuid)
        with Session(self.engine) as db:
            self.assertEqual(snapshot(db), before)
            saved = reports.get_saved_report("tutor", db)
            self.assertEqual(saved.markdown, "# Tutor Schedule")
            self.assertEqual(saved.source_stamp, "1:2")
            self.assertEqual(db.query(module_models.SavedReport).count(), 1)
        self.assertGreaterEqual(len(list(self.recovery.glob("*.mpsession"))), 2)


if __name__ == "__main__":
    unittest.main()
