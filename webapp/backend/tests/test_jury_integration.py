from datetime import date
import tempfile
import unittest
from pathlib import Path

from pydantic import ValidationError
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import models, module_models
from app.database import migrate_database
from app.jury_schemas import JuryAvailabilityIn, JuryPanelFields
from app.services.accompanist_results import (
    ResultPublicationError,
    finalize_accompanist_result,
    get_current_accompanist_result,
)
from app.services.jury import (
    JuryDataError,
    create_panel,
    readiness,
    save_availability,
    sync_roster_from_current_result,
    update_configuration,
    update_lesson_entry,
)
from app.services.module_lifecycle import (
    bump_accompanist_revision,
    ensure_pianist_identity,
    synchronize_lesson_identity,
)
from app.routers.lessons import delete_lesson


class JuryIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.path = Path(self.temporary_directory.name) / "session.sqlite3"
        self.engine = create_engine(f"sqlite:///{self.path}")
        migrate_database(self.engine)

    def tearDown(self):
        self.engine.dispose()
        self.temporary_directory.cleanup()

    def add_pianist(self, db: Session, name: str) -> tuple[models.Pianist, str]:
        pianist = models.Pianist(name=name)
        db.add(pianist)
        db.flush()
        return pianist, ensure_pianist_identity(db, pianist)

    def add_lesson(
        self,
        db: Session,
        *,
        student: str,
        student_id: str = "S-1",
        instrument: str = "Voice",
        teacher: str = "Synthetic Teacher",
        pianist: models.Pianist | None = None,
        pianist_required: bool = True,
    ) -> tuple[models.Lesson, str]:
        lesson = models.Lesson(
            student=student,
            student_id=student_id,
            instrument=instrument,
            teacher=teacher,
            need_pianist=pianist_required,
            assigned_pianist_id=pianist.id if pianist else None,
            day="Monday",
            start_minute=540,
            end_minute=590,
        )
        db.add(lesson)
        db.flush()
        student_person_uuid = synchronize_lesson_identity(
            db,
            lesson,
            identity_fields_changed=True,
        )
        return lesson, student_person_uuid

    def test_result_keeps_multiple_lesson_facts_and_pianists_distinct(self):
        with Session(self.engine) as db:
            pianist_one, pianist_one_uuid = self.add_pianist(db, "Synthetic Pianist One")
            pianist_two, pianist_two_uuid = self.add_pianist(db, "Synthetic Pianist Two")
            first, student_uuid = self.add_lesson(
                db,
                student="Alex Student",
                instrument="Voice",
                pianist=pianist_one,
            )
            second, second_student_uuid = self.add_lesson(
                db,
                student=" ALEX   STUDENT ",
                student_id=" S-1 ",
                instrument="Cello",
                teacher="Different Teacher",
                pianist=pianist_two,
            )
            self.assertEqual(student_uuid, second_student_uuid)
            result = finalize_accompanist_result(db, expected_source_revision=0)
            db.commit()

            self.assertEqual(len(result.payload.entries), 2)
            entries = {entry.instrument: entry for entry in result.payload.entries}
            voice = entries["Voice"]
            cello = entries["Cello"]
            self.assertNotEqual(voice.source_lesson_uuid, cello.source_lesson_uuid)
            self.assertEqual(voice.student_person_uuid, cello.student_person_uuid)
            self.assertEqual(voice.teacher, "Synthetic Teacher")
            self.assertEqual(cello.teacher, "Different Teacher")
            self.assertEqual(str(voice.assigned_pianist.person_uuid), pianist_one_uuid)
            self.assertEqual(str(cello.assigned_pianist.person_uuid), pianist_two_uuid)

            result_row = db.get(module_models.ModuleResult, str(result.result_uuid))
            payload_before = result_row.payload_json
            bump_accompanist_revision(db)
            db.commit()
            self.assertEqual(result_row.state, "superseded")
            self.assertEqual(result_row.payload_json, payload_before)
            self.assertIsNone(get_current_accompanist_result(db))

    def test_identity_conflict_blocks_finalization_and_blank_names_stay_distinct(self):
        with Session(self.engine) as db:
            self.add_lesson(db, student="Alex Student", student_id="CONFLICT-1")
            self.add_lesson(db, student="Different Student", student_id=" CONFLICT-1 ")
            _, blank_one = self.add_lesson(db, student="Same Blank", student_id="")
            _, blank_two = self.add_lesson(db, student="Same Blank", student_id="   ")
            self.assertNotEqual(blank_one, blank_two)
            with self.assertRaises(ResultPublicationError) as raised:
                finalize_accompanist_result(db, expected_source_revision=0)
            self.assertEqual(raised.exception.code, "STUDENT_IDENTITY_CONFLICT")

    def test_blank_id_lesson_identity_remains_stable_when_its_name_changes(self):
        with Session(self.engine) as db:
            lesson, person_uuid = self.add_lesson(db, student="Original Name", student_id="")
            lesson.student = "Updated Name"
            db.flush()
            updated_person_uuid = synchronize_lesson_identity(
                db,
                lesson,
                identity_fields_changed=True,
            )
            self.assertEqual(person_uuid, updated_person_uuid)

    def test_deleting_conflicting_lesson_recomputes_student_identity_conflict(self):
        with Session(self.engine) as db:
            first, student_uuid = self.add_lesson(
                db,
                student="Original Student",
                student_id="CONFLICT-DELETE",
            )
            self.add_lesson(
                db,
                student="Different Student",
                student_id="CONFLICT-DELETE",
            )
            profile = db.get(module_models.AccompanistStudentProfile, student_uuid)
            self.assertTrue(profile.identity_conflict)

            delete_lesson(first.id, db)

            self.assertFalse(profile.identity_conflict)
            result = finalize_accompanist_result(db, expected_source_revision=1)
            self.assertEqual(len(result.payload.entries), 1)

    def test_revision_compare_and_swap_blocks_stale_finalize(self):
        with Session(self.engine) as db:
            self.add_lesson(db, student="Synthetic Student", pianist_required=False)
            bump_accompanist_revision(db)
            db.commit()
            with self.assertRaises(ResultPublicationError) as raised:
                finalize_accompanist_result(db, expected_source_revision=0)
            self.assertEqual(raised.exception.code, "SOURCE_REVISION_CHANGED")

    def test_panel_defaults_and_validation(self):
        fields = JuryPanelFields(
            panel_name="Voice A",
            earliest_start_minute=540,
            jury_length_minutes=60,
            break_needed=True,
            break_every_x_juries=4,
        )
        self.assertEqual(fields.preferred_start_minute, 540)
        self.assertEqual(fields.break_length_minutes, 60)
        updated = JuryPanelFields(
            **{
                **fields.model_dump(),
                "jury_length_minutes": 45,
            }
        )
        self.assertEqual(updated.break_length_minutes, 60)
        with self.assertRaises(ValidationError):
            JuryPanelFields(
                panel_name="Invalid",
                earliest_start_minute=600,
                preferred_start_minute=599,
                jury_length_minutes=30,
            )
        with self.assertRaises(ValidationError):
            JuryPanelFields(
                panel_name="Invalid Meal",
                earliest_start_minute=600,
                jury_length_minutes=30,
                meal_break=True,
                meal_start_minute=700,
                meal_end_minute=690,
            )

    def test_duplicate_rooms_do_not_create_readiness_warnings(self):
        with Session(self.engine) as db:
            self.add_lesson(
                db,
                student="Synthetic Student One",
                student_id="ROOM-1",
                pianist_required=False,
            )
            self.add_lesson(
                db,
                student="Synthetic Student Two",
                student_id="ROOM-2",
                pianist_required=False,
            )
            result = finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            update_configuration(db, date(2027, 5, 1))
            panel_one = create_panel(db, JuryPanelFields(
                panel_name="Panel One",
                room="Shared Room",
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            panel_two = create_panel(db, JuryPanelFields(
                panel_name="Panel Two",
                room="Shared Room",
                earliest_start_minute=540,
                jury_length_minutes=60,
            ))
            source_by_student = {
                entry.student_display_name: entry
                for entry in result.payload.entries
            }
            update_lesson_entry(
                db,
                str(source_by_student["Synthetic Student One"].source_lesson_uuid),
                jury_required=True,
                panel_uuid=str(panel_one.panel_uuid),
            )
            update_lesson_entry(
                db,
                str(source_by_student["Synthetic Student Two"].source_lesson_uuid),
                jury_required=True,
                panel_uuid=str(panel_two.panel_uuid),
            )
            db.commit()

            report = readiness(db)

            self.assertTrue(report.ready, report.issues)
            self.assertNotIn("DUPLICATE_ROOM", {issue.code for issue in report.issues})

    def test_jury_roster_preserves_settings_and_readiness_uses_date_scoped_availability(self):
        jury_date = date(2027, 5, 1)
        with Session(self.engine) as db:
            pianist, pianist_uuid = self.add_pianist(db, "Synthetic Accompanist")
            lesson, _ = self.add_lesson(
                db,
                student="Synthetic Student",
                pianist=pianist,
            )
            result = finalize_accompanist_result(db, expected_source_revision=0)
            db.flush()
            entries = sync_roster_from_current_result(db)
            self.assertEqual(len(entries), 1)
            source_lesson_uuid = str(result.payload.entries[0].source_lesson_uuid)
            configuration = update_configuration(db, jury_date)
            panel = create_panel(db, JuryPanelFields(
                panel_name="Panel One",
                room="Room A",
                earliest_start_minute=540,
                jury_length_minutes=60,
                break_needed=True,
                break_every_x_juries=3,
            ))
            update_lesson_entry(
                db,
                source_lesson_uuid,
                jury_required=True,
                panel_uuid=str(panel.panel_uuid),
            )
            db.commit()

            blocked = readiness(db)
            self.assertFalse(blocked.ready)
            self.assertIn("PIANIST_AVAILABILITY_INCOMPLETE", {issue.code for issue in blocked.issues})

            save_availability(
                db,
                pianist_uuid,
                jury_date,
                JuryAvailabilityIn(is_complete=True, windows=[]),
            )
            db.commit()
            ready = readiness(db)
            self.assertTrue(ready.ready, ready.issues)

            # A roster refresh to the same result must retain Jury-owned choices.
            sync_roster_from_current_result(db)
            db.commit()
            row = db.get(module_models.JuryLessonEntry, (str(configuration.session_uuid), source_lesson_uuid))
            self.assertTrue(row.jury_required)
            self.assertEqual(row.panel_uuid, str(panel.panel_uuid))

            next_date = date(2027, 5, 2)
            update_configuration(db, next_date)
            db.commit()
            changed_date = readiness(db)
            self.assertFalse(changed_date.ready)
            self.assertIn("PIANIST_AVAILABILITY_INCOMPLETE", {issue.code for issue in changed_date.issues})
            self.assertIsNotNone(db.get(
                module_models.JuryPianistAvailabilityDeclaration,
                (str(configuration.session_uuid), pianist_uuid, jury_date),
            ))

    def test_availability_rejects_overlaps_and_wrong_date(self):
        with Session(self.engine) as db:
            _, pianist_uuid = self.add_pianist(db, "Synthetic Pianist")
            jury_date = date(2027, 5, 1)
            update_configuration(db, jury_date)
            payload = JuryAvailabilityIn.model_validate({
                "is_complete": True,
                "windows": [
                    {"start_minute": 500, "end_minute": 600},
                    {"start_minute": 590, "end_minute": 650},
                ],
            })
            with self.assertRaises(JuryDataError) as raised:
                save_availability(db, pianist_uuid, jury_date, payload)
            self.assertEqual(raised.exception.code, "OVERLAPPING_AVAILABILITY")
            with self.assertRaises(JuryDataError) as raised:
                save_availability(db, pianist_uuid, date(2027, 5, 2), JuryAvailabilityIn(is_complete=False))
            self.assertEqual(raised.exception.code, "JURY_DATE_MISMATCH")

if __name__ == "__main__":
    unittest.main()