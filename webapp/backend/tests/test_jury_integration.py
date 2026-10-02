from datetime import date
import tempfile
import unittest
from pathlib import Path
from uuid import UUID

from pydantic import ValidationError
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import models, module_models, schemas
from app.database import migrate_database
from app.jury_schemas import JuryAvailabilityIn, JuryGenerateRequest, JuryPanelFields
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
    update_lesson_entry,
    update_lesson_jury_required,
    update_panel,
)
from app.services.module_lifecycle import (
    bump_accompanist_revision,
    ensure_pianist_identity,
    set_lesson_jury_required,
    synchronize_lesson_identity,
)
from app.routers import jury as jury_router
from app.routers.lessons import delete_lesson, update_lesson


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
        jury_required: bool = False,
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
        set_lesson_jury_required(db, lesson.id, jury_required)
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
                jury_required=True,
            )
            second, second_student_uuid = self.add_lesson(
                db,
                student=" ALEX   STUDENT ",
                student_id=" S-1 ",
                instrument="Cello",
                teacher="Different Teacher",
                pianist=pianist_two,
                jury_required=False,
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
            self.assertTrue(voice.jury_required)
            self.assertFalse(cello.jury_required)
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
            jury_date=date(2027, 5, 1),
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
                jury_date=date(2027, 5, 1),
                earliest_start_minute=600,
                preferred_start_minute=599,
                jury_length_minutes=30,
            )
        with self.assertRaises(ValidationError):
            JuryPanelFields(
                panel_name="Invalid Meal",
                jury_date=date(2027, 5, 1),
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
                jury_required=True,
            )
            self.add_lesson(
                db,
                student="Synthetic Student Two",
                student_id="ROOM-2",
                pianist_required=False,
                jury_required=True,
            )
            result = finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panel_one = create_panel(db, JuryPanelFields(
                panel_name="Panel One",
                room="Shared Room",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            panel_two = create_panel(db, JuryPanelFields(
                panel_name="Panel Two",
                room="Shared Room",
                jury_date=date(2027, 5, 2),
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
                panel_uuid=str(panel_one.panel_uuid),
            )
            update_lesson_entry(
                db,
                str(source_by_student["Synthetic Student Two"].source_lesson_uuid),
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
                jury_required=True,
            )
            result = finalize_accompanist_result(db, expected_source_revision=0)
            db.flush()
            entries = sync_roster_from_current_result(db)
            self.assertEqual(len(entries), 1)
            source_lesson_uuid = str(result.payload.entries[0].source_lesson_uuid)
            session_uuid = db.query(models.SchedulingSession).one().session_uuid
            panel = create_panel(db, JuryPanelFields(
                panel_name="Panel One",
                room="Room A",
                jury_date=jury_date,
                earliest_start_minute=540,
                jury_length_minutes=60,
                break_needed=True,
                break_every_x_juries=3,
            ))
            update_lesson_entry(
                db,
                source_lesson_uuid,
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
                JuryAvailabilityIn(windows=[{"start_minute": 540, "end_minute": 600}]),
            )
            db.commit()
            ready = readiness(db)
            self.assertTrue(ready.ready, ready.issues)

            # A roster refresh to the same result must retain Jury-owned choices.
            refreshed_entries = sync_roster_from_current_result(db)
            db.commit()
            row = db.get(module_models.JuryLessonEntry, (session_uuid, source_lesson_uuid))
            self.assertEqual(row.panel_uuid, str(panel.panel_uuid))
            self.assertTrue(refreshed_entries[0].jury_required)

            next_date = date(2027, 5, 2)
            update_panel(db, str(panel.panel_uuid), JuryPanelFields(
                panel_name=panel.panel_name,
                room=panel.room,
                jury_date=next_date,
                earliest_start_minute=panel.earliest_start_minute,
                preferred_start_minute=panel.preferred_start_minute,
                jury_length_minutes=panel.jury_length_minutes,
                break_needed=panel.break_needed,
                break_every_x_juries=panel.break_every_x_juries,
                break_length_minutes=panel.break_length_minutes,
                meal_break=panel.meal_break,
                meal_start_minute=panel.meal_start_minute,
                meal_end_minute=panel.meal_end_minute,
            ))
            db.commit()
            changed_date = readiness(db)
            self.assertFalse(changed_date.ready)
            self.assertIn("PIANIST_AVAILABILITY_INCOMPLETE", {issue.code for issue in changed_date.issues})
            self.assertIsNotNone(db.get(
                module_models.JuryPianistAvailabilityDeclaration,
                (session_uuid, pianist_uuid, jury_date),
            ))

    def test_no_pianist_jury_is_valid_and_nonjury_lesson_is_excluded(self):
        with Session(self.engine) as db:
            required, _ = self.add_lesson(
                db,
                student="Jury Without Piano",
                student_id="NO-PIANO-JURY",
                pianist_required=False,
                jury_required=True,
            )
            excluded, _ = self.add_lesson(
                db,
                student="Recital Pianist Need Only",
                student_id="NOT-A-JURY",
                pianist_required=True,
                jury_required=False,
            )
            result = finalize_accompanist_result(db, expected_source_revision=0)
            entries = sync_roster_from_current_result(db)
            self.assertEqual(
                {entry.student_display_name: entry.jury_required for entry in entries},
                {"Jury Without Piano": True, "Recital Pianist Need Only": False},
            )
            panel = create_panel(db, JuryPanelFields(
                panel_name="Voice Panel",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            source_by_name = {entry.student_display_name: entry for entry in result.payload.entries}
            update_lesson_entry(
                db,
                str(source_by_name["Jury Without Piano"].source_lesson_uuid),
                panel_uuid=str(panel.panel_uuid),
            )
            db.commit()

            report = readiness(db)

            self.assertTrue(report.ready, report.issues)
            self.assertNotIn("FINALIZED_PIANIST_REQUIRED", {issue.code for issue in report.issues})
            self.assertNotIn("PIANIST_AVAILABILITY_INCOMPLETE", {issue.code for issue in report.issues})
            self.assertEqual(required.need_pianist, False)
            self.assertEqual(excluded.need_pianist, True)

    def test_required_fixed_pianist_without_finalized_assignment_blocks(self):
        with Session(self.engine) as db:
            _, _ = self.add_lesson(
                db,
                student="Missing Fixed Pianist",
                student_id="MISSING-PIANIST",
                pianist_required=True,
                jury_required=True,
            )
            result = finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panel = create_panel(db, JuryPanelFields(
                panel_name="Panel One",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            update_lesson_entry(
                db,
                str(result.payload.entries[0].source_lesson_uuid),
                panel_uuid=str(panel.panel_uuid),
            )
            db.commit()

            report = readiness(db)

            self.assertFalse(report.ready)
            self.assertIn("FINALIZED_PIANIST_REQUIRED", {issue.code for issue in report.issues})

    def test_new_source_results_refresh_jury_required_and_preserve_panel(self):
        with Session(self.engine) as db:
            lesson, _ = self.add_lesson(
                db,
                student="Changing Jury Requirement",
                student_id="CHANGING-JURY",
                pianist_required=False,
                jury_required=False,
            )
            first_result = finalize_accompanist_result(db, expected_source_revision=0)
            first_entries = sync_roster_from_current_result(db)
            source_uuid = str(first_result.payload.entries[0].source_lesson_uuid)
            self.assertFalse(first_entries[0].jury_required)

            update_lesson(lesson.id, schemas.LessonUpdate(jury_required=True), db)
            second_result = finalize_accompanist_result(db, expected_source_revision=1)
            second_entries = sync_roster_from_current_result(db)
            self.assertTrue(second_entries[0].jury_required)
            self.assertEqual(str(second_entries[0].source_lesson_uuid), source_uuid)
            self.assertIsNone(second_entries[0].panel_uuid)

            panel = create_panel(db, JuryPanelFields(
                panel_name="Persistent Panel",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            update_lesson_entry(db, source_uuid, panel_uuid=str(panel.panel_uuid))
            update_lesson(lesson.id, schemas.LessonUpdate(instrument="Cello"), db)
            third_result = finalize_accompanist_result(db, expected_source_revision=2)
            third_entries = sync_roster_from_current_result(db)
            self.assertEqual(third_result.payload.entries[0].instrument, "Cello")
            self.assertTrue(third_entries[0].jury_required)
            self.assertEqual(third_entries[0].panel_uuid, panel.panel_uuid)

            update_lesson(lesson.id, schemas.LessonUpdate(jury_required=False), db)
            fourth_result = finalize_accompanist_result(db, expected_source_revision=3)
            fourth_entries = sync_roster_from_current_result(db)
            self.assertFalse(fourth_result.payload.entries[0].jury_required)
            self.assertFalse(fourth_entries[0].jury_required)
            self.assertEqual(fourth_entries[0].panel_uuid, panel.panel_uuid)

    def test_availability_rejects_overlaps_and_wrong_date(self):
        with Session(self.engine) as db:
            _, pianist_uuid = self.add_pianist(db, "Synthetic Pianist")
            jury_date = date(2027, 5, 1)
            create_panel(db, JuryPanelFields(
                panel_name="Availability Date Panel",
                jury_date=jury_date,
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            payload = JuryAvailabilityIn.model_validate({
                "windows": [
                    {"start_minute": 500, "end_minute": 600},
                    {"start_minute": 590, "end_minute": 650},
                ],
            })
            with self.assertRaises(JuryDataError) as raised:
                save_availability(db, pianist_uuid, jury_date, payload)
            self.assertEqual(raised.exception.code, "OVERLAPPING_AVAILABILITY")
            with self.assertRaises(JuryDataError) as raised:
                save_availability(db, pianist_uuid, date(2027, 5, 2), JuryAvailabilityIn())
            self.assertEqual(raised.exception.code, "JURY_DATE_MISMATCH")

    def test_jury_required_edit_updates_source_and_preserves_panel_selection(self):
        with Session(self.engine) as db:
            pianist, _ = self.add_pianist(db, "Synthetic Toggle Pianist")
            lesson, _ = self.add_lesson(
                db,
                student="Synthetic Toggle Student",
                pianist=pianist,
                pianist_required=True,
                jury_required=True,
            )
            lesson_uuid = db.get(module_models.AccompanistLessonIdentity, lesson.id).lesson_uuid
            finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panel = create_panel(db, JuryPanelFields(
                panel_name="Retained Panel",
                room="Room A",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=60,
            ))
            update_lesson_entry(db, lesson_uuid, panel_uuid=str(panel.panel_uuid))
            db.commit()

            turned_off = update_lesson_jury_required(db, lesson_uuid, False)
            db.commit()
            self.assertFalse(turned_off.jury_required)
            self.assertEqual(str(turned_off.panel_uuid), str(panel.panel_uuid))
            requirement = db.get(module_models.AccompanistLessonJuryRequirement, lesson.id)
            self.assertFalse(requirement.jury_required)
            self.assertIn("ACCOMPANIST_RESULT_STALE", {issue.code for issue in readiness(db).issues})

            turned_on = update_lesson_jury_required(db, lesson_uuid, True)
            db.commit()
            self.assertTrue(turned_on.jury_required)
            self.assertEqual(str(turned_on.panel_uuid), str(panel.panel_uuid))
            issue_codes = {issue.code for issue in readiness(db).issues}
            self.assertIn("ACCOMPANIST_RESULT_STALE", issue_codes)
            self.assertIn("PIANIST_AVAILABILITY_INCOMPLETE", issue_codes)

    def test_readiness_does_not_require_availability_for_unassigned_panel_dates(self):
        with Session(self.engine) as db:
            pianist, pianist_uuid = self.add_pianist(db, "Date Scoped Pianist")
            lesson, _ = self.add_lesson(
                db,
                student="Date Scoped Student",
                pianist=pianist,
                pianist_required=True,
                jury_required=True,
            )
            lesson_uuid = db.get(module_models.AccompanistLessonIdentity, lesson.id).lesson_uuid
            finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            assigned_panel = create_panel(db, JuryPanelFields(
                panel_name="Assigned Date Panel",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=60,
            ))
            create_panel(db, JuryPanelFields(
                panel_name="Unused Date Panel",
                jury_date=date(2027, 5, 2),
                earliest_start_minute=540,
                jury_length_minutes=60,
            ))
            update_lesson_entry(db, lesson_uuid, panel_uuid=str(assigned_panel.panel_uuid))
            save_availability(db, pianist_uuid, date(2027, 5, 1), JuryAvailabilityIn(
                windows=[{"start_minute": 540, "end_minute": 600}],
            ))
            db.commit()

            result = readiness(db)
            self.assertTrue(result.ready, result.issues)
            self.assertNotIn("PIANIST_AVAILABILITY_INCOMPLETE", {issue.code for issue in result.issues})

    def test_missing_panel_reports_panel_blocker_without_invalid_availability_entity(self):
        with Session(self.engine) as db:
            pianist, _ = self.add_pianist(db, "No Panel Pianist")
            lesson, _ = self.add_lesson(
                db,
                student="No Panel Student",
                pianist=pianist,
                pianist_required=True,
                jury_required=True,
            )
            finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)

            result = readiness(db)

            issue_codes = {issue.code for issue in result.issues}
            self.assertIn("PANEL_REQUIRED", issue_codes)
            self.assertNotIn("PIANIST_AVAILABILITY_INCOMPLETE", issue_codes)

    def test_schedule_generation_and_result_routes_are_registered_and_work(self):
        route_paths = {route.path for route in jury_router.router.routes}
        self.assertTrue({
            "/api/jury/generate",
            "/api/jury/results/current",
            "/api/jury/results/history",
            "/api/jury/results/{result_uuid}",
        } <= route_paths)
        with Session(self.engine) as db:
            lesson, _ = self.add_lesson(
                db,
                student="Synthetic Schedule Route Student",
                pianist_required=False,
                jury_required=True,
            )
            lesson_uuid = db.get(module_models.AccompanistLessonIdentity, lesson.id).lesson_uuid
            finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panel = create_panel(db, JuryPanelFields(
                panel_name="Synthetic Route Panel",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            update_lesson_entry(db, lesson_uuid, panel_uuid=str(panel.panel_uuid))
            db.commit()
            readiness_report = readiness(db)
            self.assertTrue(readiness_report.ready, readiness_report.issues)
            expected_revision = readiness_report.jury_input_revision

        with Session(self.engine) as db:
            generated = jury_router.generate_schedule(
                JuryGenerateRequest(expected_jury_input_revision=expected_revision),
                db,
            )
            result_uuid = generated.result_uuid
            self.assertEqual(jury_router.get_current_schedule(db).result_uuid, result_uuid)
            history = jury_router.get_schedule_history(db)
            self.assertEqual(len(history), 1)
            self.assertEqual(jury_router.get_schedule_result(UUID(str(result_uuid)), db).result_uuid, result_uuid)

if __name__ == "__main__":
    unittest.main()