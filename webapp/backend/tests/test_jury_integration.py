from datetime import date
import asyncio
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from uuid import uuid4

from pydantic import ValidationError
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import models, module_models, schemas
from app.database import migrate_database
from app.schemas import (
    AvailabilityImportApplyRequest,
    AvailabilityImportMapping,
    AvailabilityImportPreviewRequest,
    ImportCommit,
)
from app.jury_schemas import (
    JuryAvailabilityIn,
    JuryPanelFields,
)
from app.database import get_db
from app.main import app
from app.services.accompanist_results import (
    ResultPublicationError,
    finalize_accompanist_result,
    get_current_accompanist_result,
)
from app.services import availability_importer, importer
from app.services.jury import (
    JuryDataError,
    assign_unassigned_panels_by_instrument,
    clear_availability,
    create_panel,
    get_availability,
    list_entries,
    readiness,
    save_availability,
    sync_roster_from_current_result,
    update_lesson_entry,
    update_lesson_jury_required,
    update_panel,
)
from app.services.jury_results import (
    JuryGenerationError,
    generate_schedule,
    get_current_schedule,
    get_schedule_by_uuid,
    list_schedule_history,
)
from app.services.jury_sync import sync_jury_with_accompanist
from app.services.module_lifecycle import (
    bump_accompanist_revision,
    ensure_pianist_identity,
    set_lesson_jury_required,
    synchronize_lesson_identity,
)
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

    def api_request(self, method: str, path: str, payload: dict | None = None):
        def override_get_db():
            with Session(self.engine) as db:
                yield db

        previous_override = app.dependency_overrides.get(get_db)
        app.dependency_overrides[get_db] = override_get_db
        request_body = json.dumps(payload or {}).encode("utf-8")
        messages = []
        request_sent = False

        async def receive():
            nonlocal request_sent
            if request_sent:
                return {"type": "http.disconnect"}
            request_sent = True
            return {"type": "http.request", "body": request_body, "more_body": False}

        async def send(message):
            messages.append(message)

        scope = {
            "type": "http",
            "asgi": {"version": "3.0", "spec_version": "2.3"},
            "http_version": "1.1",
            "method": method,
            "scheme": "http",
            "path": path,
            "raw_path": path.encode("ascii"),
            "query_string": b"",
            "root_path": "",
            "headers": [(b"host", b"testserver"), (b"content-type", b"application/json")],
            "client": ("testclient", 50000),
            "server": ("testserver", 80),
        }
        try:
            asyncio.run(app(scope, receive, send))
        finally:
            if previous_override is None:
                app.dependency_overrides.pop(get_db, None)
            else:
                app.dependency_overrides[get_db] = previous_override

        response_start = next(message for message in messages if message["type"] == "http.response.start")
        response_body = b"".join(
            message.get("body", b"")
            for message in messages
            if message["type"] == "http.response.body"
        )
        return response_start["status"], json.loads(response_body) if response_body else None

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

    def test_clear_availability_removes_only_selected_pianist_and_date(self):
        first_date = date(2027, 12, 14)
        second_date = date(2027, 12, 15)
        with Session(self.engine) as db:
            pianist, pianist_uuid = self.add_pianist(db, "Synthetic Accompanist")
            other_pianist, other_uuid = self.add_pianist(db, "Other Synthetic Pianist")
            lesson, _ = self.add_lesson(
                db,
                student="Availability Clear Student",
                pianist=pianist,
                jury_required=True,
            )
            finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            first_panel = create_panel(db, JuryPanelFields(
                panel_name="December 14 Panel",
                jury_date=first_date,
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            second_panel = create_panel(db, JuryPanelFields(
                panel_name="December 15 Panel",
                jury_date=second_date,
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            source_lesson_uuid = str(db.query(module_models.AccompanistLessonIdentity).filter_by(
                lesson_id=lesson.id
            ).one().lesson_uuid)
            update_lesson_entry(db, source_lesson_uuid, panel_uuid=str(first_panel.panel_uuid))
            save_availability(db, pianist_uuid, first_date, JuryAvailabilityIn(windows=[
                {"start_minute": 480, "end_minute": 540},
                {"start_minute": 600, "end_minute": 660},
            ]))
            save_availability(db, pianist_uuid, second_date, JuryAvailabilityIn(windows=[
                {"start_minute": 510, "end_minute": 570},
            ]))
            save_availability(db, other_uuid, first_date, JuryAvailabilityIn(windows=[
                {"start_minute": 540, "end_minute": 600},
            ]))
            db.commit()
            self.assertTrue(readiness(db).ready)
            configuration = db.get(module_models.JuryConfiguration, db.query(models.SchedulingSession).one().session_uuid)
            revision_before = configuration.input_revision

            deleted = clear_availability(db, pianist_uuid, first_date)
            db.commit()

            self.assertEqual(deleted, 2)
            self.assertIsNone(get_availability(db, pianist_uuid, first_date))
            self.assertEqual(len(get_availability(db, pianist_uuid, second_date).windows), 1)
            self.assertEqual(len(get_availability(db, other_uuid, first_date).windows), 1)
            self.assertIsNotNone(db.get(models.Pianist, pianist.id))
            self.assertEqual(db.query(module_models.JuryPanel).count(), 2)
            self.assertEqual(db.query(module_models.JuryPanelDate).count(), 2)
            self.assertEqual(configuration.input_revision, revision_before + 1)
            report = readiness(db)
            self.assertFalse(report.ready)
            missing = next(issue for issue in report.issues if issue.code == "PIANIST_AVAILABILITY_INCOMPLETE")
            self.assertEqual(
                missing.message,
                "Synthetic Accompanist has no Jury Availability Windows for December 14, 2027.",
            )

            self.assertEqual(clear_availability(db, pianist_uuid, first_date), 0)
            db.commit()
            self.assertEqual(configuration.input_revision, revision_before + 1)

            with self.assertRaises(JuryDataError) as wrong_date:
                clear_availability(db, pianist_uuid, date(2027, 12, 16))
            self.assertEqual(wrong_date.exception.code, "JURY_DATE_MISMATCH")
            self.assertIn("Jury Panel", str(wrong_date.exception))

            with self.assertRaises(JuryDataError) as raised:
                clear_availability(db, "00000000-0000-0000-0000-000000000000", first_date)
            self.assertEqual(raised.exception.code, "PIANIST_NOT_FOUND")
            self.assertIn("Pianist", str(raised.exception))

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
            self.assertNotIn("PIANIST_AVAILABILITY_INCOMPLETE", issue_codes)

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

    def test_readiness_blocks_required_source_lessons_missing_from_jury_entries(self):
        with Session(self.engine) as db:
            self.add_lesson(
                db,
                student="Unsynchronized Required Student",
                student_id="UNSYNCED-JURY",
                pianist_required=False,
                jury_required=True,
            )
            finalize_accompanist_result(db, expected_source_revision=0)

            result = readiness(db)

            self.assertFalse(result.ready)
            self.assertIn("JURY_ENTRY_MISSING", {issue.code for issue in result.issues})

    def test_lesson_import_replacement_removes_orphan_entries_and_stales_generated_results(self):
        with Session(self.engine) as db:
            first, _ = self.add_lesson(
                db,
                student="Old Lesson One",
                student_id="OLD-LESSON-1",
                pianist_required=False,
                jury_required=True,
            )
            second, _ = self.add_lesson(
                db,
                student="Old Lesson Two",
                student_id="OLD-LESSON-2",
                pianist_required=False,
                jury_required=True,
            )
            source = finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panel = create_panel(db, JuryPanelFields(
                panel_name="Old Lessons Panel",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            first_uuid = db.get(module_models.AccompanistLessonIdentity, first.id).lesson_uuid
            second_uuid = db.get(module_models.AccompanistLessonIdentity, second.id).lesson_uuid
            update_lesson_entry(db, first_uuid, panel_uuid=str(panel.panel_uuid))
            update_lesson_entry(db, second_uuid, panel_uuid=str(panel.panel_uuid))
            generated = generate_schedule(db)
            db.commit()

            csv_bytes = (
                "teacher,student,student_id,day,start_time,end_time,room,instrument,need_pianist,jury_required\n"
                "Synthetic Teacher,Replacement Student,NEW-LESSON-1,Monday,09:00 AM,09:50 AM,Studio 1,Voice,No,Yes\n"
            ).encode("utf-8")
            upload_token, columns, _ = importer.stage_upload("replacement-lessons.csv", csv_bytes)
            mapping = {field: field if field in columns else None for field in importer.TARGET_FIELDS}
            status, response = self.api_request(
                "POST",
                "/api/import/commit",
                {"upload_token": upload_token, "mapping": mapping},
            )

            self.assertEqual(status, 200)
            self.assertEqual(response["created"], 1)
            with Session(self.engine) as db:
                entries = db.query(module_models.JuryLessonEntry).all()
                stale = get_schedule_by_uuid(db, str(generated.result_uuid))
                self.assertEqual(entries, [])
                self.assertTrue(stale.stale)
                self.assertEqual(stale.source_result_uuid, source.result_uuid)
                self.assertEqual(db.query(module_models.JuryPanel).count(), 1)
                self.assertFalse(readiness(db).ready)

    def test_pianist_roster_replacement_removes_old_availability_and_stales_results(self):
        with Session(self.engine) as db:
            pianist, old_pianist_uuid = self.add_pianist(db, "Old Synthetic Pianist")
            lesson, _ = self.add_lesson(
                db,
                student="Retained Lesson Student",
                student_id="RETAINED-LESSON",
                pianist=pianist,
                jury_required=True,
            )
            source = finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            jury_date = date(2027, 5, 1)
            panel = create_panel(db, JuryPanelFields(
                panel_name="Retained Lesson Panel",
                jury_date=jury_date,
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            lesson_uuid = db.get(module_models.AccompanistLessonIdentity, lesson.id).lesson_uuid
            update_lesson_entry(db, lesson_uuid, panel_uuid=str(panel.panel_uuid))
            save_availability(db, old_pianist_uuid, jury_date, JuryAvailabilityIn(
                windows=[{"start_minute": 540, "end_minute": 600}],
            ))
            generated = generate_schedule(db)
            db.commit()

            csv_bytes = (
                "Pianist Name,Day,Start,End,Status\n"
                "Replacement Synthetic Pianist,Monday,08:00 AM,06:00 PM,Available\n"
            ).encode("utf-8")
            inspection = availability_importer.inspect_upload("replacement-pianists.csv", csv_bytes)
            mapping = AvailabilityImportMapping(
                layout="normalized",
                person_name_column="Pianist Name",
                day_column="Day",
                start_column="Start",
                end_column="End",
                status_column="Status",
            )
            preview = availability_importer.preview_import(
                AvailabilityImportPreviewRequest(
                    upload_token=inspection.upload_token,
                    sheet_name=inspection.selected_sheet,
                    mapping=mapping,
                ),
                db,
            )
            result = availability_importer.apply_import(
                AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
                db,
            )
            self.assertEqual(result.pianists_removed, 1)
            self.assertEqual(result.pianists_created, 1)

            self.assertIsNone(get_availability(db, old_pianist_uuid, jury_date))
            self.assertEqual(db.query(module_models.JuryPianistAvailabilityDeclaration).count(), 0)
            self.assertEqual(db.query(module_models.JuryPianistAvailableWindow).count(), 0)
            retained_entry = db.get(module_models.JuryLessonEntry, (db.query(models.SchedulingSession).one().session_uuid, lesson_uuid))
            self.assertEqual(retained_entry.panel_uuid, str(panel.panel_uuid))
            stale = get_schedule_by_uuid(db, str(generated.result_uuid))
            self.assertTrue(stale.stale)
            self.assertEqual(stale.source_result_uuid, source.result_uuid)
            self.assertFalse(readiness(db).ready)
            self.assertEqual(list_entries(db), [])

    def test_module_open_sync_removes_legacy_orphan_lesson_and_pianist_references(self):
        with Session(self.engine) as db:
            pianist, pianist_uuid = self.add_pianist(db, "Current Synthetic Pianist")
            lesson, student_uuid = self.add_lesson(
                db,
                student="Current Source Lesson",
                student_id="CURRENT-SOURCE",
                pianist_required=False,
                jury_required=True,
            )
            finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panel = create_panel(db, JuryPanelFields(
                panel_name="Current Source Panel",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            lesson_uuid = db.get(module_models.AccompanistLessonIdentity, lesson.id).lesson_uuid
            update_lesson_entry(db, lesson_uuid, panel_uuid=str(panel.panel_uuid))
            generated = generate_schedule(db)

            obsolete_pianist, obsolete_uuid = self.add_pianist(db, "Obsolete Synthetic Pianist")
            db.add(module_models.JuryPianistAvailabilityDeclaration(
                session_uuid=db.query(models.SchedulingSession).one().session_uuid,
                pianist_person_uuid=obsolete_uuid,
                jury_date=date(2027, 5, 1),
                is_complete=True,
            ))
            db.flush()
            db.add(module_models.JuryPianistAvailableWindow(
                session_uuid=db.query(models.SchedulingSession).one().session_uuid,
                pianist_person_uuid=obsolete_uuid,
                jury_date=date(2027, 5, 1),
                start_minute=540,
                end_minute=600,
            ))
            orphan_lesson_uuid = str(uuid4())
            db.add(module_models.JuryLessonEntry(
                session_uuid=db.query(models.SchedulingSession).one().session_uuid,
                source_lesson_uuid=orphan_lesson_uuid,
                student_person_uuid=student_uuid,
                panel_uuid=str(panel.panel_uuid),
            ))
            db.flush()
            db.delete(db.get(module_models.AccompanistPianistIdentity, obsolete_pianist.id))
            db.delete(obsolete_pianist)
            bump_accompanist_revision(db)
            db.commit()

        status, summary = self.api_request("POST", "/api/jury/synchronize")

        self.assertEqual(status, 200)
        self.assertEqual(summary["lesson_entries_removed"], 1)
        self.assertEqual(summary["availability_records_removed"], 1)
        self.assertEqual(summary["availability_windows_removed"], 1)
        self.assertEqual(summary["stale_results_detected"], 1)
        with Session(self.engine) as db:
            self.assertIsNone(db.get(module_models.JuryLessonEntry, (
                db.query(models.SchedulingSession).one().session_uuid,
                orphan_lesson_uuid,
            )))
            survivor = db.get(module_models.JuryLessonEntry, (
                db.query(models.SchedulingSession).one().session_uuid,
                lesson_uuid,
            ))
            self.assertEqual(survivor.panel_uuid, str(panel.panel_uuid))
            self.assertEqual(db.query(module_models.JuryPianistAvailabilityDeclaration).count(), 0)
            self.assertEqual(db.query(module_models.JuryPianistAvailableWindow).count(), 0)
            self.assertTrue(get_schedule_by_uuid(db, str(generated.result_uuid)).stale)
            self.assertFalse(readiness(db).ready)

    def test_lesson_activation_panel_matching_is_exact_case_insensitive_and_non_destructive(self):
        with Session(self.engine) as db:
            lessons = {}
            lesson_specs = (
                ("Exact Case Match", "MATCH-1", "Violin", True),
                ("Non-required Exact Match", "MATCH-2", "VOICE", False),
                ("Partial Match", "MATCH-3", "Cello", True),
                ("Existing Assignment", "MATCH-4", "Flute", True),
                ("Ambiguous Match", "MATCH-5", "oBoE", True),
                ("Whitespace Mismatch", "MATCH-6", "Cello ", True),
            )
            for student, student_id, instrument, jury_required in lesson_specs:
                lesson, _ = self.add_lesson(
                    db,
                    student=student,
                    student_id=student_id,
                    instrument=instrument,
                    pianist_required=False,
                    jury_required=jury_required,
                )
                lessons[student] = db.get(module_models.AccompanistLessonIdentity, lesson.id).lesson_uuid

            finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panels = {}
            for panel_name in ("vIoLiN", "voice", "Cell", "Flute", "Keep Existing", "Oboe", "OBoe"):
                panels[panel_name] = create_panel(db, JuryPanelFields(
                    panel_name=panel_name,
                    jury_date=date(2027, 5, 1),
                    earliest_start_minute=540,
                    jury_length_minutes=30,
                ))
            update_lesson_entry(
                db,
                lessons["Existing Assignment"],
                panel_uuid=str(panels["Keep Existing"].panel_uuid),
            )
            config = db.query(module_models.JuryConfiguration).one()
            revision_before = config.input_revision

            refreshed, assigned_count = assign_unassigned_panels_by_instrument(db)
            db.commit()

            panels_by_student = {entry.student_display_name: entry.panel_uuid for entry in refreshed}
            self.assertEqual(assigned_count, 2)
            self.assertEqual(panels_by_student["Exact Case Match"], panels["vIoLiN"].panel_uuid)
            self.assertEqual(panels_by_student["Non-required Exact Match"], panels["voice"].panel_uuid)
            self.assertIsNone(panels_by_student["Partial Match"])
            self.assertEqual(panels_by_student["Existing Assignment"], panels["Keep Existing"].panel_uuid)
            self.assertIsNone(panels_by_student["Ambiguous Match"])
            self.assertIsNone(panels_by_student["Whitespace Mismatch"])
            self.assertEqual(config.input_revision, revision_before + 1)

            _, repeated_count = assign_unassigned_panels_by_instrument(db)
            db.commit()
            self.assertEqual(repeated_count, 0)
            self.assertEqual(config.input_revision, revision_before + 1)

    def test_panel_match_endpoint_returns_automatic_assignments(self):
        with Session(self.engine) as db:
            lesson, _ = self.add_lesson(
                db,
                student="Synthetic Endpoint Match",
                instrument="Cello",
                pianist_required=False,
                jury_required=False,
            )
            finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panel = create_panel(db, JuryPanelFields(
                panel_name="cELLo",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            lesson_uuid = db.get(module_models.AccompanistLessonIdentity, lesson.id).lesson_uuid
            db.commit()

        status, response = self.api_request("POST", "/api/jury/entries/assign-panels-by-instrument")

        self.assertEqual(status, 200)
        self.assertEqual(response["assigned_count"], 1)
        matched = next(entry for entry in response["entries"] if entry["source_lesson_uuid"] == lesson_uuid)
        self.assertEqual(matched["panel_uuid"], str(panel.panel_uuid))

    def test_generate_schedule_persists_typed_result_and_dependency(self):
        with Session(self.engine) as db:
            pianist, pianist_uuid = self.add_pianist(db, "Synthetic Scheduled Pianist")
            lesson, _ = self.add_lesson(
                db,
                student="Synthetic Scheduled Student",
                pianist=pianist,
                jury_required=True,
            )
            source = finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panel = create_panel(db, JuryPanelFields(
                panel_name="Synthetic Schedule Panel",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=60,
            ))
            lesson_uuid = db.get(module_models.AccompanistLessonIdentity, lesson.id).lesson_uuid
            update_lesson_entry(db, lesson_uuid, panel_uuid=str(panel.panel_uuid))
            save_availability(db, pianist_uuid, date(2027, 5, 1), JuryAvailabilityIn(
                windows=[{"start_minute": 540, "end_minute": 660}],
            ))
            db.flush()

            generated = generate_schedule(db)
            db.commit()

            stored = db.get(module_models.ModuleResult, str(generated.result_uuid))
            dependency = db.get(
                module_models.ModuleResultDependency,
                (str(generated.result_uuid), "accompanist.assignment-result"),
            )
            loaded = get_schedule_by_uuid(db, str(generated.result_uuid))
            current = get_current_schedule(db)
            self.assertEqual(stored.state, "draft")
            self.assertEqual(stored.payload_sha256, __import__("hashlib").sha256(stored.payload_json.encode()).hexdigest())
            self.assertEqual(dependency.source_result_uuid, str(source.result_uuid))
            self.assertEqual(dependency.source_revision, source.source_revision)
            self.assertEqual(loaded, current)
            self.assertEqual(loaded.source_result_uuid, source.result_uuid)
            self.assertEqual(loaded.source_revision, source.source_revision)
            self.assertEqual(loaded.panel_timelines[0].events[0].kind, "jury")
            self.assertEqual(str(loaded.panel_timelines[0].events[0].source_lesson_uuid), lesson_uuid)
            self.assertEqual(str(loaded.panel_timelines[0].events[0].pianist_person_uuid), pianist_uuid)

    def test_repeated_generation_creates_immutable_history_and_supersedes_previous_draft(self):
        with Session(self.engine) as db:
            self.add_lesson(
                db,
                student="Synthetic No-Pianist Student",
                pianist_required=False,
                jury_required=True,
            )
            finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panel = create_panel(db, JuryPanelFields(
                panel_name="Repeat Panel",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            lesson_uuid = db.query(module_models.JuryLessonEntry).one().source_lesson_uuid
            update_lesson_entry(db, lesson_uuid, panel_uuid=str(panel.panel_uuid))
            first = generate_schedule(db)
            db.commit()
            first_row = db.get(module_models.ModuleResult, str(first.result_uuid))
            first_payload = first_row.payload_json

            second = generate_schedule(db)
            db.commit()

            self.assertNotEqual(first.result_uuid, second.result_uuid)
            self.assertEqual(first_row.payload_json, first_payload)
            self.assertEqual(first_row.state, "superseded")
            self.assertEqual(db.query(module_models.ModuleResult).filter_by(module_id="juries").count(), 2)
            self.assertEqual(len(list_schedule_history(db)), 2)
            self.assertEqual(get_current_schedule(db).result_uuid, second.result_uuid)

    def test_accompanist_revision_marks_existing_jury_result_stale_without_relinking(self):
        with Session(self.engine) as db:
            self.add_lesson(
                db,
                student="Synthetic Stale Student",
                pianist_required=False,
                jury_required=True,
            )
            source = finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panel = create_panel(db, JuryPanelFields(
                panel_name="Stale Panel",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            lesson_uuid = db.query(module_models.JuryLessonEntry).one().source_lesson_uuid
            update_lesson_entry(db, lesson_uuid, panel_uuid=str(panel.panel_uuid))
            generated = generate_schedule(db)
            db.commit()

            bump_accompanist_revision(db)
            db.commit()

            stale_current = get_current_schedule(db)
            historical = get_schedule_by_uuid(db, str(generated.result_uuid))
            self.assertTrue(stale_current.stale)
            self.assertIn("accompanist_source_revision_changed", stale_current.stale_reasons)
            self.assertTrue(historical.stale)
            self.assertEqual(historical.source_result_uuid, source.result_uuid)
            self.assertEqual(len(list_schedule_history(db)), 1)
            self.assertEqual(db.get(module_models.ModuleResult, str(generated.result_uuid)).state, "draft")

    def test_new_finalized_accompanist_result_at_same_revision_marks_jury_result_stale(self):
        with Session(self.engine) as db:
            lesson, _ = self.add_lesson(
                db,
                student="Synthetic Refinalized Student",
                pianist_required=False,
                jury_required=True,
            )
            first_source = finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panel = create_panel(db, JuryPanelFields(
                panel_name="Refinalized Source Panel",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            lesson_uuid = db.get(module_models.AccompanistLessonIdentity, lesson.id).lesson_uuid
            update_lesson_entry(db, lesson_uuid, panel_uuid=str(panel.panel_uuid))
            generated = generate_schedule(db)
            db.commit()

            replacement_source = finalize_accompanist_result(db, expected_source_revision=0)
            db.commit()

            stale = get_schedule_by_uuid(db, str(generated.result_uuid))
            self.assertEqual(first_source.source_revision, replacement_source.source_revision)
            self.assertNotEqual(first_source.result_uuid, replacement_source.result_uuid)
            self.assertTrue(stale.stale)
            self.assertIn("accompanist_result_changed", stale.stale_reasons)
            self.assertNotIn("accompanist_source_revision_changed", stale.stale_reasons)
            self.assertEqual(stale.source_result_uuid, first_source.result_uuid)

    def test_jury_input_revision_change_marks_current_schedule_stale(self):
        with Session(self.engine) as db:
            lesson, _ = self.add_lesson(
                db,
                student="Synthetic Jury Revision Student",
                pianist_required=False,
                jury_required=True,
            )
            finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panel = create_panel(db, JuryPanelFields(
                panel_name="Jury Revision Panel",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            lesson_uuid = db.get(module_models.AccompanistLessonIdentity, lesson.id).lesson_uuid
            update_lesson_entry(db, lesson_uuid, panel_uuid=str(panel.panel_uuid))
            generated = generate_schedule(db)
            db.commit()

            update_panel(db, str(panel.panel_uuid), JuryPanelFields(
                panel_name=panel.panel_name,
                jury_date=panel.jury_date,
                earliest_start_minute=panel.earliest_start_minute,
                preferred_start_minute=600,
                jury_length_minutes=panel.jury_length_minutes,
            ))
            db.commit()

            stale = get_current_schedule(db)
            self.assertEqual(stale.result_uuid, generated.result_uuid)
            self.assertTrue(stale.stale)
            self.assertEqual(stale.stale_reasons, ["jury_inputs_changed"])
            self.assertEqual(db.get(module_models.ModuleRevision, (str(generated.session_uuid), "juries")).source_revision,
                             db.get(module_models.JuryConfiguration, str(generated.session_uuid)).input_revision)

    def test_readiness_blocker_prevents_optimizer_invocation(self):
        with Session(self.engine) as db:
            self.add_lesson(
                db,
                student="Synthetic Blocked Student",
                pianist_required=False,
                jury_required=True,
            )
            finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)

            with patch("app.services.jury_results.DeterministicJuryOptimizer") as optimizer_type:
                with self.assertRaises(JuryGenerationError) as raised:
                    generate_schedule(db)

            self.assertEqual(raised.exception.code, "JURY_NOT_READY")
            self.assertIn("PANEL_REQUIRED", {issue.code for issue in raised.exception.readiness.issues})
            optimizer_type.return_value.optimize.assert_not_called()

    def test_schedule_payload_round_trips_scheduled_and_unscheduled_dispositions(self):
        with Session(self.engine) as db:
            pianist, pianist_uuid = self.add_pianist(db, "Synthetic Contended Pianist")
            first, _ = self.add_lesson(
                db,
                student="Synthetic First Contended Student",
                student_id="CONTENDED-1",
                pianist=pianist,
                jury_required=True,
            )
            second, _ = self.add_lesson(
                db,
                student="Synthetic Second Contended Student",
                student_id="CONTENDED-2",
                pianist=pianist,
                jury_required=True,
            )
            source = finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            first_panel = create_panel(db, JuryPanelFields(
                panel_name="Contended Panel One",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=60,
            ))
            second_panel = create_panel(db, JuryPanelFields(
                panel_name="Contended Panel Two",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=60,
            ))
            first_uuid = db.get(module_models.AccompanistLessonIdentity, first.id).lesson_uuid
            second_uuid = db.get(module_models.AccompanistLessonIdentity, second.id).lesson_uuid
            update_lesson_entry(db, first_uuid, panel_uuid=str(first_panel.panel_uuid))
            update_lesson_entry(db, second_uuid, panel_uuid=str(second_panel.panel_uuid))
            save_availability(db, pianist_uuid, date(2027, 5, 1), JuryAvailabilityIn(
                windows=[{"start_minute": 540, "end_minute": 600}],
            ))

            generated = generate_schedule(db)
            db.commit()
            loaded = get_schedule_by_uuid(db, str(generated.result_uuid))

            self.assertEqual(len(loaded.panel_timelines), 2)
            self.assertEqual(sum(event.kind == "jury" for panel_view in loaded.panel_timelines for event in panel_view.events), 1)
            self.assertEqual(len(loaded.unscheduled_lessons), 1)
            self.assertEqual(loaded.unscheduled_lessons[0].reason_code, "fixed_pianist_conflict")
            self.assertEqual(loaded.source_result_uuid, source.result_uuid)

    def test_jury_api_handlers_generate_current_history_and_lookup(self):
        with Session(self.engine) as db:
            self.add_lesson(
                db,
                student="Synthetic API Student",
                pianist_required=False,
                jury_required=True,
            )
            finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            panel = create_panel(db, JuryPanelFields(
                panel_name="Synthetic API Panel",
                jury_date=date(2027, 5, 1),
                earliest_start_minute=540,
                jury_length_minutes=30,
            ))
            lesson_uuid = db.query(module_models.JuryLessonEntry).one().source_lesson_uuid
            update_lesson_entry(db, lesson_uuid, panel_uuid=str(panel.panel_uuid))
            db.commit()

        status, generated = self.api_request("POST", "/api/jury/generate", {})
        current_status, current = self.api_request("GET", "/api/jury/results/current")
        history_status, history = self.api_request("GET", "/api/jury/results/history")
        result_status, historical = self.api_request("GET", f"/api/jury/results/{generated['result_uuid']}")

        self.assertEqual(status, 201)
        self.assertEqual(current_status, 200)
        self.assertEqual(history_status, 200)
        self.assertEqual(result_status, 200)
        self.assertEqual(current["result_uuid"], generated["result_uuid"])
        self.assertEqual(historical["result_uuid"], generated["result_uuid"])
        self.assertEqual(len(history), 1)
        self.assertEqual(history[0]["scheduled_count"], 1)

    def test_jury_generate_api_returns_structured_readiness_blockers(self):
        with Session(self.engine) as db:
            self.add_lesson(
                db,
                student="Synthetic API Blocked Student",
                pianist_required=False,
                jury_required=True,
            )
            finalize_accompanist_result(db, expected_source_revision=0)
            sync_roster_from_current_result(db)
            db.commit()

        status, response = self.api_request("POST", "/api/jury/generate", {})

        self.assertEqual(status, 409)
        self.assertEqual(response["detail"]["code"], "JURY_NOT_READY")
        issue_codes = {issue["code"] for issue in response["detail"]["readiness"]["issues"]}
        self.assertIn("PANEL_REQUIRED", issue_codes)
        with Session(self.engine) as db:
            self.assertEqual(db.query(module_models.ModuleResult).filter_by(module_id="juries").count(), 0)

if __name__ == "__main__":
    unittest.main()