import unittest
from io import BytesIO
from datetime import date, time
import tempfile
from uuid import uuid4
from unittest.mock import patch

from openpyxl import Workbook
from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from sqlalchemy.pool import StaticPool

from app import models, module_models, schemas
from app.database import Base, migrate_database
from app.routers import pianists
from app.services import availability_importer
from app.services.accompanist_availability import windows_to_accompanist_slots
from app.services.availability import (
    AvailabilityWindow,
    normalize_day,
    normalize_status,
    normalize_windows,
    parse_time_minutes,
)
from app.services.module_lifecycle import active_session_uuid, ensure_pianist_identity, module_revision, synchronize_lesson_identity


class AvailabilityValueTests(unittest.TestCase):
    def test_day_names_and_unambiguous_abbreviations_normalize(self):
        self.assertEqual(normalize_day("mOn."), "Monday")
        self.assertEqual(normalize_day("Thurs"), "Thursday")
        self.assertIsNone(normalize_day("T"))
        self.assertIsNone(normalize_day("Funday"))

    def test_time_parser_accepts_excel_times_and_common_text(self):
        self.assertEqual(parse_time_minutes(time(8, 0)), 480)
        self.assertEqual(parse_time_minutes(0.5), 720)
        self.assertEqual(parse_time_minutes("8:00 AM"), 480)
        self.assertEqual(parse_time_minutes("08:00"), 480)
        self.assertEqual(parse_time_minutes("1:30 PM"), 810)
        self.assertEqual(parse_time_minutes("24:00", allow_end_of_day=True), 1440)
        self.assertIsNone(parse_time_minutes("13:00 PM"))
        self.assertIsNone(parse_time_minutes("25:00"))
        self.assertIsNone(parse_time_minutes("8"))

    def test_status_vocabulary_is_case_insensitive_without_scoring_semantics(self):
        self.assertEqual(normalize_status(" tentative "), "Tentative")
        self.assertEqual(normalize_status("not sure"), None)

    def test_same_status_overlap_and_adjacency_merge(self):
        windows, issues = normalize_windows([
            AvailabilityWindow("Monday", 8 * 60, 11 * 60, "Available"),
            AvailabilityWindow("Monday", 10 * 60 + 30, 12 * 60, "Available"),
            AvailabilityWindow("Monday", 12 * 60, 13 * 60, "Available"),
        ])

        self.assertEqual([(item.start_minute, item.end_minute) for item in windows], [(480, 780)])
        self.assertEqual(issues, [])

    def test_conflicting_overlap_is_reported_and_not_merged(self):
        windows, issues = normalize_windows([
            AvailabilityWindow("Monday", 9 * 60, 12 * 60, "Available"),
            AvailabilityWindow("Monday", 10 * 60, 11 * 60, "Tentative", source_row=4),
        ])

        self.assertEqual(len(windows), 2)
        self.assertEqual([(issue.severity, issue.code, issue.row_number) for issue in issues], [
            ("error", "CONFLICTING_OVERLAP", 4),
        ])

    def test_duplicate_windows_are_deduplicated_with_warning(self):
        windows, issues = normalize_windows([
            AvailabilityWindow("Friday", 540, 600, "Unavailable", source_row=2),
            AvailabilityWindow("Friday", 540, 600, "Unavailable", source_row=3),
        ])

        self.assertEqual(len(windows), 1)
        self.assertEqual([(issue.severity, issue.code) for issue in issues], [("warning", "DUPLICATE_WINDOW")])

    def test_accompanist_adapter_preserves_existing_half_hour_fit_and_rejects_rounding(self):
        windows = [AvailabilityWindow("Monday", 8 * 60, 10 * 60, "Available")]
        slots, issues = windows_to_accompanist_slots(windows)
        availability = {windows[0].day: {slot.slot_start_minute: slot.status for slot in slots}}

        from app.services.scheduling import FIT_FULL, get_fit

        self.assertEqual(issues, [])
        self.assertEqual(get_fit("Monday", 8 * 60, 9 * 60, availability)[0], FIT_FULL)

        rounded_slots, invalid = windows_to_accompanist_slots([
            AvailabilityWindow("Monday", 8 * 60 + 15, 10 * 60, "Available", source_row=7),
        ])
        self.assertEqual(rounded_slots, [])
        self.assertEqual([(issue.code, issue.row_number) for issue in invalid], [
            ("UNSUPPORTED_ACCOMPANIST_BOUNDARY", 7),
        ])
        unavailable_slots, invalid = windows_to_accompanist_slots([
            AvailabilityWindow("Monday", 8 * 60, 10 * 60, "Unavailable"),
        ])
        self.assertEqual(unavailable_slots, [])
        self.assertEqual(invalid, [])


class AvailabilityImportServiceTests(unittest.TestCase):
    def setUp(self):
        availability_importer._STAGED_UPLOADS.clear()
        availability_importer._STAGED_PREVIEWS.clear()
        self.engine = create_engine(
            "sqlite://",
            connect_args={"check_same_thread": False},
            poolclass=StaticPool,
        )
        migrate_database(self.engine)
        self.db = Session(self.engine)
        self.db.add_all([
            models.Pianist(name="Ari Example"),
            models.Pianist(name="Bea Sample"),
        ])
        self.db.commit()

    def tearDown(self):
        availability_importer._STAGED_UPLOADS.clear()
        availability_importer._STAGED_PREVIEWS.clear()
        self.db.close()
        self.engine.dispose()

    def pianist_id(self, name):
        return self.db.query(models.Pianist).filter_by(name=name).one().id

    def csv_upload(self, content):
        return availability_importer.inspect_upload("synthetic.csv", content.encode())

    def normalized_mapping(self, inspection, person="Pianist Name", email=None):
        return schemas.AvailabilityImportMapping(
            person_name_column=person,
            email_column=email,
            days=inspection.suggested_days,
        )

    def preview_csv(self, content, *, person="Pianist Name", email=None):
        inspection = self.csv_upload(content)
        preview = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(
                upload_token=inspection.upload_token,
                mapping=self.normalized_mapping(inspection, person, email),
            ),
            self.db,
        )
        return inspection, preview

    def add_jury_availability(self, person_uuid, jury_date):
        session_uuid = active_session_uuid(self.db)
        self.db.add(module_models.JuryPianistAvailabilityDeclaration(
            session_uuid=session_uuid,
            pianist_person_uuid=person_uuid,
            jury_date=jury_date,
            is_complete=True,
        ))
        self.db.add(module_models.JuryPianistAvailableWindow(
            session_uuid=session_uuid,
            pianist_person_uuid=person_uuid,
            jury_date=jury_date,
            start_minute=480,
            end_minute=540,
        ))

    def test_import_groups_repeated_normalized_names_and_replaces_old_roster(self):
        _, preview = self.preview_csv(
            "Pianist Name,Monday\n"
            '"  Ari   Example  ",08:00-09:30\n'
            "Ari Example,T 14:00-15:00\n"
        )

        self.assertTrue(preview.can_apply)
        self.assertEqual(preview.rows_processed, 2)
        self.assertEqual(preview.existing_pianist_count, 2)
        self.assertEqual(preview.incoming_pianist_count, 1)
        self.assertEqual(preview.valid_window_count, 2)
        self.assertIn("FULL_ROSTER_REPLACEMENT", {warning.code for warning in preview.warnings})
        self.assertEqual([person.pianist_name for person in preview.pianists], ["Ari Example"])
        self.assertEqual([window.status for window in preview.windows], ["Available", "Tentative"])
        self.assertEqual(self.db.query(models.AvailabilitySlot).count(), 0)
        self.assertNotIn("pianist_code", preview.model_dump())
        self.assertNotIn("pianist_id", preview.model_dump())

        result = availability_importer.apply_import(
            schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
            self.db,
        )

        self.assertEqual((result.pianists_removed, result.pianists_created), (2, 1))
        self.assertEqual(result.slots_created, 5)
        self.assertEqual(self.db.query(models.Pianist).one().name, "Ari Example")
        self.assertEqual(self.db.query(models.AvailabilitySlot).count(), 5)

    def test_import_without_identifier_column_creates_fresh_people_with_defaults(self):
        self.db.query(models.Pianist).delete()
        self.db.commit()
        rows = [f"Person {index},,08:00-09:00" for index in range(1, 7)]
        inspection, preview = self.preview_csv(
            "Pianist Name,Email,Monday\n" + "\n".join(rows) + "\n",
        )

        self.assertTrue(preview.can_apply)
        self.assertEqual(preview.incoming_pianist_count, 6)
        self.assertNotIn("person_id_column", inspection.suggested_columns)
        self.assertTrue(all(item.action == "new" and item.email == "" for item in preview.pianists))
        self.assertTrue(all(item.max_hours_per_week == 40 for item in preview.pianists))

        result = availability_importer.apply_import(
            schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
            self.db,
        )

        pianists = self.db.query(models.Pianist).order_by(models.Pianist.name).all()
        mappings = self.db.query(module_models.AccompanistPianistIdentity).all()
        self.assertEqual((result.pianists_created, result.pianists_removed), (6, 0))
        self.assertEqual(len(pianists), 6)
        self.assertEqual(len(mappings), 6)
        self.assertEqual(len({mapping.person_uuid for mapping in mappings}), 6)
        self.assertTrue(all(pianist.email == "" and pianist.max_hours_per_week == 40 for pianist in pianists))
        self.assertTrue(all(pianist.pianist_code is None for pianist in pianists))
        self.assertEqual(self.db.query(models.AvailabilitySlot).count(), 12)
        self.assertTrue(all(pianist.availability_complete for pianist in pianists))

    def test_max_hours_column_is_imported_defaulted_and_validated(self):
        inspection = self.csv_upload(
            "Pianist Name,Max Hours Per Week,Monday\nAri Example,12.5,08:00-09:00\nBea Sample,,09:00-10:00\n"
        )
        self.assertEqual(inspection.suggested_columns["max_hours_column"], "Max Hours Per Week")
        mapping = schemas.AvailabilityImportMapping(
            person_name_column="Pianist Name", max_hours_column="Max Hours Per Week", days={"Monday": "Monday"},
        )
        preview = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(upload_token=inspection.upload_token, mapping=mapping), self.db,
        )
        self.assertEqual([p.max_hours_per_week for p in preview.pianists], [12.5, 40])
        availability_importer.apply_import(
            schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True), self.db,
        )
        hours = {p.name: p.max_hours_per_week for p in self.db.query(models.Pianist).all()}
        self.assertEqual(hours, {"Ari Example": 12.5, "Bea Sample": 40})

        bad = self.csv_upload("Pianist Name,Max Hours Per Week,Monday\nAri Example,lots,08:00-09:00\n")
        invalid = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(upload_token=bad.upload_token, mapping=mapping), self.db,
        )
        self.assertIn("INVALID_MAX_HOURS", {issue.code for issue in invalid.errors})
        self.assertFalse(invalid.can_apply)

    def test_user_models_have_no_pianist_identifier(self):
        self.assertNotIn("pianist_code", schemas.PianistCreate.model_fields)
        self.assertNotIn("pianist_code", schemas.PianistUpdate.model_fields)
        self.assertNotIn("pianist_code", schemas.PianistOut.model_fields)
        self.assertNotIn("person_id_column", schemas.AvailabilityImportMapping.model_fields)
        with self.assertRaises(Exception):
            schemas.AvailabilityImportMapping(
                person_name_column="Pianist Name",
                person_id_column="Pianist ID",
            )

    def test_imported_defaults_remain_editable_through_human_facing_api(self):
        _, preview = self.preview_csv(
            "Pianist Name,Monday\nEditable Import,08:00-09:00\n"
        )
        availability_importer.apply_import(
            schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
            self.db,
        )
        imported = self.db.query(models.Pianist).filter_by(name="Editable Import").one()
        self.assertEqual((imported.email, imported.max_hours_per_week), ("", 40))
        updated = pianists.update_pianist(imported.id, schemas.PianistUpdate(
            name="Edited Import",
            email="edited@example.invalid",
            max_hours_per_week=12,
        ), self.db)
        self.assertEqual(
            (updated.name, updated.email, updated.max_hours_per_week),
            ("Edited Import", "edited@example.invalid", 12),
        )

    def test_manual_pianist_creation_generates_internal_identity_without_id_input(self):
        created = pianists.create_pianist(schemas.PianistCreate(name="Manual Pianist"), self.db)
        self.assertIsNone(created.pianist_code)
        self.assertNotIn("pianist_code", schemas.PianistOut.model_validate(created).model_dump())
        mapping = self.db.get(module_models.AccompanistPianistIdentity, created.id)
        self.assertIsNotNone(mapping)
        self.assertIsNotNone(self.db.get(module_models.PersonIdentity, mapping.person_uuid))
        with self.assertRaises(Exception):
            schemas.PianistCreate(name="Invalid Input", pianist_code="P001")

    def test_invalid_window_blocks_new_pianist_creation_atomically(self):
        ari_id = self.pianist_id("Ari Example")
        ari_uuid = ensure_pianist_identity(self.db, self.db.get(models.Pianist, ari_id))
        self.add_jury_availability(ari_uuid, date(2027, 12, 14))
        self.db.commit()
        _, preview = self.preview_csv(
            "Pianist Name,Monday,Tuesday\n"
            "Would Be New,08:00-09:00,10:00-\n"
        )
        self.assertFalse(preview.can_apply)
        with self.assertRaises(availability_importer.AvailabilityImportError):
            availability_importer.apply_import(
                schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
                self.db,
            )
        self.assertEqual(self.db.query(models.Pianist).filter_by(name="Would Be New").count(), 0)
        self.assertEqual(self.db.query(module_models.AccompanistPianistIdentity).count(), 1)
        self.assertEqual(self.db.query(models.AvailabilitySlot).count(), 0)
        self.assertEqual(self.db.query(module_models.JuryPianistAvailabilityDeclaration).count(), 1)
        self.assertEqual(self.db.query(module_models.JuryPianistAvailableWindow).count(), 1)

    def test_cancelled_replacement_preserves_jury_availability(self):
        ari_id = self.pianist_id("Ari Example")
        ari_uuid = ensure_pianist_identity(self.db, self.db.get(models.Pianist, ari_id))
        self.add_jury_availability(ari_uuid, date(2027, 12, 14))
        self.db.commit()
        _, preview = self.preview_csv(
            "Pianist Name,Monday\nNew Pianist,08:00-09:00\n"
        )

        with self.assertRaises(availability_importer.AvailabilityImportError) as raised:
            availability_importer.apply_import(
                schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=False),
                self.db,
            )
        self.assertEqual(raised.exception.code, "CONFIRMATION_REQUIRED")
        self.assertEqual(self.db.query(models.Pianist).count(), 2)
        self.assertEqual(self.db.query(module_models.JuryPianistAvailabilityDeclaration).count(), 1)
        self.assertEqual(self.db.query(module_models.JuryPianistAvailableWindow).count(), 1)

    def test_failed_replacement_rolls_back_jury_cleanup_and_accompanist_changes(self):
        ari_id = self.pianist_id("Ari Example")
        ari_uuid = ensure_pianist_identity(self.db, self.db.get(models.Pianist, ari_id))
        self.add_jury_availability(ari_uuid, date(2027, 12, 14))
        jury_configuration = self.db.get(module_models.JuryConfiguration, active_session_uuid(self.db))
        initial_jury_revision = jury_configuration.input_revision
        lesson = models.Lesson(
            student="Rollback Student",
            day="Monday",
            start_minute=480,
            end_minute=510,
            assigned_pianist_id=ari_id,
        )
        self.db.add_all([
            lesson,
            models.AvailabilitySlot(
                pianist_id=ari_id,
                day="Monday",
                slot_start_minute=480,
                status="Available",
            ),
        ])
        self.db.commit()
        _, preview = self.preview_csv(
            "Pianist Name,Monday\nNew Pianist,08:00-09:00\n"
        )

        with patch.object(availability_importer, "bump_accompanist_revision", side_effect=RuntimeError("forced failure")):
            with self.assertRaises(availability_importer.AvailabilityImportError) as raised:
                availability_importer.apply_import(
                    schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
                    self.db,
                )

        self.assertEqual(raised.exception.code, "APPLY_FAILED")
        self.db.expire_all()
        self.assertEqual(self.db.query(models.Pianist).count(), 2)
        self.assertEqual(self.db.query(models.Pianist).filter_by(name="New Pianist").count(), 0)
        self.assertEqual(self.db.get(models.Lesson, lesson.id).assigned_pianist_id, ari_id)
        self.assertEqual(self.db.query(models.AvailabilitySlot).filter_by(pianist_id=ari_id).count(), 1)
        self.assertEqual(self.db.query(module_models.JuryPianistAvailabilityDeclaration).count(), 1)
        self.assertEqual(self.db.query(module_models.JuryPianistAvailableWindow).count(), 1)
        self.assertEqual(
            self.db.get(module_models.JuryConfiguration, active_session_uuid(self.db)).input_revision,
            initial_jury_revision,
        )

    def test_reimport_same_name_replaces_identity_instead_of_matching_old_roster(self):
        csv = "Pianist Name,Monday\n"
        first_inspection, first_preview = self.preview_csv(
            csv + "Repeat Import,08:00-09:00\n",
        )
        availability_importer.apply_import(
            schemas.AvailabilityImportApplyRequest(preview_token=first_preview.preview_token, confirmed=True),
            self.db,
        )
        first_pianist = self.db.query(models.Pianist).filter_by(name="Repeat Import").one()
        first_identity = self.db.get(module_models.AccompanistPianistIdentity, first_pianist.id).person_uuid
        second_preview = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(
                upload_token=first_inspection.upload_token,
                mapping=self.normalized_mapping(first_inspection),
            ),
            self.db,
        )
        self.assertEqual((second_preview.existing_pianist_count, second_preview.incoming_pianist_count), (1, 1))
        availability_importer.apply_import(
            schemas.AvailabilityImportApplyRequest(preview_token=second_preview.preview_token, confirmed=True),
            self.db,
        )
        second_pianist = self.db.query(models.Pianist).filter_by(name="Repeat Import").one()
        second_identity = self.db.get(module_models.AccompanistPianistIdentity, second_pianist.id).person_uuid
        self.assertNotEqual(first_identity, second_identity)
        self.assertEqual(self.db.query(models.Pianist).count(), 1)

    def test_xlsx_weekday_columns_multiple_sheets_and_suggestions(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Availability"
        sheet.append(["Pianist Name", "Pianist Email", *models.DAYS_ORDER])
        sheet.append(["Ari Example", "ari@example.invalid", "8:00 AM-9:30 AM; T 2:00 PM-3:00 PM", None, None, None, None, None, None])
        workbook.create_sheet("Unused Sheet").append(["Other", "Data"])
        contents = BytesIO()
        workbook.save(contents)

        inspection = availability_importer.inspect_upload("weekly.xlsx", contents.getvalue())

        self.assertEqual(inspection.sheets, ["Availability", "Unused Sheet"])
        self.assertEqual(inspection.suggested_columns["person_name_column"], "Pianist Name")
        self.assertEqual(inspection.suggested_columns["email_column"], "Pianist Email")
        self.assertEqual(inspection.suggested_days["Monday"], "Monday")
        self.assertEqual(inspection.suggested_days["Sunday"], "Sunday")
        other_sheet = availability_importer.inspect_sheet(inspection.upload_token, "Unused Sheet")
        self.assertEqual(other_sheet.columns, ["Other", "Data"])
        preview = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(
                upload_token=inspection.upload_token,
                mapping=self.normalized_mapping(inspection, email="Pianist Email"),
            ),
            self.db,
        )

        self.assertEqual(preview.valid_window_count, 2)
        self.assertEqual([window.start_minute for window in preview.windows], [480, 840])
        self.assertEqual([window.status for window in preview.windows], ["Available", "Tentative"])
        self.assertEqual(preview.pianists[0].email, "ari@example.invalid")
        self.assertTrue(preview.can_apply)

    def test_meaningfully_different_names_create_separate_pianists(self):
        _, preview = self.preview_csv(
            "Pianist Name,Monday\n"
            "Alex Smith,08:00-09:00\n"
            "Alex J. Smith,09:00-10:00\n"
        )

        self.assertTrue(preview.can_apply)
        self.assertEqual(preview.incoming_pianist_count, 2)
        self.assertEqual({person.pianist_name for person in preview.pianists}, {"Alex Smith", "Alex J. Smith"})
        availability_importer.apply_import(
            schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
            self.db,
        )
        self.assertEqual(self.db.query(models.Pianist).count(), 2)

    def test_invalid_mapping_day_time_range_status_and_conflict_block_all_writes(self):
        missing_mapping_file = self.csv_upload(
            "Pianist Name,Monday\nAri Example,08:00-09:00\n"
        )
        missing_mapping = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(
                upload_token=missing_mapping_file.upload_token,
                mapping=schemas.AvailabilityImportMapping(),
            ),
            self.db,
        )
        self.assertIn("MISSING_MAPPING", {issue.code for issue in missing_mapping.errors})
        self.assertFalse(missing_mapping.can_apply)

        _, missing_mapping = self.preview_csv(
            "Pianist Name,Monday\nAri Example,08:00-09:00\n",
            person="Does Not Exist",
        )
        self.assertIn("INVALID_MAPPING", {issue.code for issue in missing_mapping.errors})

        _, invalid_rows = self.preview_csv(
            "Pianist Name,Monday,Tuesday,Wednesday,Thursday\n"
            "Ari Example,25:00-26:00,10:00-10:00,08:00,junk\n"
        )
        self.assertEqual(len(invalid_rows.errors), 4)
        self.assertEqual({issue.code for issue in invalid_rows.errors}, {"INVALID_AVAILABILITY"})
        self.assertFalse(invalid_rows.can_apply)

        _, conflict = self.preview_csv(
            "Pianist Name,Monday\n"
            'Ari Example,"9:00-12:00; T 10:00-11:00"\n'
        )
        self.assertFalse(conflict.can_apply)
        self.assertIn("overlaps", conflict.errors[0].message)

        inspection = self.csv_upload("Pianist Name,Monday\nAri Example,9:00-10:00\n")
        duplicated = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(
                upload_token=inspection.upload_token,
                mapping=schemas.AvailabilityImportMapping(
                    person_name_column="Pianist Name",
                    days={"Monday": "Pianist Name"},
                ),
            ),
            self.db,
        )
        self.assertIn("DUPLICATE_MAPPING", {issue.code for issue in duplicated.errors})
        self.assertEqual(self.db.query(models.AvailabilitySlot).count(), 0)

    def test_full_replacement_clears_old_assignments_availability_and_absent_pianists(self):
        ari_id = self.pianist_id("Ari Example")
        bea_id = self.pianist_id("Bea Sample")
        old_person_uuid = ensure_pianist_identity(self.db, self.db.get(models.Pianist, ari_id))
        bea_person_uuid = ensure_pianist_identity(self.db, self.db.get(models.Pianist, bea_id))
        first_date = date(2027, 12, 14)
        second_date = date(2027, 12, 15)
        lesson = models.Lesson(
            student="Synthetic Student",
            day="Monday",
            start_minute=480,
            end_minute=510,
            assigned_pianist_id=ari_id,
        )
        self.db.add(lesson)
        self.db.flush()
        student_person_uuid = synchronize_lesson_identity(self.db, lesson)
        lesson_identity = self.db.get(module_models.AccompanistLessonIdentity, lesson.id)
        jury_panel = module_models.JuryPanel(
            session_uuid=active_session_uuid(self.db),
            panel_name="Panel A",
            earliest_start_minute=480,
            preferred_start_minute=540,
            jury_length_minutes=10,
        )
        second_panel = module_models.JuryPanel(
            session_uuid=active_session_uuid(self.db),
            panel_name="Panel B",
            earliest_start_minute=480,
            preferred_start_minute=540,
            jury_length_minutes=10,
        )
        self.db.add_all([jury_panel, second_panel])
        self.db.flush()
        self.db.add_all([
            module_models.JuryPanelDate(
                panel_uuid=jury_panel.panel_uuid,
                session_uuid=active_session_uuid(self.db),
                jury_date=first_date,
            ),
            module_models.JuryPanelDate(
                panel_uuid=second_panel.panel_uuid,
                session_uuid=active_session_uuid(self.db),
                jury_date=second_date,
            ),
        ])
        self.db.add_all([
            module_models.AccompanistLessonJuryRequirement(lesson_id=lesson.id, jury_required=True),
            module_models.JuryLessonEntry(
                session_uuid=active_session_uuid(self.db),
                source_lesson_uuid=lesson_identity.lesson_uuid,
                student_person_uuid=student_person_uuid,
                jury_required=True,
                panel_uuid=jury_panel.panel_uuid,
            ),
        ])
        revision = module_revision(self.db)
        initial_source_revision = revision.source_revision
        previous_result_uuid = str(uuid4())
        self.db.add(module_models.ModuleResult(
            result_uuid=previous_result_uuid,
            session_uuid=revision.session_uuid,
            module_id="accompanists",
            contract_id="accompanist.assignment-result",
            contract_version=2,
            payload_schema_version=1,
            result_version=1,
            source_revision=revision.source_revision,
            state="finalized",
            payload_json="{}",
            payload_sha256="0" * 64,
        ))
        revision.current_result_uuid = previous_result_uuid
        for pianist_uuid in (old_person_uuid, bea_person_uuid):
            for jury_date in (first_date, second_date):
                self.add_jury_availability(pianist_uuid, jury_date)
        self.db.add_all([
            models.AvailabilitySlot(pianist_id=ari_id, day="Monday", slot_start_minute=480, status="Available"),
            models.AvailabilitySlot(pianist_id=ari_id, day="Monday", slot_start_minute=510, status="Tentative"),
            models.AvailabilitySlot(pianist_id=ari_id, day="Tuesday", slot_start_minute=540, status="Available"),
            models.AvailabilitySlot(pianist_id=bea_id, day="Monday", slot_start_minute=600, status="Available"),
        ])
        self.db.commit()
        self.assertEqual(self.db.query(module_models.JuryPianistAvailabilityDeclaration).count(), 4)
        self.assertEqual(self.db.query(module_models.JuryPianistAvailableWindow).count(), 4)
        jury_configuration = self.db.get(module_models.JuryConfiguration, active_session_uuid(self.db))
        initial_jury_revision = jury_configuration.input_revision
        _, preview = self.preview_csv(
            "Pianist Name,Tuesday,Thursday\n"
            "Ari Example,10:00-11:00,T 09:00-10:00\n"
        )

        self.assertEqual(preview.existing_pianist_count, 2)
        self.assertEqual(preview.existing_assignment_count, 1)
        self.assertEqual(preview.existing_slots_in_scope, 4)
        result = availability_importer.apply_import(
            schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
            self.db,
        )

        imported = self.db.query(models.Pianist).one()
        new_person_uuid = self.db.get(
            module_models.AccompanistPianistIdentity,
            imported.id,
        ).person_uuid
        slots = self.db.query(models.AvailabilitySlot).order_by(
            models.AvailabilitySlot.day, models.AvailabilitySlot.slot_start_minute
        ).all()
        self.assertEqual(result.pianists_removed, 2)
        self.assertEqual(result.assignments_cleared, 1)
        self.assertEqual(result.slots_replaced, 4)
        self.assertEqual(result.slots_created, 4)
        self.assertEqual(result.jury_availability_windows_removed, 4)
        self.assertEqual(
            self.db.get(module_models.JuryConfiguration, active_session_uuid(self.db)).input_revision,
            initial_jury_revision + 1,
        )
        self.assertEqual(module_revision(self.db).source_revision, initial_source_revision + 1)
        self.assertIsNone(module_revision(self.db).current_result_uuid)
        self.assertEqual(self.db.get(module_models.ModuleResult, previous_result_uuid).state, "superseded")
        self.assertIsNone(self.db.get(models.Lesson, lesson.id).assigned_pianist_id)
        self.assertNotEqual(new_person_uuid, old_person_uuid)
        self.assertEqual(self.db.query(models.Pianist).filter_by(name="Bea Sample").count(), 0)
        self.assertEqual(
            self.db.get(module_models.AccompanistLessonIdentity, lesson.id).lesson_uuid,
            lesson_identity.lesson_uuid,
        )
        self.assertEqual(
            self.db.get(
                module_models.JuryLessonEntry,
                (active_session_uuid(self.db), lesson_identity.lesson_uuid),
            ).panel_uuid,
            jury_panel.panel_uuid,
        )
        self.assertTrue(self.db.get(module_models.AccompanistLessonJuryRequirement, lesson.id).jury_required)
        self.assertEqual(self.db.query(models.Lesson).count(), 1)
        self.assertEqual(self.db.query(module_models.JuryPianistAvailabilityDeclaration).count(), 0)
        self.assertEqual(self.db.query(module_models.JuryPianistAvailableWindow).count(), 0)
        self.assertEqual(self.db.query(module_models.JuryPanel).count(), 2)
        self.assertEqual(self.db.query(module_models.JuryPanelDate).count(), 2)
        self.assertEqual(
            {row.jury_date for row in self.db.query(module_models.JuryPanelDate).all()},
            {first_date, second_date},
        )
        self.assertEqual(
            [(slot.pianist_id, slot.day, slot.slot_start_minute, slot.status) for slot in slots],
            [
                (imported.id, "Thursday", 540, "Tentative"),
                (imported.id, "Thursday", 570, "Tentative"),
                (imported.id, "Tuesday", 600, "Available"),
                (imported.id, "Tuesday", 630, "Available"),
            ],
        )
        self.assertTrue(imported.availability_complete)
        self.assertNotIn("Monday", {slot.day for slot in slots})

        from app.services.scheduling import FIT_FULL, FIT_NONE, get_fit

        availability_by_day = {}
        for slot in slots:
            availability_by_day.setdefault(slot.day, {})[slot.slot_start_minute] = slot.status
        self.assertEqual(get_fit("Monday", 480, 510, availability_by_day)[0], FIT_NONE)
        self.assertEqual(get_fit("Wednesday", 480, 510, availability_by_day)[0], FIT_NONE)

        pianists.set_availability(
            imported.id,
            schemas.AvailabilityBulkIn(slots=[
                schemas.AvailabilitySlotIn(day="Monday", slot_start_minute=480, status="Available"),
            ]),
            self.db,
        )
        availability = {"Monday": {slot.slot_start_minute: slot.status for slot in self.db.query(models.AvailabilitySlot).filter_by(pianist_id=imported.id).all()}}
        self.assertTrue(self.db.get(models.Pianist, imported.id).availability_complete)
        self.assertEqual(get_fit("Monday", 480, 510, availability)[0], FIT_FULL)

    def test_valid_zero_window_submission_clears_old_slots_and_marks_week_complete(self):
        ari_id = self.pianist_id("Ari Example")
        self.db.add(models.AvailabilitySlot(pianist_id=ari_id, day="Tuesday", slot_start_minute=480, status="Available"))
        self.db.commit()
        _, preview = self.preview_csv(
            "Pianist Name,Monday\nAri Example,\n",
        )
        self.assertTrue(preview.can_apply)
        self.assertEqual(preview.valid_window_count, 0)
        self.assertEqual(preview.errors, [])

        result = availability_importer.apply_import(
            schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
            self.db,
        )
        slots = self.db.query(models.AvailabilitySlot).filter_by(pianist_id=ari_id).all()
        self.assertEqual(result.slots_created, 0)
        self.assertEqual(slots, [])
        self.assertTrue(self.db.get(models.Pianist, ari_id).availability_complete)

    def test_malformed_complete_submission_does_not_replace_or_complete_availability(self):
        ari_id = self.pianist_id("Ari Example")
        self.db.add(models.AvailabilitySlot(
            pianist_id=ari_id,
            day="Tuesday",
            slot_start_minute=600,
            status="Available",
        ))
        self.db.commit()
        inspection = self.csv_upload(
            "Pianist Name,Tuesday\nAri Example,10:00-\n"
        )
        preview = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(
                upload_token=inspection.upload_token,
                mapping=self.normalized_mapping(inspection),
            ),
            self.db,
        )

        self.assertFalse(preview.can_apply)
        self.assertIn("INVALID_AVAILABILITY", {issue.code for issue in preview.errors})
        with self.assertRaises(availability_importer.AvailabilityImportError):
            availability_importer.apply_import(
                schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
                self.db,
            )
        self.assertEqual(
            [(slot.day, slot.slot_start_minute, slot.status) for slot in self.db.query(models.AvailabilitySlot).all()],
            [("Tuesday", 600, "Available")],
        )
        self.assertFalse(self.db.get(models.Pianist, ari_id).availability_complete)

    def test_imported_availability_survives_session_archive_round_trip(self):
        from app.services.session_files import create_new_session, export_session_archive, restore_session

        with tempfile.TemporaryDirectory() as directory:
            database_path = f"{directory}/session.sqlite3"
            engine = create_engine(f"sqlite:///{database_path}", connect_args={"check_same_thread": False})
            migrate_database(engine)
            session = Session(engine)
            inspection = availability_importer.inspect_upload(
                "round-trip.csv",
                b"Pianist Name,Monday\nRound Trip Pianist,08:00-09:00\n",
            )
            preview = availability_importer.preview_import(
                schemas.AvailabilityImportPreviewRequest(
                    upload_token=inspection.upload_token,
                    mapping=schemas.AvailabilityImportMapping(
                        person_name_column="Pianist Name",
                        days={"Monday": "Monday"},
                    ),
                ),
                session,
            )
            availability_importer.apply_import(
                schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
                session,
            )
            imported_pianist = session.query(models.Pianist).one()
            imported_person_uuid = session.get(
                module_models.AccompanistPianistIdentity,
                imported_pianist.id,
            ).person_uuid
            session.close()
            archive = export_session_archive(engine)
            create_new_session(
                engine,
                schemas.SessionMetadataIn(
                    institution_name="Synthetic University",
                    program_name="Music",
                    term_label="Spring 2027",
                ),
            )
            restore_session(engine, archive)
            with Session(engine) as restored_db:
                restored_slots = restored_db.query(models.AvailabilitySlot).all()
                self.assertEqual(
                    [(slot.day, slot.slot_start_minute, slot.status) for slot in restored_slots],
                    [("Monday", minute, "Available") for minute in (480, 510)],
                )
                restored_pianist = restored_db.query(models.Pianist).one()
                self.assertTrue(restored_pianist.availability_complete)
                identity_mapping = restored_db.get(module_models.AccompanistPianistIdentity, restored_pianist.id)
                self.assertIsNotNone(identity_mapping)
                self.assertEqual(identity_mapping.person_uuid, imported_person_uuid)
                self.assertIsNotNone(restored_db.get(module_models.PersonIdentity, identity_mapping.person_uuid))
            engine.dispose()


if __name__ == "__main__":
    unittest.main()