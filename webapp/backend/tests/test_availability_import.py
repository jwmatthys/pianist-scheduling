import unittest
from io import BytesIO
from datetime import time
import tempfile
from unittest.mock import patch

from openpyxl import Workbook
from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from sqlalchemy.pool import StaticPool

from app import models, schemas
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
        Base.metadata.create_all(self.engine)
        self.db = Session(self.engine)
        self.db.add(models.Organization(id=1, name="Synthetic Availability Program"))
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

    def normalized_mapping(self, inspection, person="Pianist Name"):
        return schemas.AvailabilityImportMapping(
            layout="normalized",
            person_name_column=person,
            day_column="Weekday",
            start_column="Start",
            end_column="End",
            status_column="Status",
        )

    def preview_csv(self, content, *, person="Pianist Name"):
        inspection = self.csv_upload(content)
        preview = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(
                upload_token=inspection.upload_token,
                mapping=self.normalized_mapping(inspection, person),
            ),
            self.db,
        )
        return inspection, preview

    def test_normalized_csv_preview_and_apply_import_multiple_windows(self):
        _, preview = self.preview_csv(
            "Pianist Name,Weekday,Start,End,Status\n"
            "Ari Example,Mon,08:00,09:30,Available\n"
            "Ari Example,Monday,14:00,15:00,Tentative\n"
        )

        self.assertTrue(preview.can_apply)
        self.assertEqual(preview.rows_processed, 2)
        self.assertEqual(preview.matched_pianist_count, 1)
        self.assertEqual(preview.valid_window_count, 2)
        self.assertEqual([window.status for window in preview.windows], ["Available", "Tentative"])
        self.assertEqual(self.db.query(models.AvailabilitySlot).count(), 0)

        result = availability_importer.apply_import(
            schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
            self.db,
        )

        self.assertEqual((result.pianists_updated, result.slots_created), (1, 5))
        self.assertEqual(self.db.query(models.AvailabilitySlot).count(), 5)

    def test_xlsx_actual_time_cells_multiple_sheets_and_forms_wide_suggestions(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Form Responses"
        headers = ["Full Name"]
        row = ["Ari Example"]
        for day in models.DAYS_ORDER:
            if day == "Monday":
                headers.extend(["Monday Start Time", "Monday End Time", "Monday Start 2", "Monday End 2"])
                row.extend([time(8, 0), time(9, 30), time(14, 0), time(15, 0)])
            else:
                headers.extend([f"{day} Start Time", f"{day} End Time"])
                row.extend([None, None])
        sheet.append(headers)
        sheet.append(row)
        workbook.create_sheet("Unused Sheet").append(["Other", "Data"])
        contents = BytesIO()
        workbook.save(contents)

        inspection = availability_importer.inspect_upload("forms.xlsx", contents.getvalue())

        self.assertEqual(inspection.sheets, ["Form Responses", "Unused Sheet"])
        self.assertEqual(len(inspection.suggested_wide_windows["Monday"]), 2)
        other_sheet = availability_importer.inspect_sheet(inspection.upload_token, "Unused Sheet")
        self.assertEqual(other_sheet.columns, ["Other", "Data"])
        mapping = schemas.AvailabilityImportMapping(
            layout="wide",
            person_name_column="Full Name",
            wide_status="Available",
            wide_windows=inspection.suggested_wide_windows,
        )
        preview = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(upload_token=inspection.upload_token, mapping=mapping),
            self.db,
        )

        self.assertEqual(preview.valid_window_count, 2)
        self.assertEqual([window.start_minute for window in preview.windows], [480, 840])
        self.assertTrue(preview.can_apply)

    def test_normalized_xlsx_reads_excel_time_cells(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Pianist Name", "Weekday", "Start", "End", "Status"])
        sheet.append(["Ari Example", "Mon", time(8, 0), time(9, 30), "Available"])
        contents = BytesIO()
        workbook.save(contents)
        inspection = availability_importer.inspect_upload("normalized.xlsx", contents.getvalue())
        preview = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(
                upload_token=inspection.upload_token,
                mapping=schemas.AvailabilityImportMapping(
                    layout="normalized",
                    person_name_column="Pianist Name",
                    day_column="Weekday",
                    start_column="Start",
                    end_column="End",
                    status_column="Status",
                ),
            ),
            self.db,
        )

        self.assertEqual(preview.valid_window_count, 1)
        self.assertEqual((preview.windows[0].start_minute, preview.windows[0].end_minute), (480, 570))

    def test_unknown_and_ambiguous_pianist_names_are_errors(self):
        self.db.add(models.Pianist(name="Ari Example"))
        self.db.commit()
        _, preview = self.preview_csv(
            "Pianist Name,Weekday,Start,End,Status\n"
            "Missing Person,Mon,08:00,09:00,Available\n"
            "Ari Example,Mon,09:00,10:00,Available\n"
        )

        self.assertFalse(preview.can_apply)
        self.assertEqual({issue.code for issue in preview.errors}, {"UNKNOWN_PIANIST", "AMBIGUOUS_PIANIST"})
        with self.assertRaises(availability_importer.AvailabilityImportError):
            availability_importer.apply_import(
                schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
                self.db,
            )
        self.assertEqual(self.db.query(models.AvailabilitySlot).count(), 0)

    def test_invalid_mapping_day_time_range_status_and_conflict_block_all_writes(self):
        missing_mapping_file = self.csv_upload(
            "Pianist Name,Weekday,Start,End,Status\nAri Example,Mon,08:00,09:00,Available\n"
        )
        missing_mapping = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(
                upload_token=missing_mapping_file.upload_token,
                mapping=schemas.AvailabilityImportMapping(layout="normalized"),
            ),
            self.db,
        )
        self.assertIn("MISSING_MAPPING", {issue.code for issue in missing_mapping.errors})
        self.assertFalse(missing_mapping.can_apply)

        _, missing_mapping = self.preview_csv(
            "Pianist Name,Weekday,Start,End,Status\nAri Example,Mon,08:00,09:00,Available\n",
            person="Does Not Exist",
        )
        self.assertIn("INVALID_MAPPING", {issue.code for issue in missing_mapping.errors})

        _, invalid_rows = self.preview_csv(
            "Pianist Name,Weekday,Start,End,Status\n"
            "Ari Example,Funday,08:00,09:00,Available\n"
            "Ari Example,Mon,25:00,26:00,Available\n"
            "Ari Example,Tue,10:00,09:00,Available\n"
            "Ari Example,Wed,08:00,09:00,Maybe\n"
        )
        self.assertEqual(
            {issue.code for issue in invalid_rows.errors},
            {"INVALID_DAY", "INVALID_START_TIME", "INVALID_END_TIME", "INVALID_TIME_RANGE", "INVALID_STATUS"},
        )

        _, conflict = self.preview_csv(
            "Pianist Name,Weekday,Start,End,Status\n"
            "Ari Example,Mon,09:00,12:00,Available\n"
            "Ari Example,Mon,10:00,11:00,Tentative\n"
        )
        self.assertIn("CONFLICTING_OVERLAP", {issue.code for issue in conflict.errors})
        self.assertEqual(self.db.query(models.AvailabilitySlot).count(), 0)

    def test_complete_reimport_replaces_matched_week_and_leaves_absent_pianists_unchanged(self):
        ari_id = self.pianist_id("Ari Example")
        bea_id = self.pianist_id("Bea Sample")
        self.db.add_all([
            models.AvailabilitySlot(pianist_id=ari_id, day="Monday", slot_start_minute=480, status="Available"),
            models.AvailabilitySlot(pianist_id=ari_id, day="Monday", slot_start_minute=510, status="Tentative"),
            models.AvailabilitySlot(pianist_id=ari_id, day="Tuesday", slot_start_minute=540, status="Available"),
            models.AvailabilitySlot(pianist_id=bea_id, day="Monday", slot_start_minute=600, status="Available"),
        ])
        self.db.commit()
        _, preview = self.preview_csv(
            "Pianist Name,Weekday,Start,End,Status\n"
            "Ari Example,Tue,10:00,11:00,Available\n"
            "Ari Example,Thu,09:00,10:00,Tentative\n"
        )

        self.assertEqual(preview.existing_slots_in_scope, 3)
        self.assertEqual(preview.absent_pianist_count, 1)
        result = availability_importer.apply_import(
            schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
            self.db,
        )

        slots = self.db.query(models.AvailabilitySlot).order_by(
            models.AvailabilitySlot.pianist_id, models.AvailabilitySlot.day, models.AvailabilitySlot.slot_start_minute
        ).all()
        self.assertEqual(result.slots_replaced, 3)
        self.assertEqual(result.slots_created, 4)
        self.assertEqual(
            [(slot.pianist_id, slot.day, slot.slot_start_minute, slot.status) for slot in slots],
            [
                (ari_id, "Thursday", 540, "Tentative"),
                (ari_id, "Thursday", 570, "Tentative"),
                (ari_id, "Tuesday", 600, "Available"),
                (ari_id, "Tuesday", 630, "Available"),
                (bea_id, "Monday", 600, "Available"),
            ],
        )
        self.assertTrue(self.db.get(models.Pianist, ari_id).availability_complete)
        self.assertNotIn("Monday", {slot.day for slot in slots if slot.pianist_id == ari_id})

        from app.services.scheduling import FIT_NONE, get_fit

        availability_by_day = {}
        for slot in slots:
            if slot.pianist_id == ari_id:
                availability_by_day.setdefault(slot.day, {})[slot.slot_start_minute] = slot.status
        self.assertEqual(get_fit("Monday", 480, 510, availability_by_day)[0], FIT_NONE)
        self.assertEqual(get_fit("Wednesday", 480, 510, availability_by_day)[0], FIT_NONE)

        pianists.set_availability(
            ari_id,
            schemas.AvailabilityBulkIn(slots=[
                schemas.AvailabilitySlotIn(day="Monday", slot_start_minute=480, status="Available"),
            ]),
            self.db,
        )
        availability = {"Monday": {slot.slot_start_minute: slot.status for slot in self.db.query(models.AvailabilitySlot).filter_by(pianist_id=ari_id).all()}}
        from app.services.scheduling import FIT_FULL, get_fit

        self.assertTrue(self.db.get(models.Pianist, ari_id).availability_complete)
        self.assertEqual(get_fit("Monday", 480, 510, availability)[0], FIT_FULL)
        self.assertEqual(get_fit("Wednesday", 480, 510, availability)[0], FIT_NONE)

    def test_valid_zero_window_submission_clears_old_slots_and_marks_week_complete(self):
        ari_id = self.pianist_id("Ari Example")
        self.db.add(models.AvailabilitySlot(pianist_id=ari_id, day="Tuesday", slot_start_minute=480, status="Available"))
        self.db.commit()
        _, preview = self.preview_csv(
            "Pianist Name,Weekday,Start,End,Status\nAri Example,,,,\n",
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
            "Pianist Name,Weekday,Start,End,Status\nAri Example,Tuesday,10:00,,Available\n"
        )
        preview = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(
                upload_token=inspection.upload_token,
                mapping=self.normalized_mapping(inspection),
            ),
            self.db,
        )

        self.assertFalse(preview.can_apply)
        self.assertIn("INVALID_END_TIME", {issue.code for issue in preview.errors})
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

    def test_wide_import_missing_a_weekday_mapping_is_structurally_incomplete(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Name", "Monday Start", "Monday End"])
        sheet.append(["Ari Example", time(8, 0), time(9, 0)])
        content = BytesIO()
        workbook.save(content)
        inspection = availability_importer.inspect_upload("incomplete-wide.xlsx", content.getvalue())
        preview = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(
                upload_token=inspection.upload_token,
                mapping=schemas.AvailabilityImportMapping(
                    layout="wide",
                    person_name_column="Name",
                    wide_windows=inspection.suggested_wide_windows,
                ),
            ),
            self.db,
        )

        self.assertFalse(preview.can_apply)
        self.assertGreaterEqual(
            sum(issue.code == "MISSING_MAPPING" for issue in preview.errors),
            len(models.DAYS_ORDER) - 1,
        )
        self.assertEqual(self.db.query(models.AvailabilitySlot).count(), 0)

    def test_complete_wide_blank_days_are_unavailable_without_warnings(self):
        workbook = Workbook()
        sheet = workbook.active
        headers = ["Name"]
        row = ["Ari Example"]
        for day in models.DAYS_ORDER:
            headers.extend([f"{day} Start", f"{day} End"])
            if day == "Monday":
                row.extend([time(8, 0), time(9, 0)])
            else:
                row.extend([None, None])
        sheet.append(headers)
        sheet.append(row)
        content = BytesIO()
        workbook.save(content)
        inspection = availability_importer.inspect_upload("wide.xlsx", content.getvalue())
        mapping = schemas.AvailabilityImportMapping(
            layout="wide",
            person_name_column="Name",
            wide_windows=inspection.suggested_wide_windows,
        )
        preview = availability_importer.preview_import(
            schemas.AvailabilityImportPreviewRequest(upload_token=inspection.upload_token, mapping=mapping),
            self.db,
        )
        self.assertEqual(preview.pianists[0].days, ["Monday"])
        self.assertEqual(preview.valid_window_count, 1)
        self.assertFalse(any(issue.code in {"MISSING_VS_UNAVAILABLE", "MISSING_DAY_DATA", "NO_AVAILABILITY"} for issue in preview.warnings))
        self.assertTrue(preview.can_apply)

        availability_importer.apply_import(
            schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
            self.db,
        )
        slots = self.db.query(models.AvailabilitySlot).filter_by(pianist_id=self.pianist_id("Ari Example")).all()
        self.assertEqual(len(slots), 2)
        self.assertTrue(all(slot.status == "Available" for slot in slots))
        self.assertTrue(self.db.get(models.Pianist, self.pianist_id("Ari Example")).availability_complete)

    def test_imported_availability_survives_session_archive_round_trip(self):
        from app.services.session_files import create_new_session, export_session_archive, restore_session

        with tempfile.TemporaryDirectory() as directory:
            database_path = f"{directory}/session.sqlite3"
            engine = create_engine(f"sqlite:///{database_path}", connect_args={"check_same_thread": False})
            migrate_database(engine)
            session = Session(engine)
            session.add(models.Pianist(name="Round Trip Pianist"))
            session.commit()
            inspection = availability_importer.inspect_upload(
                "round-trip.csv",
                b"Pianist Name,Weekday,Start,End,Status\nRound Trip Pianist,Mon,08:00,09:00,Available\n",
            )
            preview = availability_importer.preview_import(
                schemas.AvailabilityImportPreviewRequest(
                    upload_token=inspection.upload_token,
                    mapping=schemas.AvailabilityImportMapping(
                        layout="normalized",
                        person_name_column="Pianist Name",
                        day_column="Weekday",
                        start_column="Start",
                        end_column="End",
                        status_column="Status",
                    ),
                ),
                session,
            )
            availability_importer.apply_import(
                schemas.AvailabilityImportApplyRequest(preview_token=preview.preview_token, confirmed=True),
                session,
            )
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
                self.assertTrue(restored_db.query(models.Pianist).one().availability_complete)
            engine.dispose()


if __name__ == "__main__":
    unittest.main()