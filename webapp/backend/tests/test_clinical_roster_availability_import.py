"""Synthetic-data tests for weekday availability columns in the unified Clinical Student Roster import."""

import asyncio
import io
import json
import unittest

import pandas as pd
from fastapi import HTTPException, UploadFile

from app.clinical_models import ClinicalPlacement, PlacementType
from app.routers import clinical
from app.services import clinical_students_importer as importer
from app.services.availability import AvailabilityWindow
from app.services.clinical_import_common import ClinicalImportError
from tests.test_clinical_foundation import ClinicalTestCase, opportunity_input

HEADERS = ["Student ID", "Mon", "Tue", "Wed", "Notes"]
MAPPING = {"student_id": "Student ID", "days": {"Monday": "Mon", "Tuesday": "Tue", "Wednesday": "Wed"}}


def csv_bytes(rows, headers=HEADERS) -> bytes:
    return pd.DataFrame(rows, columns=headers).to_csv(index=False).encode()


def xlsx_bytes(rows, headers=HEADERS) -> bytes:
    buffer = io.BytesIO()
    pd.DataFrame(rows, columns=headers).to_excel(buffer, index=False)
    return buffer.getvalue()


def upload(data: bytes, name="roster.csv") -> UploadFile:
    return UploadFile(file=io.BytesIO(data), filename=name)


ROSTER = {
    "s1": importer.ExistingStudent("pk1", "Ada Example", "", 3),
    "s2": importer.ExistingStudent("pk2", "Ben Sample", "", 0),
}


def parse(rows, name="roster.csv", data=None, mapping=None, roster=ROSTER):
    return importer.parse_students(
        name,
        data if data is not None else csv_bytes(rows),
        None,
        importer.ClinicalStudentsImportMapping(**(mapping or MAPPING)),
        roster,
    )


def triples(row):
    return [(w.day, w.start_minute, w.end_minute, w.status) for w in row.windows]


class ParserTests(unittest.TestCase):
    def test_single_range_defaults_to_available(self):
        parsed = parse([["S1", "8:00-12:00", "", "", ""]])
        self.assertEqual(triples(parsed.rows[0]), [("Monday", 480, 720, "Available")])
        self.assertEqual(parsed.rows[0].name, "Ada Example")
        self.assertEqual(parsed.rows[0].errors, [])

    def test_time_formats_and_whitespace(self):
        for text in ("8-12", "8:00 AM - 12:00 PM", "  8 AM-12 PM ", "08:00 to 12:00", "8:00–12:00"):
            parsed = parse([["s1", text, "", "", ""]])
            self.assertEqual(triples(parsed.rows[0]), [("Monday", 480, 720, "Available")], text)
        parsed = parse([["S1", "13:00-15:30", "", "", ""]])
        self.assertEqual(triples(parsed.rows[0]), [("Monday", 780, 930, "Available")])

    def test_unmarked_reversed_ranges_assume_afternoon_end(self):
        for text, expected in (
            ("9:00-3:00", (540, 900)),
            ("8:00-1:00", (480, 780)),
            ("10:30-2:15", (630, 855)),
        ):
            with self.subTest(text=text):
                row = parse([["S1", text, "", "", ""]]).rows[0]
                self.assertEqual([(window.start_minute, window.end_minute) for window in row.windows], [expected])
                self.assertEqual(row.errors, [])

    def test_explicit_meridiem_is_not_adjusted(self):
        row = parse([["S1", "9:00 AM-3:00 AM", "", "", ""]]).rows[0]
        self.assertTrue(any("ends at or before" in error for error in row.errors), row.errors)

    def test_tentative_prefixes(self):
        for text in ("T 1:00 PM-4:00 PM", "Tentative 1:00 PM-4:00 PM", "t 13:00-16:00", "T 1:00-4:00"):
            parsed = parse([["S1", text, "", "", ""]])
            window = parsed.rows[0].windows[0]
            self.assertEqual(window.status, "Tentative", text)
        self.assertEqual(parse([["S1", "Available: 8:00-12:00", "", "", ""]]).rows[0].windows[0].status, "Available")

    def test_multiple_and_mixed_entries(self):
        parsed = parse([["S1", "8:00-10:00; T 1:00-4:00", "9-10", "", ""]])
        self.assertEqual(
            triples(parsed.rows[0]),
            [
                ("Monday", 480, 600, "Available"),
                ("Monday", 780, 960, "Tentative"),
                ("Tuesday", 540, 600, "Available"),
            ],
        )

    def test_unmarked_early_hours_read_as_afternoon(self):
        parsed = parse([["S1", "8:00-10:00; T 1:00-4:00", "6-8 AM", "", ""]])
        self.assertEqual(
            triples(parsed.rows[0]),
            [("Monday", 480, 600, "Available"), ("Monday", 780, 960, "Tentative"),
             ("Tuesday", 360, 480, "Available")],
        )

    def test_blank_cells_have_no_windows(self):
        parsed = parse([["S1", "", "  ", "", ""]])
        self.assertEqual(parsed.rows[0].windows, [])
        self.assertEqual(parsed.rows[0].errors, [])

    def test_xlsx_matches_csv(self):
        rows = [["S2", "8-10; T 1-4", "", "", ""]]
        from_csv = parse(rows)
        from_xlsx = parse(rows, name="roster.xlsx", data=xlsx_bytes(rows))
        self.assertEqual(triples(from_csv.rows[0]), triples(from_xlsx.rows[0]))

    def test_same_status_overlaps_merge(self):
        parsed = parse([["S1", "8-10; 9-12; 8-10", "", "", ""]])
        self.assertEqual(triples(parsed.rows[0]), [("Monday", 480, 720, "Available")])
        self.assertEqual(parsed.rows[0].errors, [])

    def test_invalid_times_and_ranges(self):
        cases = {
            "noon-5": "could not be read",
            "8:00": "not a valid time range",
            "12:00 PM-8:00 AM": "ends at or before",
            "9-9": "ends at or before",
            "8-10-12": "not a valid time range",
        }
        for text, fragment in cases.items():
            row = parse([["S1", text, "", "", ""]]).rows[0]
            self.assertTrue(any(fragment in error for error in row.errors), (text, row.errors))

    def test_conflicting_overlaps_are_errors(self):
        row = parse([["S1", "8-12; T 10:00-14:00", "", "", ""]]).rows[0]
        self.assertTrue(any("overlaps" in error for error in row.errors), row.errors)

    def test_student_matching(self):
        parsed = parse([["", "8-9", "", "", ""], ["S9", "8-9", "", "", ""], ["S1", "8-9", "", "", ""],
                        ["s1", "8-9", "", "", ""]])
        self.assertIn("Student ID is required.", parsed.rows[0].errors)
        self.assertIn("Name is required to create a new student; map the Name column.", parsed.rows[1].errors)
        self.assertTrue(any("duplicated" in error for error in parsed.rows[2].errors))
        self.assertTrue(any("duplicated" in error for error in parsed.rows[3].errors))

    def test_mapping_validation(self):
        columns = list(HEADERS)
        bad = [
            importer.ClinicalStudentsImportMapping(days={"Monday": "Mon"}),
            importer.ClinicalStudentsImportMapping(student_id="Student ID", days={"Monday": "Missing"}),
            importer.ClinicalStudentsImportMapping(student_id="Mon", days={"Monday": "Mon"}),
            importer.ClinicalStudentsImportMapping(student_id="Student ID", days={"Someday": "Mon"}),
        ]
        for mapping in bad:
            with self.assertRaises(ClinicalImportError) as caught:
                importer.validate_mapping(mapping, columns)
            self.assertEqual(caught.exception.code, "INVALID_MAPPING")

    def test_inspect_suggests_mapping(self):
        headers = ["Student ID", "Monday", "tue", "Wed.", "Thursday", "Notes"]
        inspection = importer.inspect_file("a.csv", csv_bytes([["S1", "", "", "", "", ""]], headers), None)
        mapping = inspection.suggested_mapping
        self.assertEqual(mapping.student_id, "Student ID")
        self.assertEqual(mapping.days["Monday"], "Monday")
        self.assertEqual(mapping.days["Tuesday"], "tue")
        self.assertEqual(mapping.days["Wednesday"], "Wed.")
        self.assertIsNone(mapping.days["Friday"])

    def test_preview_change_descriptions(self):
        parsed = parse([["S1", "8-9; 10-11", "", "", ""], ["S2", "8-9", "", "", ""], ["S1", "", "", "", ""]][:2])
        preview = importer.build_preview(parsed)
        self.assertEqual(
            [row.availability_change for row in preview.rows],
            ["Replace 3 availability windows with 2", "Add 1 availability window"],
        )
        cleared = importer.build_preview(parse([["S1", "", "", "", ""]]))
        self.assertEqual(cleared.rows[0].availability_change, "Clear availability (3 windows)")
        self.assertTrue(cleared.can_apply)

    def test_invalid_rows_block_apply(self):
        preview = importer.build_preview(parse([["S1", "8-9", "", "", ""], ["S9", "8-9", "", "", ""]]))
        self.assertFalse(preview.can_apply)
        self.assertEqual((preview.valid_rows, preview.invalid_rows), (1, 1))


class RepositoryAndApiTests(ClinicalTestCase):
    def seed(self):
        self.ada = self.repo.create_student("S1", "Ada Example", require_name=True)
        self.ben = self.repo.create_student("S2", "Ben Sample", require_name=True)
        self.cy = self.repo.create_student("S3", "Cy Untouched", require_name=True)
        self.repo.set_student_availability(self.ada.id, [AvailabilityWindow("Monday", 480, 600, "Available")])
        self.repo.set_student_availability(self.cy.id, [AvailabilityWindow("Friday", 600, 660, "Tentative")])
        self.site = self.repo.create_opportunity(opportunity_input())
        self.repo.set_placement(self.ada.id, PlacementType.MANUAL, self.site.id, 480)
        self.db.commit()

    def call(self, name, rows, confirmed=True, mapping=MAPPING, data=None):
        data = data if data is not None else csv_bytes(rows)
        if name == "apply":
            return asyncio.run(clinical.apply_students_import(
                file=upload(data), mapping=json.dumps(mapping), sheet_name=None, confirmed=confirmed, db=self.db
            ))
        return asyncio.run(clinical.preview_students_import(
            file=upload(data), mapping=json.dumps(mapping), sheet_name=None, db=self.db
        ))

    def test_preview_writes_nothing(self):
        self.seed()
        state = self.repo.get_state()
        revision, stale = state.source_data_revision, state.is_schedule_stale
        preview = self.call("preview", [["S1", "8-12", "", "", ""], ["S2", "T 1-4", "", "", ""]])
        self.assertTrue(preview.can_apply)
        self.assertEqual(preview.rows[0].availability_change, "Replace 1 availability window with 1")
        self.db.rollback()
        state = self.repo.get_state()
        self.assertEqual((state.source_data_revision, state.is_schedule_stale), (revision, stale))
        self.assertEqual(len(self.repo.student_availability(self.ada.id)), 1)
        self.assertEqual(self.repo.student_availability(self.ben.id), [])

    def test_apply_replaces_only_matched_and_preserves_everything_else(self):
        self.seed()
        before = self.repo.get_state().source_data_revision
        result = self.call("apply", [["s1", "8-10; T 1-4", "9-11", "", ""], ["S2", "", "", "", ""]])
        self.assertEqual((result.students_created, result.students_updated, result.availability_updated, result.windows_imported), (0, 2, 2, 3))
        ada = self.repo.student_availability(self.ada.id)
        self.assertEqual(
            [(w.day, w.start_minute, w.end_minute, w.status, w.source) for w in ada],
            [("Monday", 480, 600, "Available", "import"), ("Monday", 780, 960, "Tentative", "import"),
             ("Tuesday", 540, 660, "Available", "import")],
        )
        self.assertEqual(self.repo.student_availability(self.ben.id), [])
        self.assertEqual(len(self.repo.student_availability(self.cy.id)), 1)
        self.assertEqual([s.student_id for s in self.repo.list_students()], ["S1", "S2", "S3"])
        self.assertEqual(len(self.repo.list_opportunities()), 1)
        self.assertEqual(self.db.query(ClinicalPlacement).count(), 1)
        state = self.repo.get_state()
        self.assertEqual(state.source_data_revision, before + 1)
        self.assertTrue(state.is_schedule_stale)

    def test_confirmation_required(self):
        self.seed()
        with self.assertRaises(HTTPException) as caught:
            self.call("apply", [["S1", "8-9", "", "", ""]], confirmed=False)
        self.assertEqual(caught.exception.detail["code"], "CONFIRMATION_REQUIRED")
        self.assertEqual(self.repo.student_availability(self.ada.id)[0].end_minute, 600)

    def test_any_invalid_row_rejects_entire_import(self):
        self.seed()
        before = self.repo.get_state().source_data_revision
        with self.assertRaises(HTTPException) as caught:
            self.call("apply", [["S1", "8-9", "", "", ""], ["S9", "8-9", "", "", ""], ["S2", "bad", "", "", ""]])
        self.assertEqual(caught.exception.detail["code"], "VALIDATION_FAILED")
        self.assertEqual(len(caught.exception.detail["issues"]), 2)
        self.assertEqual(self.repo.student_availability(self.ada.id)[0].end_minute, 600)
        self.assertEqual(self.repo.student_availability(self.ben.id), [])
        self.assertEqual(self.repo.get_state().source_data_revision, before)

    def test_empty_import_rejected(self):
        self.seed()
        with self.assertRaises(HTTPException) as caught:
            self.call("apply", [])
        self.assertEqual(caught.exception.detail["code"], "EMPTY_IMPORT")

    def test_stale_only_when_schedule_exists(self):
        self.ada = self.repo.create_student("S1", "Ada Example", require_name=True)
        self.db.commit()
        self.call("apply", [["S1", "8-9", "", "", ""]])
        self.assertFalse(self.repo.get_state().is_schedule_stale)

    def test_xlsx_apply(self):
        self.seed()
        rows = [["S2", "T 2-3", "", "", ""]]
        asyncio.run(clinical.apply_students_import(
            file=upload(xlsx_bytes(rows), "a.xlsx"), mapping=json.dumps(MAPPING), sheet_name=None,
            confirmed=True, db=self.db,
        ))
        self.assertEqual([w.status for w in self.repo.student_availability(self.ben.id)], ["Tentative"])


FULL_HEADERS = ["Student ID", "Name", "Email", "Mon", "Tue"]
FULL_MAPPING = {
    "student_id": "Student ID", "name": "Name", "email": "Email", "days": {"Monday": "Mon", "Tuesday": "Tue"},
}
ROSTER_ONLY = {"student_id": "Student ID", "name": "Name", "email": "Email"}


class UnifiedImportTests(ClinicalTestCase):
    def seed(self):
        self.ada = self.repo.create_student("S1", "Ada Example", "ada@example.test", require_name=True)
        self.repo.set_student_availability(self.ada.id, [AvailabilityWindow("Monday", 480, 600, "Available")])
        self.site = self.repo.create_opportunity(opportunity_input())
        self.repo.set_placement(self.ada.id, PlacementType.MANUAL, self.site.id, 480)
        self.db.commit()

    def run_apply(self, rows, mapping):
        return asyncio.run(clinical.apply_students_import(
            file=upload(csv_bytes(rows, FULL_HEADERS)), mapping=json.dumps(mapping), sheet_name=None,
            confirmed=True, db=self.db,
        ))

    def run_preview(self, rows, mapping):
        return asyncio.run(clinical.preview_students_import(
            file=upload(csv_bytes(rows, FULL_HEADERS)), mapping=json.dumps(mapping), sheet_name=None, db=self.db
        ))

    def revision(self):
        return self.repo.get_state().source_data_revision

    def test_roster_and_availability_in_one_import_bumps_once(self):
        self.seed()
        before = self.revision()
        result = self.run_apply(
            [["s1", "Ada E.", "new@example.test", "8-9", "T 1-2"], ["S2", "Ben Sample", "", "", ""]], FULL_MAPPING
        )
        self.assertEqual(
            (result.students_created, result.students_updated, result.availability_updated, result.windows_imported),
            (1, 1, 2, 2),
        )
        ada = next(s for s in self.repo.list_students() if s.student_id == "S1")
        self.assertEqual((ada.name, ada.email), ("Ada E.", "new@example.test"))
        self.assertEqual(len(self.repo.student_availability(ada.id)), 2)
        self.assertEqual(self.revision(), before + 1)
        self.assertTrue(self.repo.get_state().is_schedule_stale)
        self.assertEqual(self.db.query(ClinicalPlacement).count(), 1)

    def test_roster_only_leaves_availability_and_bumps_once(self):
        self.seed()
        before = self.revision()
        result = self.run_apply([["S1", "Ada Renamed", "", "x", "y"]], ROSTER_ONLY)
        self.assertEqual((result.availability_updated, result.windows_imported), (0, 0))
        self.assertEqual(len(self.repo.student_availability(self.ada.id)), 1)
        self.assertEqual(self.revision(), before + 1)

    def test_availability_only_keeps_name_and_email(self):
        self.seed()
        before = self.revision()
        mapping = {"student_id": "Student ID", "days": FULL_MAPPING["days"]}
        result = self.run_apply([["S1", "", "", "", "9-10"]], mapping)
        self.assertEqual((result.students_created, result.students_updated, result.availability_updated), (0, 1, 1))
        ada = self.repo.list_students()[0]
        self.assertEqual((ada.name, ada.email), ("Ada Example", "ada@example.test"))
        self.assertEqual([w.day for w in self.repo.student_availability(self.ada.id)], ["Tuesday"])
        self.assertEqual(self.revision(), before + 1)

    def test_all_blank_mapped_days_clear_availability(self):
        self.seed()
        self.run_apply([["S1", "Ada Example", "ada@example.test", "", ""]], FULL_MAPPING)
        self.assertEqual(self.repo.student_availability(self.ada.id), [])

    def test_absent_students_and_placements_are_preserved(self):
        self.seed()
        self.run_apply([["S2", "Ben Sample", "", "8-9", ""]], FULL_MAPPING)
        self.assertEqual({s.student_id for s in self.repo.list_students()}, {"S1", "S2"})
        self.assertEqual(len(self.repo.student_availability(self.ada.id)), 1)
        self.assertEqual(self.db.query(ClinicalPlacement).count(), 1)

    def test_preview_messages_and_no_writes(self):
        self.seed()
        before = self.revision()
        preview = self.run_preview(
            [["S1", "Ada Example", "other@example.test", "8-9; 10-11", ""], ["S2", "Ben Sample", "", "", ""]],
            FULL_MAPPING,
        )
        first, second = preview.rows
        self.assertEqual((first.action, first.student_change), ("update", "Update email"))
        self.assertEqual(first.availability_change, "Replace 1 availability window with 2")
        self.assertEqual((second.action, second.student_change), ("create", "Create student"))
        self.assertEqual((preview.students_to_create, preview.students_to_update), (1, 1))
        self.db.rollback()
        self.assertEqual(self.revision(), before)
        self.assertEqual(len(self.repo.list_students()), 1)
        unchanged = self.run_preview([["S1", "Ada Example", "ada@example.test", "", ""]], ROSTER_ONLY).rows[0]
        self.assertEqual((unchanged.student_change, unchanged.availability_change),
                         ("No student change", "No availability change"))

    def test_new_student_requires_name(self):
        self.seed()
        preview = self.run_preview([["S9", "", "", "8-9", ""]], FULL_MAPPING)
        self.assertIn("Name is required.", preview.rows[0].errors)
        self.assertFalse(preview.can_apply)
        with self.assertRaises(HTTPException):
            self.run_apply([["S9", "", "", "8-9", ""]], FULL_MAPPING)
        self.assertEqual(len(self.repo.list_students()), 1)

    def test_availability_error_rejects_whole_import(self):
        self.seed()
        before = self.revision()
        with self.assertRaises(HTTPException) as caught:
            self.run_apply([["S1", "New Name", "", "8-9", ""], ["S2", "Ben", "", "bad", ""]], FULL_MAPPING)
        self.assertEqual(caught.exception.detail["code"], "VALIDATION_FAILED")
        self.assertEqual(self.repo.list_students()[0].name, "Ada Example")
        self.assertEqual(self.revision(), before)


if __name__ == "__main__":
    unittest.main()
