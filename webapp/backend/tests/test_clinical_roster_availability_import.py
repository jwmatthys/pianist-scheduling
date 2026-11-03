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
from app.services.import_common import ImportFileError
from tests.test_clinical_foundation import ClinicalTestCase, opportunity_input

HEADERS = ["Student ID", "Mon", "Tue", "Wed", "Notes"]
MAPPING = {
    "student_id": "Student ID", "name": "Name",
    "days": {"Monday": "Mon", "Tuesday": "Tue", "Wednesday": "Wed"},
}


def frame(rows, headers):
    """Every roster row needs a Name; tests using HEADERS get a synthetic Name column."""
    if headers is HEADERS:
        return pd.DataFrame([[row[0], "Test Name", *row[1:]] for row in rows], columns=["Student ID", "Name", *headers[1:]])
    return pd.DataFrame(rows, columns=headers)


def csv_bytes(rows, headers=HEADERS) -> bytes:
    return frame(rows, headers).to_csv(index=False).encode()


def xlsx_bytes(rows, headers=HEADERS) -> bytes:
    buffer = io.BytesIO()
    frame(rows, headers).to_excel(buffer, index=False)
    return buffer.getvalue()


def upload(data: bytes, name="roster.csv") -> UploadFile:
    return UploadFile(file=io.BytesIO(data), filename=name)


def parse(rows, name="roster.csv", data=None, mapping=None):
    return importer.parse_students(
        name,
        data if data is not None else csv_bytes(rows),
        None,
        importer.ClinicalStudentsImportMapping(**(mapping or MAPPING)),
    )


def triples(row):
    return [(w.day, w.start_minute, w.end_minute, w.status) for w in row.windows]


class ParserTests(unittest.TestCase):
    def test_single_range_defaults_to_available(self):
        parsed = parse([["S1", "8:00-12:00", "", "", ""]])
        self.assertEqual(triples(parsed.rows[0]), [("Monday", 480, 720, "Available")])
        self.assertEqual(parsed.rows[0].name, "Test Name")
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

    def test_every_row_creates_a_record(self):
        parsed = parse([["", "8-9", "", "", ""], ["S9", "8-9", "", "", ""], ["S1", "8-9", "", "", ""],
                        ["s1", "8-9", "", "", ""]])
        self.assertIn("Student ID is required.", parsed.rows[0].errors)
        self.assertEqual(parsed.rows[1].errors, [])
        # Repeated Student IDs are separate placement requests; every row is a new record.
        self.assertEqual([parsed.rows[2].errors, parsed.rows[3].errors], [[], []])
        self.assertEqual([row.action for row in parsed.rows[1:]], ["create", "create", "create"])

    def test_mapping_validation(self):
        columns = list(HEADERS)
        bad = [
            importer.ClinicalStudentsImportMapping(days={"Monday": "Mon"}),
            importer.ClinicalStudentsImportMapping(student_id="Student ID", days={"Monday": "Missing"}),
            importer.ClinicalStudentsImportMapping(student_id="Mon", days={"Monday": "Mon"}),
            importer.ClinicalStudentsImportMapping(student_id="Student ID", days={"Someday": "Mon"}),
        ]
        for mapping in bad:
            with self.assertRaises(ImportFileError) as caught:
                importer.validate_mapping(mapping, columns)
            self.assertEqual(caught.exception.code, "INVALID_MAPPING")

    def test_inspect_suggests_mapping(self):
        headers = ["Student ID", "Monday", "tue", "Wed.", "Thursday", "Notes"]
        inspection = importer.inspect_file("a.csv", csv_bytes([["S1", "Ada Example", "", "", "", ""]], headers), None)
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
            ["Add 2 availability windows", "Add 1 availability window"],
        )
        cleared = importer.build_preview(parse([["S1", "", "", "", ""]]))
        self.assertEqual(cleared.rows[0].availability_change, "No availability change")
        self.assertTrue(cleared.can_apply)

    def test_invalid_rows_block_apply(self):
        preview = importer.build_preview(parse([["S1", "8-9", "", "", ""], ["S9", "bad", "", "", ""]]))
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
        self.assertEqual(preview.rows[0].availability_change, "Add 1 availability window")
        self.db.rollback()
        state = self.repo.get_state()
        self.assertEqual((state.source_data_revision, state.is_schedule_stale), (revision, stale))
        self.assertEqual(len(self.repo.student_availability(self.ada.id)), 1)
        self.assertEqual(self.repo.student_availability(self.ben.id), [])
        self.assertEqual((preview.students_to_remove, preview.placements_to_remove), (3, 1))

    def test_apply_replaces_roster_and_removes_absent_students(self):
        self.seed()
        before = self.repo.get_state().source_data_revision
        result = self.call("apply", [["s1", "8-10; T 1-4", "9-11", "", ""], ["S2", "", "", "", ""]])
        self.assertEqual((result.students_created, result.students_updated, result.availability_updated, result.windows_imported), (2, 0, 2, 3))
        self.assertEqual((result.students_removed, result.placements_removed), (3, 1))
        by_id = {s.student_id.casefold(): s for s in self.repo.list_students()}
        ada = self.repo.student_availability(by_id["s1"].id)
        self.assertEqual(
            [(w.day, w.start_minute, w.end_minute, w.status, w.source) for w in ada],
            [("Monday", 480, 600, "Available", "import"), ("Monday", 780, 960, "Tentative", "import"),
             ("Tuesday", 540, 660, "Available", "import")],
        )
        self.assertEqual(self.repo.student_availability(by_id["s2"].id), [])
        self.assertEqual(sorted(by_id), ["s1", "s2"])
        self.assertEqual(len(self.repo.list_opportunities()), 1)
        self.assertEqual(self.db.query(ClinicalPlacement).count(), 0)
        state = self.repo.get_state()
        self.assertEqual(state.source_data_revision, before + 1)
        self.assertFalse(state.is_schedule_stale)  # every placement was cleared

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
        self.assertEqual(len(caught.exception.detail["issues"]), 1)
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
        (ben,) = self.repo.list_students()
        self.assertEqual([w.status for w in self.repo.student_availability(ben.id)], ["Tentative"])


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
            (2, 0, 2, 2),
        )
        ada = next(s for s in self.repo.list_students() if s.student_id.casefold() == "s1")
        self.assertEqual((ada.name, ada.email), ("Ada E.", "new@example.test"))
        self.assertEqual(len(self.repo.student_availability(ada.id)), 2)
        self.assertEqual(self.revision(), before + 1)
        self.assertEqual(self.db.query(ClinicalPlacement).count(), 0)

    def test_roster_only_import_creates_students_without_availability_and_bumps_once(self):
        self.seed()
        before = self.revision()
        result = self.run_apply([["S1", "Ada Renamed", "", "x", "y"]], ROSTER_ONLY)
        self.assertEqual((result.availability_updated, result.windows_imported), (1, 0))
        self.assertEqual(self.repo.student_availability(self.repo.list_students()[0].id), [])
        self.assertEqual(self.revision(), before + 1)

    def test_name_column_is_required(self):
        self.seed()
        with self.assertRaises(HTTPException) as caught:
            self.run_apply([["S1", "", "", "", "9-10"]], {"student_id": "Student ID", "days": FULL_MAPPING["days"]})
        self.assertEqual(caught.exception.detail["code"], "INVALID_MAPPING")
        self.assertEqual(len(self.repo.student_availability(self.ada.id)), 1)

    def test_all_blank_mapped_days_create_a_student_without_availability(self):
        self.seed()
        self.run_apply([["S1", "Ada Example", "ada@example.test", "", ""]], FULL_MAPPING)
        self.assertEqual(self.repo.student_availability(self.repo.list_students()[0].id), [])

    def test_absent_students_are_removed_with_their_placements(self):
        self.seed()
        preview = self.run_preview([["S2", "Ben Sample", "", "8-9", ""]], FULL_MAPPING)
        self.assertEqual((preview.students_to_remove, preview.placements_to_remove), (1, 1))
        result = self.run_apply([["S2", "Ben Sample", "", "8-9", ""]], FULL_MAPPING)
        self.assertEqual((result.students_removed, result.placements_removed), (1, 1))
        self.assertEqual({s.student_id for s in self.repo.list_students()}, {"S2"})
        self.assertEqual(self.db.query(ClinicalPlacement).count(), 0)

    def test_reimporting_a_listed_student_recreates_it_and_clears_placements(self):
        self.seed()
        self.run_apply([["s1", "Ada Example", "", "8-9", ""]], FULL_MAPPING)
        self.assertEqual(self.db.query(ClinicalPlacement).count(), 0)
        self.assertEqual(len(self.repo.list_students()), 1)
        self.assertEqual(self.db.query(ClinicalPlacement).count(), 0)

    def test_row_order_has_no_effect_on_the_resulting_roster(self):
        rows = [["S1", "Ada", "", "8-9", ""], ["S2", "Ben", "", "9-10", ""], ["S1", "Ada", "", "10-11", ""]]

        def snapshot():
            return sorted(
                (s.student_id, s.name, tuple((w.day, w.start_minute) for w in self.repo.student_availability(s.id)))
                for s in self.repo.list_students()
            )

        self.seed()
        self.run_apply(rows, FULL_MAPPING)
        forward = snapshot()
        self.run_apply(list(reversed(rows)), FULL_MAPPING)
        self.assertEqual(snapshot(), forward)

    def test_preview_messages_and_no_writes(self):
        self.seed()
        before = self.revision()
        preview = self.run_preview(
            [["S1", "Ada Example", "other@example.test", "8-9; 10-11", ""], ["S2", "Ben Sample", "", "", ""]],
            FULL_MAPPING,
        )
        first, second = preview.rows
        self.assertEqual((first.action, first.student_change), ("create", "Create student"))
        self.assertEqual(first.availability_change, "Add 2 availability windows")
        self.assertEqual((second.action, second.student_change), ("create", "Create student"))
        self.assertEqual((preview.students_to_create, preview.students_to_update), (2, 0))
        self.assertEqual((preview.students_to_remove, preview.placements_to_remove), (1, 1))
        self.db.rollback()
        self.assertEqual(self.revision(), before)
        self.assertEqual(len(self.repo.list_students()), 1)

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


class TagImportTests(ClinicalTestCase):
    HEADERS = ["Student ID", "Name", "Required Tags", "Preferred Tags", "Restricted Tags", "Notes"]
    MAPPING = {
        "student_id": "Student ID", "name": "Name", "required_tags": "Required Tags",
        "preferred_tags": "Preferred Tags", "restricted_tags": "Restricted Tags",
    }

    def seed(self):
        self.ada = self.repo.create_student("S1", "Ada Example", "", require_name=True)
        self.ben = self.repo.create_student("S2", "Ben Sample", "", require_name=True)
        self.adult = self.repo.create_tag("Adult")
        self.old = self.repo.create_tag("Old")
        self.repo.set_student_tag_states(self.ada.id, {self.old.id: "required"})
        self.repo.set_student_tag_states(self.ben.id, {self.old.id: "preferred"})
        self.db.commit()

    def call(self, fn, rows, mapping, **kw):
        return asyncio.run(fn(
            file=upload(csv_bytes(rows, self.HEADERS)), mapping=json.dumps(mapping), sheet_name=None,
            db=self.db, **kw,
        ))

    def states(self, student):
        names = {t.id: t.name for t in self.repo.list_tags()}
        return {names[tag]: state for tag, state in self.repo.student_tag_states().get(student.id, {}).items()}

    def test_inspect_suggests_tag_columns(self):
        info = importer.inspect_file("r.csv", csv_bytes([], self.HEADERS), None)
        self.assertEqual(info.suggested_mapping.required_tags, "Required Tags")
        self.assertEqual(info.suggested_mapping.restricted_tags, "Restricted Tags")

    def test_preview_lists_new_and_reused_without_writing(self):
        self.seed()
        preview = self.call(clinical.preview_students_import, [["S1", "Ada Example", "adult; Virtual", "", "", ""]], self.MAPPING)
        self.assertEqual((preview.new_tags, preview.reused_tags), (["VIRTUAL"], ["ADULT"]))
        self.assertEqual(preview.rows[0].required_tags, ["ADULT", "VIRTUAL"])
        self.db.rollback()
        self.assertEqual(len(self.repo.list_tags()), 2)

    def test_apply_replaces_listed_students_and_creates_tags(self):
        self.seed()
        result = self.call(
            clinical.apply_students_import,
            [["S1", "Ada Example", "adult", "virtual", "", ""]], self.MAPPING, confirmed=True,
        )
        self.assertEqual((result.tags_created, result.tag_states_updated), (1, 1))
        (ada,) = self.repo.list_students()
        self.assertEqual(self.states(ada), {"ADULT": "required", "VIRTUAL": "preferred"})
        self.assertEqual(ada.student_id, "S1")

    def test_blank_cells_clear_listed_student(self):
        self.seed()
        self.call(clinical.apply_students_import, [["S1", "Ada Example", "", "", "", ""]], self.MAPPING, confirmed=True)
        self.assertEqual(self.states(self.repo.list_students()[0]), {})

    def test_unmapped_tag_columns_clear_listed_students(self):
        self.seed()
        self.call(
            clinical.apply_students_import, [["S1", "Ada Example", "adult", "", "", ""]],
            {"student_id": "Student ID", "name": "Name"}, confirmed=True,
        )
        self.assertEqual(self.states(self.repo.list_students()[0]), {})
        self.assertEqual(len(self.repo.list_tags()), 2)

    def test_conflicting_states_reject_import(self):
        self.seed()
        parsed = importer.parse_students(
            "r.csv", csv_bytes([["S1", "Ada Example", "Adult", "adult", "", ""]], self.HEADERS), None,
            importer.ClinicalStudentsImportMapping(**self.MAPPING),
        )
        self.assertTrue(parsed.rows[0].errors)
        with self.assertRaises(HTTPException):
            self.call(
                clinical.apply_students_import, [["S1", "Ada Example", "Adult", "adult", "", ""]], self.MAPPING, confirmed=True
            )
        self.assertEqual(self.states(self.ada), {"OLD": "required"})

    def test_column_reuse_is_rejected(self):
        with self.assertRaises(ImportFileError):
            importer.validate_mapping(
                importer.ClinicalStudentsImportMapping(
                    student_id="Student ID", required_tags="Notes", preferred_tags="Notes"
                ), self.HEADERS,
            )


if __name__ == "__main__":
    unittest.main()
