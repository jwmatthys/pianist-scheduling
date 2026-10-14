"""Synthetic-data tests for the Clinical Student Roster repository, API, and import."""

import asyncio
import io
import json
import unittest

import pandas as pd
from fastapi import HTTPException, UploadFile

from app.clinical_models import ClinicalStudentTagState, PlacementType, TagState
from app.routers import clinical
from app.services import clinical_students_importer as importer
from app.services.availability import AvailabilityWindow
from app.services.clinical_repository import ClinicalValidationError
from tests.test_clinical_foundation import ClinicalTestCase, opportunity_input

HEADERS = ["ID", "Student", "Mail", "Monday"]
MAPPING = {"student_id": "ID", "name": "Student", "email": "Mail"}


def csv_bytes(rows, headers=HEADERS) -> bytes:
    return pd.DataFrame(rows, columns=headers).to_csv(index=False).encode()


def xlsx_bytes(rows, headers=HEADERS) -> bytes:
    buffer = io.BytesIO()
    pd.DataFrame(rows, columns=headers).to_excel(buffer, index=False)
    return buffer.getvalue()


def upload(data: bytes, name="roster.csv") -> UploadFile:
    return UploadFile(file=io.BytesIO(data), filename=name)


def parse(rows, name="roster.csv", data=None, **overrides):
    return importer.parse_students(
        name, data if data is not None else csv_bytes(rows), None,
        importer.ClinicalStudentsImportMapping(**{**MAPPING, **overrides}),
    )


GOOD = [["S100", "Ada Example", "ada@example.test", "9-5"], ["S200", "Ben Sample", "", ""]]


class StudentRepositoryTests(ClinicalTestCase):
    def test_crud_and_validation(self):
        student = self.repo.create_student(" S1 ", " Pat Test ", " pat@example.test ", require_name=True)
        self.assertEqual((student.student_id, student.name, student.email), ("S1", "Pat Test", "pat@example.test"))
        updated = self.repo.update_student(student.id, "S9", "Pat Renamed", "")
        self.assertEqual((updated.student_id, updated.name, updated.email), ("S9", "Pat Renamed", ""))
        self.assertEqual([s.student_id for s in self.repo.list_students()], ["S9"])
        self.repo.delete_student(student.id)
        self.assertEqual(self.repo.list_students(), [])

    def test_required_fields_and_duplicate_ids(self):
        with self.assertRaises(ClinicalValidationError) as caught:
            self.repo.create_student("S1", "  ", require_name=True)
        self.assertEqual(caught.exception.code, "INVALID_STUDENT_NAME")
        with self.assertRaises(ClinicalValidationError):
            self.repo.create_student(" ", "Name", require_name=True)
        first = self.repo.create_student("abc", "One", require_name=True)
        second = self.repo.create_student("S2", "Two", require_name=True)
        with self.assertRaises(ClinicalValidationError) as caught:
            self.repo.create_student("ABC", "Dup", require_name=True)
        self.assertEqual(caught.exception.code, "DUPLICATE_STUDENT_ID")
        with self.assertRaises(ClinicalValidationError):
            self.repo.update_student(second.id, "Abc", "Two")
        self.repo.update_student(first.id, "ABC", "One Edited")  # own ID is not a duplicate
        with self.assertRaises(ClinicalValidationError) as caught:
            self.repo.update_student(first.id, "S1", " ")
        self.assertEqual(caught.exception.code, "INVALID_STUDENT_NAME")

    def test_student_edits_bump_revision_and_flag_stale_only_with_placements(self):
        site = self.repo.create_opportunity(opportunity_input())
        student = self.repo.create_student("S1", "Pat", require_name=True)
        before = self.repo.get_state().source_data_revision
        self.repo.update_student(student.id, "S1", "Pat Two")
        self.assertEqual(self.repo.get_state().source_data_revision, before + 1)
        self.assertFalse(self.repo.get_state().is_schedule_stale)
        self.repo.set_placement(student.id, PlacementType.MANUAL, site.id, 480)
        self.repo.update_student(student.id, "S1", "Pat Three")
        self.assertTrue(self.repo.get_state().is_schedule_stale)

    def test_tag_state_transitions_and_neutral_is_absent(self):
        student = self.repo.create_student("S1", "Pat", require_name=True)
        tag = self.repo.create_tag("kids")
        for state in (TagState.REQUIRED, TagState.PREFERRED, TagState.RESTRICTED):
            self.repo.set_student_tag_states(student.id, {tag.id: state})
            self.assertEqual(self.repo.get_tag_state(student.id, tag.id), state)
            self.assertEqual(self.repo.student_tag_states()[student.id][tag.id], state)
        self.repo.set_student_tag_states(student.id, {tag.id: TagState.NEUTRAL})
        self.assertEqual(self.db.query(ClinicalStudentTagState).count(), 0)
        self.assertEqual(self.repo.student_tag_states(), {})

    def test_bulk_tag_states_replace_and_bump_once(self):
        student = self.repo.create_student("S1", "Pat", require_name=True)
        a, b = self.repo.create_tag("a"), self.repo.create_tag("b")
        self.repo.set_student_tag_states(student.id, {a.id: "required", b.id: "restricted"})
        before = self.repo.get_state().source_data_revision
        self.repo.set_student_tag_states(student.id, {b.id: "preferred"})
        self.assertEqual(self.repo.get_state().source_data_revision, before + 1)
        self.assertEqual(self.repo.student_tag_states()[student.id], {b.id: TagState.PREFERRED})
        with self.assertRaises(ClinicalValidationError) as caught:
            self.repo.set_student_tag_states(student.id, {b.id: "bogus"})
        self.assertEqual(caught.exception.code, "INVALID_TAG_STATE")
        with self.assertRaises(ClinicalValidationError):
            self.repo.set_student_tag_states(student.id, {"missing": "required"})

    def test_delete_student_clears_placement_and_tag_states(self):
        site = self.repo.create_opportunity(opportunity_input())
        student = self.repo.create_student("S1", "Pat", require_name=True)
        tag = self.repo.create_tag("kids")
        self.repo.set_student_tag_states(student.id, {tag.id: "required"})
        self.repo.set_placement(student.id, PlacementType.AUTOMATIC, site.id, 480)
        self.repo.delete_student(student.id)
        self.assertEqual(self.db.query(ClinicalStudentTagState).count(), 0)
        self.assertIsNone(self.repo.get_placement(student.id))

    def test_availability_counts(self):
        student = self.repo.create_student(
            "S1", "Pat", availability=[AvailabilityWindow("Monday", 60, 120, "Available")], require_name=True
        )
        self.repo.create_student("S2", "Sam", require_name=True)
        self.assertEqual(self.repo.student_availability_counts(), {student.id: 1})


class StudentApiTests(ClinicalTestCase):
    def test_create_update_tag_states_delete(self):
        tag = self.repo.create_tag("kids")
        self.db.commit()
        created = clinical.create_student(clinical.schemas.ClinicalStudentFields(student_id="S1", name="Pat"), self.db)
        self.assertEqual(created.availability_window_count, 0)
        with self.assertRaises(HTTPException) as caught:
            clinical.create_student(clinical.schemas.ClinicalStudentFields(student_id="s1", name="Dup"), self.db)
        self.assertEqual(caught.exception.status_code, 409)
        with self.assertRaises(HTTPException) as caught:
            clinical.create_student(clinical.schemas.ClinicalStudentFields(student_id="S2", name=" "), self.db)
        self.assertEqual(caught.exception.status_code, 422)
        updated = clinical.set_student_tag_states(
            created.id, clinical.schemas.ClinicalStudentTagStatesIn(tag_states={tag.id: "required"}), self.db
        )
        self.assertEqual(updated.tag_states, {tag.id: "required"})
        listed = clinical.list_students(self.db)
        self.assertEqual([s.student_id for s in listed], ["S1"])
        clinical.delete_student(created.id, self.db)
        self.assertEqual(clinical.list_students(self.db), [])
        with self.assertRaises(HTTPException) as caught:
            clinical.delete_student(created.id, self.db)
        self.assertEqual(caught.exception.status_code, 404)


class StudentImportParsingTests(unittest.TestCase):
    def test_csv_and_ignored_columns(self):
        rows = parse(GOOD).rows
        self.assertEqual([(r.student_id, r.name, r.email, r.errors) for r in rows],
                         [("S100", "Ada Example", "ada@example.test", []), ("S200", "Ben Sample", "", [])])
        self.assertEqual(rows[0].row_number, 2)

    def test_xlsx_matches_csv(self):
        from_xlsx = parse(None, name="roster.xlsx", data=xlsx_bytes(GOOD)).rows
        self.assertEqual([(r.student_id, r.name) for r in from_xlsx], [("S100", "Ada Example"), ("S200", "Ben Sample")])

    def test_numeric_ids_are_not_floats(self):
        rows = parse(None, name="r.xlsx", data=xlsx_bytes([[1001, "Num", "", ""]])).rows
        self.assertEqual(rows[0].student_id, "1001")

    def test_required_fields_and_blank_rows(self):
        rows = parse([["", "No Id", "", ""], ["S1", "", "", ""], ["", "", "", ""]]).rows
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0].errors, ["Student ID is required."])
        self.assertEqual(rows[1].errors, ["Name is required."])

    def test_duplicate_ids_case_insensitive(self):
        rows = parse([["S1", "A", "", ""], ["s1", "B", "", ""], ["S2", "C", "", ""]]).rows
        self.assertTrue(any("duplicated" in e for e in rows[0].errors))
        self.assertTrue(any("row 2" in e for e in rows[1].errors))
        self.assertEqual(rows[2].errors, [])

    def test_mapping_errors_and_suggestions(self):
        with self.assertRaises(importer.ClinicalImportError) as caught:
            importer.parse_students(
                "r.csv", csv_bytes(GOOD), None, importer.ClinicalStudentsImportMapping(name="Student")
            )
        self.assertEqual(caught.exception.code, "INVALID_MAPPING")
        with self.assertRaises(importer.ClinicalImportError):
            parse(GOOD, name="ID")
        inspection = importer.inspect_file("r.csv", csv_bytes(GOOD, ["Student ID", "Name", "Email", "x"]), None)
        self.assertEqual(inspection.suggested_mapping, {"student_id": "Student ID", "name": "Name", "email": "Email"})

    def test_preview(self):
        preview = importer.build_preview(parse(GOOD + [["", "Bad", "", ""]]), 3, 4)
        self.assertEqual((preview.total_rows, preview.valid_rows, preview.invalid_rows), (3, 2, 1))
        self.assertEqual((preview.existing_student_count, preview.existing_placement_count), (3, 4))
        self.assertFalse(preview.can_apply)
        self.assertTrue(importer.build_preview(parse(GOOD), 0, 0).can_apply)


class StudentReplacementTests(ClinicalTestCase):
    def seed(self):
        tag = self.repo.create_tag("kids")
        site = self.repo.create_opportunity(opportunity_input(site_name="Keep Site"), tag_ids=[tag.id])
        old = self.repo.create_student(
            "OLD1", "Old Student", availability=[AvailabilityWindow("Monday", 60, 120, "Available")], require_name=True
        )
        self.repo.set_student_tag_states(old.id, {tag.id: "required"})
        self.repo.set_placement(old.id, PlacementType.MANUAL, site.id, 480)
        self.repo.set_placement(self.repo.create_student("OLD2", "Unplaced").id, PlacementType.MANUALLY_UNPLACED)
        self.db.commit()

    def apply(self, rows=GOOD, confirmed=True):
        return asyncio.run(
            clinical.apply_students_import(
                file=upload(csv_bytes(rows)), mapping=json.dumps(MAPPING), sheet_name=None,
                confirmed=confirmed, db=self.db,
            )
        )

    def test_full_replacement(self):
        self.seed()
        before = self.repo.get_state().source_data_revision
        result = self.apply()
        self.assertEqual((result.students_removed, result.placements_removed, result.students_created), (2, 2, 2))
        self.assertEqual([s.student_id for s in self.repo.list_students()], ["S100", "S200"])
        self.assertEqual(self.repo.student_tag_states(), {})
        self.assertEqual(self.repo.student_availability_counts(), {})
        self.assertEqual(self.db.query(clinical.ClinicalPlacement).count(), 0)
        self.assertEqual([o.site_name for o in self.repo.list_opportunities()], ["Keep Site"])
        self.assertEqual([t.name for t in self.repo.list_tags()], ["KIDS"])
        self.assertEqual([t.name for t in self.repo.list_opportunities()[0].tags], ["KIDS"])
        state = self.repo.get_state()
        self.assertEqual(state.source_data_revision, before + 1)
        self.assertEqual(state.schedule_revision, 0)
        self.assertFalse(state.is_schedule_stale)

    def test_requires_confirmation(self):
        self.seed()
        with self.assertRaises(HTTPException) as caught:
            self.apply(confirmed=False)
        self.assertEqual(caught.exception.detail["code"], "CONFIRMATION_REQUIRED")
        self.assertEqual(len(self.repo.list_students()), 2)

    def test_invalid_rows_reject_entire_import(self):
        self.seed()
        with self.assertRaises(HTTPException) as caught:
            self.apply(GOOD + [["S100", "Dup", "", ""]])
        self.assertEqual(caught.exception.detail["code"], "VALIDATION_FAILED")
        self.assertEqual({s.student_id for s in self.repo.list_students()}, {"OLD1", "OLD2"})
        self.assertEqual(self.db.query(clinical.ClinicalPlacement).count(), 2)

    def test_empty_file_rejected(self):
        self.seed()
        with self.assertRaises(HTTPException) as caught:
            self.apply(rows=[])
        self.assertEqual(caught.exception.detail["code"], "EMPTY_IMPORT")
        self.assertEqual(len(self.repo.list_students()), 2)

    def test_preview_writes_nothing(self):
        self.seed()
        preview = asyncio.run(
            clinical.preview_students_import(
                file=upload(csv_bytes(GOOD)), mapping=json.dumps(MAPPING), sheet_name=None, db=self.db
            )
        )
        self.assertTrue(preview.can_apply)
        self.assertEqual((preview.existing_student_count, preview.existing_placement_count), (2, 2))
        self.db.rollback()
        self.assertEqual({s.student_id for s in self.repo.list_students()}, {"OLD1", "OLD2"})


if __name__ == "__main__":
    unittest.main()
