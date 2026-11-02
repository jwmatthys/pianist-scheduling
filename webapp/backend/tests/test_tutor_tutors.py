"""Tutor Roster: CRUD, tags, availability, import, revisions, and Clinical isolation.

All data is synthetic.
"""

import asyncio
import io
import json
import unittest

import pandas as pd
from fastapi import HTTPException, UploadFile

from app import tutor_schemas as schemas
from app.tutor_models import TutorPlacementType, TutorTagState
from app.routers import tutor_scheduling as tutor_api
from app.services import tutor_tutor_importer as importer
from app.services.clinical_repository import ClinicalRepository
from app.services.tutor_repository import TutorInput, TutorValidationError
from tests.test_tutor_foundation import TutorTestCase, window

HEADERS = ["Tutor Name", "Tutor Email", "Room", "Lesson Length", "Capacity",
           "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Tags", "Notes"]
MAPPING = {
    "tutor_name": "Tutor Name", "tutor_email": "Tutor Email", "room": "Room",
    "lesson_length": "Lesson Length", "capacity": "Capacity",
    "days": {"Monday": "Mon", "Tuesday": "Tue", "Wednesday": "Wed", "Thursday": "Thu",
             "Friday": "Fri", "Saturday": "Sat", "Sunday": "Sun"},
    "tags": "Tags",
}
GOOD_ROW = ["Synthetic Teacher", "teacher@example.test", "Room 101", 45, 2,
            "9:00 AM-12:00 PM; T 1:00 PM-3:00 PM", "", "", "", "", "", "", "violin; Strings; violin;", ""]
BLANK_DAYS = ["", "", "", "", "", "", ""]


def row_with(name="T", email="", room="", length="", capacity="", monday="", tags=("", "")):
    return [name, email, room, length, capacity, monday, *BLANK_DAYS[1:], *tags]


def csv_bytes(rows, headers=HEADERS) -> bytes:
    return pd.DataFrame(rows, columns=headers).to_csv(index=False).encode()


def xlsx_bytes(rows, headers=HEADERS) -> bytes:
    buffer = io.BytesIO()
    pd.DataFrame(rows, columns=headers).to_excel(buffer, index=False)
    return buffer.getvalue()


def upload(data: bytes, name="tutors.csv") -> UploadFile:
    return UploadFile(file=io.BytesIO(data), filename=name)


def mapping(**overrides):
    return importer.TutorTutorImportMapping(**{**MAPPING, **overrides})


def parse(rows, **overrides):
    return importer.parse_tutors("tutors.csv", csv_bytes(rows), None, mapping(**overrides))


class TutorApiTests(TutorTestCase):
    def payload(self, **overrides):
        return schemas.TutorTutorIn(**{"name": "Synthetic Teacher", **overrides})

    def test_create_update_delete_and_defaults(self):
        created = tutor_api.create_tutor(self.payload(), db=self.db)
        self.assertEqual((created.lesson_length_minutes, created.capacity), (60, 1))
        updated = tutor_api.update_tutor(
            created.id, self.payload(name="Renamed", room="B2", lesson_length_minutes=30, capacity=3), db=self.db
        )
        self.assertEqual((updated.name, updated.room, updated.lesson_length_minutes, updated.capacity),
                         ("Renamed", "B2", 30, 3))
        self.assertEqual([i.name for i in tutor_api.list_tutors(db=self.db)], ["Renamed"])
        tutor_api.delete_tutor(created.id, db=self.db)
        self.assertEqual(tutor_api.list_tutors(db=self.db), [])

    def test_validation_errors(self):
        for overrides, code in (
            ({"name": "  "}, "INVALID_TUTOR_NAME"),
            ({"lesson_length_minutes": 0}, "INVALID_LESSON_LENGTH"),
            ({"capacity": 0}, "INVALID_CAPACITY"),
        ):
            with self.assertRaises(HTTPException) as caught:
                tutor_api.create_tutor(self.payload(**overrides), db=self.db)
            self.assertEqual((caught.exception.status_code, caught.exception.detail["code"]), (422, code))
        self.assertEqual(self.repo.list_tutors(), [])

    def test_missing_tutor_is_404(self):
        with self.assertRaises(HTTPException) as caught:
            tutor_api.update_tutor("missing", self.payload(), db=self.db)
        self.assertEqual(caught.exception.status_code, 404)

    def test_delete_preserves_assignments_and_reports_them_for_review(self):
        first = self.repo.create_tutor("First")
        second = self.repo.create_tutor("Second")
        a, b = self.repo.create_student("S1", "Pat Test"), self.repo.create_student("S2", "Sam Test")
        self.repo.set_placement(a.id, TutorPlacementType.MANUAL, first.id, 540)
        self.repo.set_placement(b.id, TutorPlacementType.MANUAL, second.id, 540)
        result = tutor_api.delete_tutor(first.id, db=self.db)
        self.assertEqual((result.removed_placement_count, result.placements_preserved), (0, 1))
        self.assertIsNotNone(self.repo.get_placement(a.id))
        self.assertIsNotNone(self.repo.get_placement(b.id))

    def test_update_with_tags_is_one_revision_and_none_keeps_tags(self):
        tag = self.repo.create_tag("strings")
        created = tutor_api.create_tutor(self.payload(tag_ids=[tag.id]), db=self.db)
        self.assertEqual([t.name for t in created.tags], ["STRINGS"])
        before = self.revision()
        tutor_api.update_tutor(created.id, self.payload(name="Again"), db=self.db)
        self.assertEqual(self.revision(), before + 1)
        self.assertEqual([t.name for t in tutor_api.list_tutors(db=self.db)[0].tags], ["STRINGS"])
        tutor_api.update_tutor(created.id, self.payload(tag_ids=[]), db=self.db)
        self.assertEqual(tutor_api.list_tutors(db=self.db)[0].tags, [])

    def test_tag_assignment_endpoint(self):
        created = tutor_api.create_tutor(self.payload(), db=self.db)
        tag = self.repo.create_tag("brass")
        assigned = tutor_api.set_tutor_tags(created.id, schemas.TutorTutorTagsIn(tag_ids=[tag.id]), db=self.db)
        self.assertEqual([t.name for t in assigned.tags], ["BRASS"])
        with self.assertRaises(HTTPException) as caught:
            tutor_api.set_tutor_tags(created.id, schemas.TutorTutorTagsIn(tag_ids=["nope"]), db=self.db)
        self.assertEqual(caught.exception.status_code, 404)


class TagApiTests(TutorTestCase):
    def test_create_normalizes_and_rejects_duplicates(self):
        tag = tutor_api.create_tag(schemas.TutorTagIn(name="  mental health "), db=self.db)
        self.assertEqual(tag.name, "MENTAL HEALTH")
        self.assertTrue(tag.is_unused)
        with self.assertRaises(HTTPException) as caught:
            tutor_api.create_tag(schemas.TutorTagIn(name="Mental Health"), db=self.db)
        self.assertEqual(caught.exception.status_code, 409)

    def test_rename_and_duplicate_rename(self):
        a = tutor_api.create_tag(schemas.TutorTagIn(name="a"), db=self.db)
        tutor_api.create_tag(schemas.TutorTagIn(name="b"), db=self.db)
        self.assertEqual(tutor_api.rename_tag(a.id, schemas.TutorTagIn(name=" c "), db=self.db).name, "C")
        with self.assertRaises(HTTPException):
            tutor_api.rename_tag(a.id, schemas.TutorTagIn(name="B"), db=self.db)

    def test_usage_counts_and_delete_reports_affected(self):
        tag = self.repo.create_tag("x")
        self.repo.create_tag("unused")
        self.repo.create_tutor("Synthetic Teacher", tag_ids=[tag.id])
        listed = {t.name: t for t in tutor_api.list_tags(db=self.db)}
        self.assertEqual(listed["X"].tutor_count, 1)
        self.assertFalse(listed["X"].is_unused)
        self.assertTrue(listed["UNUSED"].is_unused)
        result = tutor_api.delete_tag(tag.id, db=self.db)
        self.assertEqual(result.affected_tutor_count, 1)
        self.assertEqual(self.repo.list_tutors()[0].tags, [])
        self.assertEqual([t.name for t in self.repo.list_tags()], ["UNUSED"])

    def test_student_use_counts_as_in_use(self):
        tag = self.repo.create_tag("x")
        student = self.repo.create_student("S1", "Pat Test")
        self.repo.set_tag_state(student.id, tag.id, TutorTagState.REQUIRED)
        listed = tutor_api.list_tags(db=self.db)[0]
        self.assertEqual(listed.student_count, 1)
        self.assertFalse(listed.is_unused)

    def test_each_tag_change_is_one_source_change(self):
        before = self.revision()
        tag = tutor_api.create_tag(schemas.TutorTagIn(name="x"), db=self.db)
        tutor_api.rename_tag(tag.id, schemas.TutorTagIn(name="y"), db=self.db)
        tutor_api.delete_tag(tag.id, db=self.db)
        self.assertEqual(self.revision(), before + 3)


class AvailabilityApiTests(TutorTestCase):
    def test_edit_persists_and_replaces(self):
        tutor = self.repo.create_tutor("Synthetic Teacher")
        io_in = schemas.TutorTutorAvailabilityIO(windows=[
            schemas.TutorAvailabilityWindowIO(day="Mon", start_minute=540, end_minute=600, status="Available"),
            schemas.TutorAvailabilityWindowIO(day="Monday", start_minute=600, end_minute=660, status="Tentative"),
            schemas.TutorAvailabilityWindowIO(day="Friday", start_minute=720, end_minute=780, status="Unavailable"),
        ])
        saved = tutor_api.set_tutor_availability(tutor.id, io_in, db=self.db)
        self.assertEqual(
            sorted((w.day, w.start_minute, w.status) for w in saved.windows),
            [("Friday", 720, "Unavailable"), ("Monday", 540, "Available"), ("Monday", 600, "Tentative")],
        )
        self.assertEqual(len(tutor_api.get_tutor_availability(tutor.id, db=self.db).windows), 3)
        self.assertEqual(tutor_api.list_tutors(db=self.db)[0].availability_window_count, 3)
        listed = tutor_api.list_tutors(db=self.db)[0]
        self.assertEqual((listed.available_minutes, listed.tentative_minutes), (60, 60))
        tutor_api.set_tutor_availability(tutor.id, schemas.TutorTutorAvailabilityIO(windows=[]), db=self.db)
        self.assertEqual(tutor_api.get_tutor_availability(tutor.id, db=self.db).windows, [])

    def test_invalid_window_rejected_without_partial_write(self):
        tutor = self.repo.create_tutor("Synthetic Teacher", availability=[window()])
        self.db.commit()
        bad = schemas.TutorTutorAvailabilityIO(windows=[
            schemas.TutorAvailabilityWindowIO(day="Monday", start_minute=600, end_minute=540, status="Available"),
        ])
        with self.assertRaises(HTTPException) as caught:
            tutor_api.set_tutor_availability(tutor.id, bad, db=self.db)
        self.assertEqual(caught.exception.status_code, 422)
        self.assertEqual(len(self.repo.tutor_availability(tutor.id)), 1)

    def test_availability_edit_is_one_revision(self):
        tutor = self.repo.create_tutor("Synthetic Teacher")
        before = self.revision()
        tutor_api.set_tutor_availability(
            tutor.id,
            schemas.TutorTutorAvailabilityIO(windows=[
                schemas.TutorAvailabilityWindowIO(day="Monday", start_minute=540, end_minute=600, status="Available"),
            ]),
            db=self.db,
        )
        self.assertEqual(self.revision(), before + 1)


class RevisionTests(TutorTestCase):
    def test_tutor_changes_bump_once_and_mark_existing_schedule_stale(self):
        student = self.repo.create_student("S1", "Pat Test")
        tutor = self.repo.create_tutor("Synthetic Teacher")
        self.repo.set_placement(student.id, TutorPlacementType.MANUAL, tutor.id, 540)
        self.assertFalse(self.repo.get_state().is_schedule_stale)
        before = self.revision()
        self.repo.update_tutor(tutor.id, "Synthetic Teacher", room="B")
        self.assertEqual(self.revision(), before + 1)
        self.assertTrue(self.repo.get_state().is_schedule_stale)

    def test_changes_without_a_schedule_never_mark_stale(self):
        tutor = self.repo.create_tutor("Synthetic Teacher")
        self.repo.set_tutor_availability(tutor.id, [window()])
        self.repo.update_tutor(tutor.id, "Synthetic Teacher", room="B")
        self.assertFalse(self.repo.get_state().is_schedule_stale)


class ImportParsingTests(unittest.TestCase):
    def test_csv_fields_availability_and_tags(self):
        row = parse([GOOD_ROW]).rows[0]
        self.assertEqual(row.errors, [])
        self.assertEqual(
            (row.tutor_name, row.tutor_email, row.room, row.lesson_length_minutes, row.capacity),
            ("Synthetic Teacher", "teacher@example.test", "Room 101", 45, 2),
        )
        self.assertEqual(row.tags, ["VIOLIN", "STRINGS"])
        self.assertEqual(
            [(w.day, w.start_minute, w.end_minute, w.status, w.source) for w in row.windows],
            [("Monday", 540, 720, "Available", "import"), ("Monday", 780, 900, "Tentative", "import")],
        )

    def test_xlsx_is_supported(self):
        parsed = importer.parse_tutors("i.xlsx", xlsx_bytes([GOOD_ROW]), None, mapping())
        self.assertEqual(parsed.rows[0].tutor_name, "Synthetic Teacher")

    def test_defaults_when_blank_or_unmapped(self):
        row = parse([row_with(name="Synthetic Teacher")]).rows[0]
        self.assertEqual((row.lesson_length_minutes, row.capacity, row.errors), (60, 1, []))
        unmapped = parse([GOOD_ROW], lesson_length=None, capacity=None).rows[0]
        self.assertEqual((unmapped.lesson_length_minutes, unmapped.capacity), (60, 1))

    def test_validation_errors_are_reported_per_row(self):
        errors = parse([row_with(name="", length="abc", capacity=0, monday="nonsense")]).rows[0].errors
        self.assertTrue(any("Tutor Name is required" in e for e in errors))
        self.assertTrue(any("Lesson Length 'abc'" in e for e in errors))
        self.assertTrue(any("Capacity must be greater than 0" in e for e in errors))
        self.assertTrue(any("Monday" in e for e in errors))
        self.assertIn("Lesson Length must be greater than 0.", parse([row_with(length=-5)]).rows[0].errors)

    def test_conflicting_availability_is_an_error(self):
        row = parse([row_with(monday="9:00 AM-11:00 AM; T 10:00 AM-12:00 PM")]).rows[0]
        self.assertTrue(any("overlaps" in e for e in row.errors))

    def test_unmarked_reversed_availability_ranges_assume_afternoon_end(self):
        for value, expected in (
            ("9:00-3:00", (540, 900)),
            ("8:00-1:00", (480, 780)),
            ("10:30-2:15", (630, 855)),
        ):
            with self.subTest(value=value):
                row = parse([row_with(monday=value)]).rows[0]
                self.assertEqual([(window.start_minute, window.end_minute) for window in row.windows], [expected])
                self.assertEqual(row.errors, [])

    def test_explicit_meridiem_availability_is_not_adjusted(self):
        row = parse([row_with(monday="9:00 AM-3:00 AM")]).rows[0]
        self.assertTrue(any("ends at or before" in error for error in row.errors), row.errors)

    def test_same_status_overlaps_merge(self):
        row = parse([row_with(monday="9:00 AM-11:00 AM; 10:00 AM-12:00 PM")]).rows[0]
        self.assertEqual([(w.start_minute, w.end_minute) for w in row.windows], [(540, 720)])

    def test_duplicate_name_warns(self):
        rows = parse([GOOD_ROW, GOOD_ROW]).rows
        self.assertTrue(any("more than once" in w for w in rows[1].warnings))

    def test_mapping_validation(self):
        with self.assertRaises(importer.TutorTutorImportError) as caught:
            parse([GOOD_ROW], tutor_name=None)
        self.assertEqual(caught.exception.code, "INVALID_MAPPING")
        with self.assertRaises(importer.TutorTutorImportError):
            parse([GOOD_ROW], room="Tutor Name")
        with self.assertRaises(importer.TutorTutorImportError):
            parse([GOOD_ROW], capacity="Missing Column")

    def test_empty_file_rejected(self):
        with self.assertRaises(importer.TutorTutorImportError) as caught:
            parse([])
        self.assertEqual(caught.exception.code, "NO_ROWS")

    def test_inspect_suggests_mapping_including_days(self):
        suggested = importer.inspect_file("i.csv", csv_bytes([GOOD_ROW]), None).suggested_mapping
        self.assertEqual(suggested.tutor_name, "Tutor Name")
        self.assertEqual(suggested.lesson_length, "Lesson Length")
        self.assertEqual(suggested.days["Monday"], "Mon")
        self.assertEqual(suggested.days["Sunday"], "Sun")


class ImportFlowTests(TutorTestCase):
    def seed(self):
        keep = self.repo.create_tag("violin")
        self.repo.create_tag("orphan")
        old = self.repo.create_tutor("Old Teacher", tag_ids=[keep.id], availability=[window()])
        student = self.repo.create_student("S1", "Pat Test")
        self.repo.set_placement(student.id, TutorPlacementType.MANUAL, old.id, 540)
        self.db.commit()
        return student

    def call(self, endpoint, rows=(GOOD_ROW,), **kwargs):
        return asyncio.run(endpoint(
            file=upload(csv_bytes(list(rows))), mapping=json.dumps(MAPPING), sheet_name=None, db=self.db, **kwargs
        ))

    def test_preview_writes_nothing_and_reports_impact(self):
        self.seed()
        before = (self.revision(), len(self.repo.list_tutors()), len(self.repo.list_tags()))
        preview = self.call(tutor_api.preview_tutor_import)
        self.assertTrue(preview.can_apply)
        self.assertEqual((preview.existing_tutor_count, preview.existing_placement_count), (1, 1))
        self.assertEqual(preview.new_tags, ["STRINGS"])
        self.assertEqual(preview.reused_tags, ["VIOLIN"])
        self.assertEqual(preview.rows[0].availability, ["Mon 9:00 AM-12:00 PM", "Tentative Mon 1:00 PM-3:00 PM"])
        self.assertEqual(before, (self.revision(), len(self.repo.list_tutors()), len(self.repo.list_tags())))

    def test_preview_flags_invalid_rows(self):
        preview = self.call(tutor_api.preview_tutor_import, rows=[GOOD_ROW, row_with(name="", room="Room X")])
        self.assertEqual((preview.valid_rows, preview.invalid_rows, preview.can_apply), (1, 1, False))

    def test_apply_requires_confirmation(self):
        self.seed()
        with self.assertRaises(HTTPException) as caught:
            self.call(tutor_api.apply_tutor_import, confirmed=False)
        self.assertEqual(caught.exception.detail["code"], "CONFIRMATION_REQUIRED")
        self.assertEqual([i.name for i in self.repo.list_tutors()], ["Old Teacher"])

    def test_apply_replaces_tutors_preserving_tags_and_students(self):
        student = self.seed()
        before = self.revision()
        result = self.call(tutor_api.apply_tutor_import, confirmed=True)
        self.assertEqual(
            (
                result.tutors_removed, result.placements_removed, result.placements_preserved,
                result.tutors_created, result.tags_created,
            ),
            (1, 0, 1, 1, 1),
        )
        tutors = self.repo.list_tutors()
        self.assertEqual([i.name for i in tutors], ["Synthetic Teacher"])
        self.assertEqual({t.name for t in self.repo.list_tags()}, {"VIOLIN", "ORPHAN", "STRINGS"})
        self.assertEqual({t.name for t in tutors[0].tags}, {"VIOLIN", "STRINGS"})
        self.assertEqual([s.name for s in self.repo.list_students()], ["Pat Test"])
        windows = self.repo.tutor_availability(tutors[0].id)
        self.assertEqual((len(windows), windows[0].source), (2, "import"))
        self.assertIsNotNone(self.repo.get_placement(student.id))
        state = self.repo.get_state()
        self.assertEqual(state.source_data_revision, before + 1)
        self.assertTrue(state.is_schedule_stale)

    def test_apply_reparses_and_rejects_invalid_rows_atomically(self):
        student = self.seed()
        before = self.revision()
        with self.assertRaises(HTTPException) as caught:
            self.call(tutor_api.apply_tutor_import, rows=[GOOD_ROW, row_with(name="", room="Room X")], confirmed=True)
        self.assertEqual(caught.exception.detail["code"], "VALIDATION_FAILED")
        self.assertEqual([i.name for i in self.repo.list_tutors()], ["Old Teacher"])
        self.assertEqual(self.revision(), before)
        self.assertIsNotNone(self.repo.get_placement(student.id))

    def test_inspect_endpoint(self):
        inspection = asyncio.run(tutor_api.inspect_tutor_import(file=upload(csv_bytes([GOOD_ROW])), sheet_name=None))
        self.assertIn("Tutor Name", inspection.columns)

    def test_invalid_mapping_json_is_400(self):
        with self.assertRaises(HTTPException) as caught:
            asyncio.run(tutor_api.preview_tutor_import(
                file=upload(csv_bytes([GOOD_ROW])), mapping="not json", sheet_name=None, db=self.db
            ))
        self.assertEqual(caught.exception.status_code, 400)

    def test_replace_tutors_repository_contract(self):
        self.seed()
        result = self.repo.replace_tutors([
            (TutorInput(name="A"), [window()], ["new tag", "NEW TAG"]),
            (TutorInput(name="B", lesson_length_minutes=30, capacity=2), [], []),
        ])
        self.assertEqual((result.tutors_created, result.tags_created), (2, 1))
        with self.assertRaises(TutorValidationError):
            self.repo.replace_tutors([(TutorInput(name=""), [], [])])


class ClinicalIsolationTests(TutorTestCase):
    def test_import_and_tags_never_touch_clinical(self):
        clinical = ClinicalRepository(self.db)
        clinical.create_tag("violin")
        clinical.create_student(student_id="S1", name="Synthetic Teacher", email="")
        self.db.commit()
        clinical_revision = clinical.get_state().source_data_revision
        asyncio.run(tutor_api.apply_tutor_import(
            file=upload(csv_bytes([GOOD_ROW])), mapping=json.dumps(MAPPING), sheet_name=None,
            confirmed=True, db=self.db,
        ))
        self.assertEqual([t.name for t in clinical.list_tags()], ["VIOLIN"])
        self.assertEqual([t.name for t in self.repo.list_tags()], ["STRINGS", "VIOLIN"])
        self.assertEqual(len(clinical.list_students()), 1)
        self.assertEqual(clinical.get_state().source_data_revision, clinical_revision)
        self.assertEqual(self.repo.list_students(), [])


if __name__ == "__main__":
    unittest.main()
