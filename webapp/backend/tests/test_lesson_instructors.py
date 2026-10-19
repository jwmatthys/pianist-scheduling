"""Instructor Roster: CRUD, tags, availability, import, revisions, and Clinical isolation.

All data is synthetic.
"""

import asyncio
import io
import json
import unittest

import pandas as pd
from fastapi import HTTPException, UploadFile

from app import lesson_schemas as schemas
from app.lesson_models import LessonPlacementType, LessonTagState
from app.routers import lesson
from app.services import lesson_instructor_importer as importer
from app.services.clinical_repository import ClinicalRepository
from app.services.lesson_repository import InstructorInput, LessonValidationError
from tests.test_lesson_foundation import LessonTestCase, window

HEADERS = ["Instructor Name", "Instructor Email", "Room", "Lesson Length", "Capacity",
           "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Tag 1", "Tag 2"]
MAPPING = {
    "instructor_name": "Instructor Name", "instructor_email": "Instructor Email", "room": "Room",
    "lesson_length": "Lesson Length", "capacity": "Capacity",
    "days": {"Monday": "Mon", "Tuesday": "Tue", "Wednesday": "Wed", "Thursday": "Thu",
             "Friday": "Fri", "Saturday": "Sat", "Sunday": "Sun"},
}
GOOD_ROW = ["Synthetic Teacher", "teacher@example.test", "Room 101", 45, 2,
            "9:00 AM-12:00 PM; T 1:00 PM-3:00 PM", "", "", "", "", "", "", "violin", "Strings"]
BLANK_DAYS = ["", "", "", "", "", "", ""]


def row_with(name="T", email="", room="", length="", capacity="", monday="", tags=("", "")):
    return [name, email, room, length, capacity, monday, *BLANK_DAYS[1:], *tags]


def csv_bytes(rows, headers=HEADERS) -> bytes:
    return pd.DataFrame(rows, columns=headers).to_csv(index=False).encode()


def xlsx_bytes(rows, headers=HEADERS) -> bytes:
    buffer = io.BytesIO()
    pd.DataFrame(rows, columns=headers).to_excel(buffer, index=False)
    return buffer.getvalue()


def upload(data: bytes, name="instructors.csv") -> UploadFile:
    return UploadFile(file=io.BytesIO(data), filename=name)


def mapping(**overrides):
    return importer.LessonInstructorImportMapping(**{**MAPPING, **overrides})


def parse(rows, **overrides):
    return importer.parse_instructors("instructors.csv", csv_bytes(rows), None, mapping(**overrides))


class InstructorApiTests(LessonTestCase):
    def payload(self, **overrides):
        return schemas.LessonInstructorIn(**{"name": "Synthetic Teacher", **overrides})

    def test_create_update_delete_and_defaults(self):
        created = lesson.create_instructor(self.payload(), db=self.db)
        self.assertEqual((created.lesson_length_minutes, created.capacity), (60, 1))
        updated = lesson.update_instructor(
            created.id, self.payload(name="Renamed", room="B2", lesson_length_minutes=30, capacity=3), db=self.db
        )
        self.assertEqual((updated.name, updated.room, updated.lesson_length_minutes, updated.capacity),
                         ("Renamed", "B2", 30, 3))
        self.assertEqual([i.name for i in lesson.list_instructors(db=self.db)], ["Renamed"])
        lesson.delete_instructor(created.id, db=self.db)
        self.assertEqual(lesson.list_instructors(db=self.db), [])

    def test_validation_errors(self):
        for overrides, code in (
            ({"name": "  "}, "INVALID_INSTRUCTOR_NAME"),
            ({"lesson_length_minutes": 0}, "INVALID_LESSON_LENGTH"),
            ({"capacity": 0}, "INVALID_CAPACITY"),
        ):
            with self.assertRaises(HTTPException) as caught:
                lesson.create_instructor(self.payload(**overrides), db=self.db)
            self.assertEqual((caught.exception.status_code, caught.exception.detail["code"]), (422, code))
        self.assertEqual(self.repo.list_instructors(), [])

    def test_missing_instructor_is_404(self):
        with self.assertRaises(HTTPException) as caught:
            lesson.update_instructor("missing", self.payload(), db=self.db)
        self.assertEqual(caught.exception.status_code, 404)

    def test_delete_preserves_assignments_and_reports_them_for_review(self):
        first = self.repo.create_instructor("First")
        second = self.repo.create_instructor("Second")
        a, b = self.repo.create_student("S1", "Pat Test"), self.repo.create_student("S2", "Sam Test")
        self.repo.set_placement(a.id, LessonPlacementType.MANUAL, first.id, 540)
        self.repo.set_placement(b.id, LessonPlacementType.MANUAL, second.id, 540)
        result = lesson.delete_instructor(first.id, db=self.db)
        self.assertEqual((result.removed_placement_count, result.placements_preserved), (0, 1))
        self.assertIsNotNone(self.repo.get_placement(a.id))
        self.assertIsNotNone(self.repo.get_placement(b.id))

    def test_update_with_tags_is_one_revision_and_none_keeps_tags(self):
        tag = self.repo.create_tag("strings")
        created = lesson.create_instructor(self.payload(tag_ids=[tag.id]), db=self.db)
        self.assertEqual([t.name for t in created.tags], ["STRINGS"])
        before = self.revision()
        lesson.update_instructor(created.id, self.payload(name="Again"), db=self.db)
        self.assertEqual(self.revision(), before + 1)
        self.assertEqual([t.name for t in lesson.list_instructors(db=self.db)[0].tags], ["STRINGS"])
        lesson.update_instructor(created.id, self.payload(tag_ids=[]), db=self.db)
        self.assertEqual(lesson.list_instructors(db=self.db)[0].tags, [])

    def test_tag_assignment_endpoint(self):
        created = lesson.create_instructor(self.payload(), db=self.db)
        tag = self.repo.create_tag("brass")
        assigned = lesson.set_instructor_tags(created.id, schemas.LessonInstructorTagsIn(tag_ids=[tag.id]), db=self.db)
        self.assertEqual([t.name for t in assigned.tags], ["BRASS"])
        with self.assertRaises(HTTPException) as caught:
            lesson.set_instructor_tags(created.id, schemas.LessonInstructorTagsIn(tag_ids=["nope"]), db=self.db)
        self.assertEqual(caught.exception.status_code, 404)


class TagApiTests(LessonTestCase):
    def test_create_normalizes_and_rejects_duplicates(self):
        tag = lesson.create_tag(schemas.LessonTagIn(name="  mental health "), db=self.db)
        self.assertEqual(tag.name, "MENTAL HEALTH")
        self.assertTrue(tag.is_unused)
        with self.assertRaises(HTTPException) as caught:
            lesson.create_tag(schemas.LessonTagIn(name="Mental Health"), db=self.db)
        self.assertEqual(caught.exception.status_code, 409)

    def test_rename_and_duplicate_rename(self):
        a = lesson.create_tag(schemas.LessonTagIn(name="a"), db=self.db)
        lesson.create_tag(schemas.LessonTagIn(name="b"), db=self.db)
        self.assertEqual(lesson.rename_tag(a.id, schemas.LessonTagIn(name=" c "), db=self.db).name, "C")
        with self.assertRaises(HTTPException):
            lesson.rename_tag(a.id, schemas.LessonTagIn(name="B"), db=self.db)

    def test_usage_counts_and_delete_reports_affected(self):
        tag = self.repo.create_tag("x")
        self.repo.create_tag("unused")
        self.repo.create_instructor("Synthetic Teacher", tag_ids=[tag.id])
        listed = {t.name: t for t in lesson.list_tags(db=self.db)}
        self.assertEqual(listed["X"].instructor_count, 1)
        self.assertFalse(listed["X"].is_unused)
        self.assertTrue(listed["UNUSED"].is_unused)
        result = lesson.delete_tag(tag.id, db=self.db)
        self.assertEqual(result.affected_instructor_count, 1)
        self.assertEqual(self.repo.list_instructors()[0].tags, [])
        self.assertEqual([t.name for t in self.repo.list_tags()], ["UNUSED"])

    def test_student_use_counts_as_in_use(self):
        tag = self.repo.create_tag("x")
        student = self.repo.create_student("S1", "Pat Test")
        self.repo.set_tag_state(student.id, tag.id, LessonTagState.REQUIRED)
        listed = lesson.list_tags(db=self.db)[0]
        self.assertEqual(listed.student_count, 1)
        self.assertFalse(listed.is_unused)

    def test_each_tag_change_is_one_source_change(self):
        before = self.revision()
        tag = lesson.create_tag(schemas.LessonTagIn(name="x"), db=self.db)
        lesson.rename_tag(tag.id, schemas.LessonTagIn(name="y"), db=self.db)
        lesson.delete_tag(tag.id, db=self.db)
        self.assertEqual(self.revision(), before + 3)


class AvailabilityApiTests(LessonTestCase):
    def test_edit_persists_and_replaces(self):
        instructor = self.repo.create_instructor("Synthetic Teacher")
        io_in = schemas.LessonInstructorAvailabilityIO(windows=[
            schemas.LessonAvailabilityWindowIO(day="Mon", start_minute=540, end_minute=600, status="Available"),
            schemas.LessonAvailabilityWindowIO(day="Monday", start_minute=600, end_minute=660, status="Tentative"),
            schemas.LessonAvailabilityWindowIO(day="Friday", start_minute=720, end_minute=780, status="Unavailable"),
        ])
        saved = lesson.set_instructor_availability(instructor.id, io_in, db=self.db)
        self.assertEqual(
            sorted((w.day, w.start_minute, w.status) for w in saved.windows),
            [("Friday", 720, "Unavailable"), ("Monday", 540, "Available"), ("Monday", 600, "Tentative")],
        )
        self.assertEqual(len(lesson.get_instructor_availability(instructor.id, db=self.db).windows), 3)
        self.assertEqual(lesson.list_instructors(db=self.db)[0].availability_window_count, 3)
        lesson.set_instructor_availability(instructor.id, schemas.LessonInstructorAvailabilityIO(windows=[]), db=self.db)
        self.assertEqual(lesson.get_instructor_availability(instructor.id, db=self.db).windows, [])

    def test_invalid_window_rejected_without_partial_write(self):
        instructor = self.repo.create_instructor("Synthetic Teacher", availability=[window()])
        self.db.commit()
        bad = schemas.LessonInstructorAvailabilityIO(windows=[
            schemas.LessonAvailabilityWindowIO(day="Monday", start_minute=600, end_minute=540, status="Available"),
        ])
        with self.assertRaises(HTTPException) as caught:
            lesson.set_instructor_availability(instructor.id, bad, db=self.db)
        self.assertEqual(caught.exception.status_code, 422)
        self.assertEqual(len(self.repo.instructor_availability(instructor.id)), 1)

    def test_availability_edit_is_one_revision(self):
        instructor = self.repo.create_instructor("Synthetic Teacher")
        before = self.revision()
        lesson.set_instructor_availability(
            instructor.id,
            schemas.LessonInstructorAvailabilityIO(windows=[
                schemas.LessonAvailabilityWindowIO(day="Monday", start_minute=540, end_minute=600, status="Available"),
            ]),
            db=self.db,
        )
        self.assertEqual(self.revision(), before + 1)


class RevisionTests(LessonTestCase):
    def test_instructor_changes_bump_once_and_mark_existing_schedule_stale(self):
        student = self.repo.create_student("S1", "Pat Test")
        instructor = self.repo.create_instructor("Synthetic Teacher")
        self.repo.set_placement(student.id, LessonPlacementType.MANUAL, instructor.id, 540)
        self.assertFalse(self.repo.get_state().is_schedule_stale)
        before = self.revision()
        self.repo.update_instructor(instructor.id, "Synthetic Teacher", room="B")
        self.assertEqual(self.revision(), before + 1)
        self.assertTrue(self.repo.get_state().is_schedule_stale)

    def test_changes_without_a_schedule_never_mark_stale(self):
        instructor = self.repo.create_instructor("Synthetic Teacher")
        self.repo.set_instructor_availability(instructor.id, [window()])
        self.repo.update_instructor(instructor.id, "Synthetic Teacher", room="B")
        self.assertFalse(self.repo.get_state().is_schedule_stale)


class ImportParsingTests(unittest.TestCase):
    def test_csv_fields_availability_and_tags(self):
        row = parse([GOOD_ROW]).rows[0]
        self.assertEqual(row.errors, [])
        self.assertEqual(
            (row.instructor_name, row.instructor_email, row.room, row.lesson_length_minutes, row.capacity),
            ("Synthetic Teacher", "teacher@example.test", "Room 101", 45, 2),
        )
        self.assertEqual(row.tags, ["VIOLIN", "STRINGS"])
        self.assertEqual(
            [(w.day, w.start_minute, w.end_minute, w.status, w.source) for w in row.windows],
            [("Monday", 540, 720, "Available", "import"), ("Monday", 780, 900, "Tentative", "import")],
        )

    def test_xlsx_is_supported(self):
        parsed = importer.parse_instructors("i.xlsx", xlsx_bytes([GOOD_ROW]), None, mapping())
        self.assertEqual(parsed.rows[0].instructor_name, "Synthetic Teacher")

    def test_defaults_when_blank_or_unmapped(self):
        row = parse([row_with(name="Synthetic Teacher")]).rows[0]
        self.assertEqual((row.lesson_length_minutes, row.capacity, row.errors), (60, 1, []))
        unmapped = parse([GOOD_ROW], lesson_length=None, capacity=None).rows[0]
        self.assertEqual((unmapped.lesson_length_minutes, unmapped.capacity), (60, 1))

    def test_validation_errors_are_reported_per_row(self):
        errors = parse([row_with(name="", length="abc", capacity=0, monday="nonsense")]).rows[0].errors
        self.assertTrue(any("Instructor Name is required" in e for e in errors))
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
        with self.assertRaises(importer.LessonInstructorImportError) as caught:
            parse([GOOD_ROW], instructor_name=None)
        self.assertEqual(caught.exception.code, "INVALID_MAPPING")
        with self.assertRaises(importer.LessonInstructorImportError):
            parse([GOOD_ROW], room="Instructor Name")
        with self.assertRaises(importer.LessonInstructorImportError):
            parse([GOOD_ROW], capacity="Missing Column")

    def test_empty_file_rejected(self):
        with self.assertRaises(importer.LessonInstructorImportError) as caught:
            parse([])
        self.assertEqual(caught.exception.code, "NO_ROWS")

    def test_inspect_suggests_mapping_including_days(self):
        suggested = importer.inspect_file("i.csv", csv_bytes([GOOD_ROW]), None).suggested_mapping
        self.assertEqual(suggested.instructor_name, "Instructor Name")
        self.assertEqual(suggested.lesson_length, "Lesson Length")
        self.assertEqual(suggested.days["Monday"], "Mon")
        self.assertEqual(suggested.days["Sunday"], "Sun")


class ImportFlowTests(LessonTestCase):
    def seed(self):
        keep = self.repo.create_tag("violin")
        self.repo.create_tag("orphan")
        old = self.repo.create_instructor("Old Teacher", tag_ids=[keep.id], availability=[window()])
        student = self.repo.create_student("S1", "Pat Test")
        self.repo.set_placement(student.id, LessonPlacementType.MANUAL, old.id, 540)
        self.db.commit()
        return student

    def call(self, endpoint, rows=(GOOD_ROW,), **kwargs):
        return asyncio.run(endpoint(
            file=upload(csv_bytes(list(rows))), mapping=json.dumps(MAPPING), sheet_name=None, db=self.db, **kwargs
        ))

    def test_preview_writes_nothing_and_reports_impact(self):
        self.seed()
        before = (self.revision(), len(self.repo.list_instructors()), len(self.repo.list_tags()))
        preview = self.call(lesson.preview_instructor_import)
        self.assertTrue(preview.can_apply)
        self.assertEqual((preview.existing_instructor_count, preview.existing_placement_count), (1, 1))
        self.assertEqual(preview.new_tags, ["STRINGS"])
        self.assertEqual(preview.reused_tags, ["VIOLIN"])
        self.assertEqual(preview.rows[0].availability, ["Mon 9:00 AM-12:00 PM", "Tentative Mon 1:00 PM-3:00 PM"])
        self.assertEqual(before, (self.revision(), len(self.repo.list_instructors()), len(self.repo.list_tags())))

    def test_preview_flags_invalid_rows(self):
        preview = self.call(lesson.preview_instructor_import, rows=[GOOD_ROW, row_with(name="", room="Room X")])
        self.assertEqual((preview.valid_rows, preview.invalid_rows, preview.can_apply), (1, 1, False))

    def test_apply_requires_confirmation(self):
        self.seed()
        with self.assertRaises(HTTPException) as caught:
            self.call(lesson.apply_instructor_import, confirmed=False)
        self.assertEqual(caught.exception.detail["code"], "CONFIRMATION_REQUIRED")
        self.assertEqual([i.name for i in self.repo.list_instructors()], ["Old Teacher"])

    def test_apply_replaces_instructors_preserving_tags_and_students(self):
        student = self.seed()
        before = self.revision()
        result = self.call(lesson.apply_instructor_import, confirmed=True)
        self.assertEqual(
            (
                result.instructors_removed, result.placements_removed, result.placements_preserved,
                result.instructors_created, result.tags_created,
            ),
            (1, 0, 1, 1, 1),
        )
        instructors = self.repo.list_instructors()
        self.assertEqual([i.name for i in instructors], ["Synthetic Teacher"])
        self.assertEqual({t.name for t in self.repo.list_tags()}, {"VIOLIN", "ORPHAN", "STRINGS"})
        self.assertEqual({t.name for t in instructors[0].tags}, {"VIOLIN", "STRINGS"})
        self.assertEqual([s.name for s in self.repo.list_students()], ["Pat Test"])
        windows = self.repo.instructor_availability(instructors[0].id)
        self.assertEqual((len(windows), windows[0].source), (2, "import"))
        self.assertIsNotNone(self.repo.get_placement(student.id))
        state = self.repo.get_state()
        self.assertEqual(state.source_data_revision, before + 1)
        self.assertTrue(state.is_schedule_stale)

    def test_apply_reparses_and_rejects_invalid_rows_atomically(self):
        student = self.seed()
        before = self.revision()
        with self.assertRaises(HTTPException) as caught:
            self.call(lesson.apply_instructor_import, rows=[GOOD_ROW, row_with(name="", room="Room X")], confirmed=True)
        self.assertEqual(caught.exception.detail["code"], "VALIDATION_FAILED")
        self.assertEqual([i.name for i in self.repo.list_instructors()], ["Old Teacher"])
        self.assertEqual(self.revision(), before)
        self.assertIsNotNone(self.repo.get_placement(student.id))

    def test_inspect_endpoint(self):
        inspection = asyncio.run(lesson.inspect_instructor_import(file=upload(csv_bytes([GOOD_ROW])), sheet_name=None))
        self.assertIn("Instructor Name", inspection.columns)

    def test_invalid_mapping_json_is_400(self):
        with self.assertRaises(HTTPException) as caught:
            asyncio.run(lesson.preview_instructor_import(
                file=upload(csv_bytes([GOOD_ROW])), mapping="not json", sheet_name=None, db=self.db
            ))
        self.assertEqual(caught.exception.status_code, 400)

    def test_replace_instructors_repository_contract(self):
        self.seed()
        result = self.repo.replace_instructors([
            (InstructorInput(name="A"), [window()], ["new tag", "NEW TAG"]),
            (InstructorInput(name="B", lesson_length_minutes=30, capacity=2), [], []),
        ])
        self.assertEqual((result.instructors_created, result.tags_created), (2, 1))
        with self.assertRaises(LessonValidationError):
            self.repo.replace_instructors([(InstructorInput(name=""), [], [])])


class ClinicalIsolationTests(LessonTestCase):
    def test_import_and_tags_never_touch_clinical(self):
        clinical = ClinicalRepository(self.db)
        clinical.create_tag("violin")
        clinical.create_student(student_id="S1", name="Synthetic Teacher", email="")
        self.db.commit()
        clinical_revision = clinical.get_state().source_data_revision
        asyncio.run(lesson.apply_instructor_import(
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
