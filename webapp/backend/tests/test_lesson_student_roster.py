"""Student Roster: CRUD, validation, tag states, availability, import, revisions, and Clinical isolation.

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
from app.services import lesson_student_importer as importer
from app.services.clinical_repository import ClinicalRepository
from app.services.lesson_repository import LessonValidationError
from tests.test_lesson_foundation import LessonTestCase, window

HEADERS = ["Student ID", "Name", "Email", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
MAPPING = {
    "student_id": "Student ID", "name": "Name", "email": "Email",
    "days": {"Monday": "Mon", "Tuesday": "Tue", "Wednesday": "Wed", "Thursday": "Thu",
             "Friday": "Fri", "Saturday": "Sat", "Sunday": "Sun"},
}
GOOD_ROW = ["S100", "Synthetic Student", "student@example.test", "9:00 AM-12:00 PM; T 1:00 PM-3:00 PM",
            "", "", "", "", "", ""]
BLANK_DAYS = ["", "", "", "", "", "", ""]


def row_with(student_id="S1", name="Pat Test", email="", monday=""):
    return [student_id, name, email, monday, *BLANK_DAYS[1:]]


def csv_bytes(rows, headers=HEADERS) -> bytes:
    return pd.DataFrame(rows, columns=headers).to_csv(index=False).encode()


def xlsx_bytes(rows, headers=HEADERS) -> bytes:
    buffer = io.BytesIO()
    pd.DataFrame(rows, columns=headers).to_excel(buffer, index=False)
    return buffer.getvalue()


def upload(data: bytes, name="students.csv") -> UploadFile:
    return UploadFile(file=io.BytesIO(data), filename=name)


def mapping(**overrides):
    return importer.LessonStudentImportMapping(**{**MAPPING, **overrides})


def parse(rows, **overrides):
    return importer.parse_students("students.csv", csv_bytes(rows), None, mapping(**overrides))


class StudentApiTests(LessonTestCase):
    def payload(self, **overrides):
        return schemas.LessonStudentIn(**{"student_id": "S1", "name": "Pat Test", **overrides})

    def test_create_update_delete(self):
        created = lesson.create_student(self.payload(email="pat@example.test"), db=self.db)
        self.assertEqual((created.student_id, created.name, created.email), ("S1", "Pat Test", "pat@example.test"))
        self.assertEqual((created.availability_window_count, created.tag_states), (0, {}))
        updated = lesson.update_student(created.id, self.payload(student_id="S2", name="Renamed"), db=self.db)
        self.assertEqual((updated.student_id, updated.name, updated.email), ("S2", "Renamed", ""))
        self.assertEqual([s.name for s in lesson.list_students(db=self.db)], ["Renamed"])
        lesson.delete_student(created.id, db=self.db)
        self.assertEqual(lesson.list_students(db=self.db), [])

    def test_validation_errors(self):
        for overrides, code in (
            ({"student_id": "  "}, "INVALID_STUDENT_ID"),
            ({"name": "  "}, "INVALID_STUDENT_NAME"),
        ):
            with self.assertRaises(HTTPException) as caught:
                lesson.create_student(self.payload(**overrides), db=self.db)
            self.assertEqual((caught.exception.status_code, caught.exception.detail["code"]), (422, code))
        self.assertEqual(self.repo.list_students(), [])

    def test_email_is_optional(self):
        created = lesson.create_student(self.payload(), db=self.db)
        self.assertEqual(created.email, "")

    def test_duplicate_student_ids_are_case_insensitive(self):
        lesson.create_student(self.payload(student_id="ab-1"), db=self.db)
        with self.assertRaises(HTTPException) as caught:
            lesson.create_student(self.payload(student_id=" AB-1 ", name="Other"), db=self.db)
        self.assertEqual((caught.exception.status_code, caught.exception.detail["code"]), (409, "DUPLICATE_STUDENT_ID"))
        other = lesson.create_student(self.payload(student_id="AB-2", name="Other"), db=self.db)
        with self.assertRaises(HTTPException) as caught:
            lesson.update_student(other.id, self.payload(student_id="Ab-1", name="Other"), db=self.db)
        self.assertEqual(caught.exception.detail["code"], "DUPLICATE_STUDENT_ID")
        # A student may keep its own Student ID, including a case change.
        same = lesson.update_student(other.id, self.payload(student_id="ab-2", name="Other"), db=self.db)
        self.assertEqual(same.student_id, "ab-2")

    def test_missing_student_is_404(self):
        with self.assertRaises(HTTPException) as caught:
            lesson.update_student("missing", self.payload(), db=self.db)
        self.assertEqual(caught.exception.status_code, 404)
        with self.assertRaises(HTTPException) as caught:
            lesson.get_student_availability("missing", db=self.db)
        self.assertEqual(caught.exception.status_code, 404)

    def test_delete_removes_source_dependents_but_preserves_schedule_decision(self):
        tag = self.repo.create_tag("violin")
        instructor = self.repo.create_instructor("Synthetic Teacher")
        student = self.repo.create_student("S1", "Pat Test", availability=[window()], tag_states={tag.id: "required"})
        self.repo.set_placement(student.id, LessonPlacementType.MANUAL, instructor.id, 540)
        self.db.commit()
        lesson.delete_student(student.id, db=self.db)
        self.assertEqual(len(self.repo.list_placements()), 1)
        self.assertEqual(self.repo.student_tag_states(), {})
        self.assertEqual(self.repo.student_availability_counts(), {})


class TagStateTests(LessonTestCase):
    def setUp(self):
        super().setUp()
        self.violin = self.repo.create_tag("violin")
        self.piano = self.repo.create_tag("piano")
        self.student = self.repo.create_student("S1", "Pat Test")
        self.db.commit()

    def test_states_persist_and_neutral_is_absence_of_a_row(self):
        out = lesson.set_student_tag_states(
            self.student.id,
            schemas.LessonStudentTagStatesIn(
                tag_states={self.violin.id: "required", self.piano.id: "restricted"}
            ),
            db=self.db,
        )
        self.assertEqual(out.tag_states, {self.violin.id: "required", self.piano.id: "restricted"})
        out = lesson.set_student_tag_states(
            self.student.id,
            schemas.LessonStudentTagStatesIn(tag_states={self.violin.id: "neutral", self.piano.id: "preferred"}),
            db=self.db,
        )
        self.assertEqual(out.tag_states, {self.piano.id: "preferred"})
        self.assertEqual(self.repo.get_tag_state(self.student.id, self.violin.id), LessonTagState.NEUTRAL)
        self.assertEqual(self.repo.student_tag_states(), {self.student.id: {self.piano.id: LessonTagState.PREFERRED}})
        self.assertEqual(lesson.list_students(db=self.db)[0].tag_states, {self.piano.id: "preferred"})

    def test_invalid_state_and_unknown_tag_are_rejected_without_changes(self):
        for states, code in (({self.violin.id: "maybe"}, "INVALID_TAG_STATE"), ({"missing": "required"}, "TAG_NOT_FOUND")):
            with self.assertRaises(HTTPException) as caught:
                lesson.set_student_tag_states(
                    self.student.id, schemas.LessonStudentTagStatesIn(tag_states=states), db=self.db
                )
            self.assertEqual(caught.exception.detail["code"], code)
        self.assertEqual(self.repo.student_tag_states(), {})

    def test_create_and_update_carry_tag_states_in_the_same_revision(self):
        before = self.revision()
        created = lesson.create_student(
            schemas.LessonStudentIn(student_id="S2", name="Two", tag_states={self.violin.id: "required"}), db=self.db
        )
        self.assertEqual((self.revision(), created.tag_states), (before + 1, {self.violin.id: "required"}))
        lesson.update_student(
            created.id,
            schemas.LessonStudentIn(student_id="S2", name="Two", tag_states={self.piano.id: "preferred"}),
            db=self.db,
        )
        self.assertEqual(self.revision(), before + 2)
        self.assertEqual(self.repo.student_tag_states()[created.id], {self.piano.id: LessonTagState.PREFERRED})

    def test_update_without_tag_states_leaves_them_unchanged(self):
        self.repo.set_tag_state(self.student.id, self.violin.id, LessonTagState.REQUIRED)
        lesson.update_student(self.student.id, schemas.LessonStudentIn(student_id="S1", name="Renamed"), db=self.db)
        self.assertEqual(self.repo.get_tag_state(self.student.id, self.violin.id), LessonTagState.REQUIRED)

    def test_tag_usage_counts_non_neutral_students_and_delete_cascades(self):
        self.repo.set_tag_state(self.student.id, self.violin.id, LessonTagState.RESTRICTED)
        self.db.commit()
        counts = {t.name: (t.student_count, t.is_unused) for t in lesson.list_tags(db=self.db)}
        self.assertEqual(counts, {"PIANO": (0, True), "VIOLIN": (1, False)})
        deleted = lesson.delete_tag(self.violin.id, db=self.db)
        self.assertEqual(deleted.affected_student_count, 1)
        self.assertEqual(self.repo.student_tag_states(), {})


class StudentAvailabilityTests(LessonTestCase):
    def setUp(self):
        super().setUp()
        self.student = self.repo.create_student("S1", "Pat Test")
        self.db.commit()

    @staticmethod
    def io(*items):
        return schemas.LessonStudentAvailabilityIO(
            windows=[schemas.LessonAvailabilityWindowIO(day=d, start_minute=s, end_minute=e, status=st)
                     for d, s, e, st in items]
        )

    def test_save_replaces_windows_and_normalizes_day_names(self):
        lesson.set_student_availability(
            self.student.id, self.io(("Mon", 540, 600, "Available"), ("Mon", 600, 660, "Tentative")), db=self.db
        )
        replaced = lesson.set_student_availability(
            self.student.id, self.io(("Tue", 480, 540, "Available")), db=self.db
        )
        self.assertEqual([(w.day, w.start_minute, w.end_minute, w.status) for w in replaced.windows],
                         [("Tuesday", 480, 540, "Available")])
        fetched = lesson.get_student_availability(self.student.id, db=self.db)
        self.assertEqual(fetched.windows, replaced.windows)
        self.assertEqual(lesson.list_students(db=self.db)[0].availability_window_count, 1)

    def test_clearing_removes_every_window(self):
        lesson.set_student_availability(self.student.id, self.io(("Monday", 540, 600, "Available")), db=self.db)
        lesson.set_student_availability(self.student.id, self.io(), db=self.db)
        self.assertEqual(lesson.get_student_availability(self.student.id, db=self.db).windows, [])

    def test_invalid_windows_are_rejected_and_keep_existing(self):
        lesson.set_student_availability(self.student.id, self.io(("Monday", 540, 600, "Available")), db=self.db)
        with self.assertRaises(HTTPException) as caught:
            lesson.set_student_availability(self.student.id, self.io(("Monday", 600, 540, "Available")), db=self.db)
        self.assertEqual(caught.exception.status_code, 422)
        self.assertEqual(len(lesson.get_student_availability(self.student.id, db=self.db).windows), 1)

    def test_availability_does_not_touch_instructor_availability(self):
        instructor = self.repo.create_instructor("Synthetic Teacher", availability=[window()])
        lesson.set_student_availability(self.student.id, self.io(("Monday", 540, 600, "Available")), db=self.db)
        self.assertEqual(len(self.repo.instructor_availability(instructor.id)), 1)


class RevisionTests(LessonTestCase):
    def seed_schedule(self):
        student = self.repo.create_student("S1", "Pat Test")
        instructor = self.repo.create_instructor("Synthetic Teacher")
        tag = self.repo.create_tag("violin")
        self.repo.set_placement(student.id, LessonPlacementType.MANUAL, instructor.id, 540)
        self.assertFalse(self.repo.get_state().is_schedule_stale)
        return student, tag

    def test_each_student_change_bumps_once_and_marks_an_existing_schedule_stale(self):
        student, tag = self.seed_schedule()
        operations = (
            lambda: self.repo.update_student(student.id, "S1", "Renamed"),
            lambda: self.repo.set_student_availability(student.id, [window()]),
            lambda: self.repo.set_student_tag_states(student.id, {tag.id: "required"}),
            lambda: self.repo.set_tag_state(student.id, tag.id, LessonTagState.PREFERRED),
            lambda: self.repo.create_student("S2", "Other"),
        )
        for operation in operations:
            self.repo.get_state().is_schedule_stale = False
            self.db.flush()
            before = self.revision()
            operation()
            self.assertEqual(self.revision(), before + 1)
            self.assertTrue(self.repo.get_state().is_schedule_stale)

    def test_delete_bumps_once_and_marks_stale_when_other_placements_remain(self):
        keep, _ = self.seed_schedule()
        other = self.repo.create_student("S2", "Other")
        before = self.revision()
        self.repo.delete_student(other.id)
        self.assertEqual(self.revision(), before + 1)
        self.assertTrue(self.repo.get_state().is_schedule_stale)
        self.assertEqual(len(self.repo.list_placements()), 1)

    def test_changes_without_a_schedule_never_mark_stale(self):
        student = self.repo.create_student("S1", "Pat Test")
        self.repo.set_student_availability(student.id, [window()])
        self.repo.update_student(student.id, "S1", "Renamed")
        self.assertFalse(self.repo.get_state().is_schedule_stale)

    def test_failed_save_does_not_bump(self):
        student, _ = self.seed_schedule()
        before = self.revision()
        with self.assertRaises(LessonValidationError):
            self.repo.update_student(student.id, "S1", "  ")
        with self.assertRaises(LessonValidationError):
            self.repo.set_student_tag_states(student.id, {"missing": "required"})
        self.assertEqual(self.revision(), before)

    def test_deleting_last_placed_student_keeps_remaining_schedule_stale(self):
        student, _ = self.seed_schedule()
        before = self.revision()
        self.repo.delete_student(student.id)
        self.assertEqual(self.revision(), before + 1)
        self.assertEqual(len(self.repo.list_placements()), 1)
        self.assertTrue(self.repo.get_state().is_schedule_stale)


class ImportParsingTests(unittest.TestCase):
    def test_csv_fields_and_availability(self):
        parsed = parse([GOOD_ROW])
        row = parsed.rows[0]
        self.assertEqual((row.student_id, row.name, row.email, row.errors), ("S100", "Synthetic Student",
                                                                           "student@example.test", []))
        self.assertEqual([(w.day, w.start_minute, w.end_minute, w.status) for w in row.windows],
                         [("Monday", 540, 720, "Available"), ("Monday", 780, 900, "Tentative")])
        self.assertTrue(parsed.availability_mapped)

    def test_xlsx_is_supported(self):
        parsed = importer.parse_students("students.xlsx", xlsx_bytes([GOOD_ROW]), None, mapping())
        self.assertEqual(parsed.rows[0].student_id, "S100")

    def test_required_fields(self):
        parsed = parse([row_with(student_id="", name="Pat"), row_with(student_id="S2", name="", email="a@b.test")])
        self.assertIn("Student ID is required.", parsed.rows[0].errors)
        self.assertIn("Name is required.", parsed.rows[1].errors)

    def test_email_is_optional(self):
        self.assertEqual(parse([row_with()]).rows[0].errors, [])

    def test_duplicate_ids_in_file_are_case_insensitive_errors(self):
        parsed = parse([row_with("ab-1", "One"), row_with("AB-1", "Two"), row_with("AB-2", "Three")])
        self.assertTrue(any("duplicated" in message for message in parsed.rows[0].errors))
        self.assertTrue(any("duplicated" in message for message in parsed.rows[1].errors))
        self.assertEqual(parsed.rows[2].errors, [])

    def test_invalid_availability_is_a_row_error(self):
        parsed = parse([row_with(monday="not a time")])
        self.assertTrue(parsed.rows[0].errors)

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

    def test_availability_columns_are_optional(self):
        parsed = importer.parse_students(
            "students.csv", csv_bytes([row_with()], HEADERS), None,
            importer.LessonStudentImportMapping(student_id="Student ID", name="Name", email="Email"),
        )
        self.assertFalse(parsed.availability_mapped)
        self.assertEqual(parsed.rows[0].windows, [])

    def test_mapping_validation(self):
        data = csv_bytes([row_with()])
        for bad in (
            mapping(student_id=None),
            mapping(name=None),
            mapping(name="Student ID"),
            mapping(email="Missing"),
        ):
            with self.assertRaises(importer.LessonStudentImportError) as caught:
                importer.parse_students("students.csv", data, None, bad)
            self.assertEqual(caught.exception.code, "INVALID_MAPPING")

    def test_blank_rows_are_skipped_and_empty_file_is_rejected(self):
        self.assertEqual(len(parse([row_with(), ["", "", "", "", "", "", "", "", "", ""]]).rows), 1)
        with self.assertRaises(importer.LessonStudentImportError) as caught:
            parse([])
        self.assertEqual(caught.exception.code, "NO_ROWS")

    def test_inspect_suggests_mapping(self):
        inspection = importer.inspect_file("students.csv", csv_bytes([GOOD_ROW]), None)
        suggested = inspection.suggested_mapping
        self.assertEqual((suggested.student_id, suggested.name, suggested.email),
                         ("Student ID", "Name", "Email"))
        self.assertEqual(suggested.days["Monday"], "Mon")


class ImportFlowTests(LessonTestCase):
    def seed(self):
        tag = self.repo.create_tag("violin")
        self.repo.create_tag("orphan")
        instructor = self.repo.create_instructor("Old Teacher", tag_ids=[tag.id], availability=[window()])
        old = self.repo.create_student("OLD1", "Old Student", availability=[window()], tag_states={tag.id: "required"})
        self.repo.set_placement(old.id, LessonPlacementType.MANUAL, instructor.id, 540)
        self.db.commit()
        return old, instructor, tag

    def call(self, endpoint, rows=(GOOD_ROW,), **kwargs):
        return asyncio.run(endpoint(
            file=upload(csv_bytes(list(rows))), mapping=json.dumps(MAPPING), sheet_name=None, db=self.db, **kwargs
        ))

    def test_preview_writes_nothing_and_reports_impact(self):
        self.seed()
        before = (self.revision(), len(self.repo.list_students()), len(self.repo.list_placements()))
        preview = self.call(lesson.preview_student_import)
        self.assertTrue(preview.can_apply)
        self.assertEqual((preview.existing_student_count, preview.existing_placement_count), (1, 1))
        self.assertEqual(preview.rows[0].availability, ["Mon 9:00 AM-12:00 PM", "Tentative Mon 1:00 PM-3:00 PM"])
        self.assertEqual(before, (self.revision(), len(self.repo.list_students()), len(self.repo.list_placements())))

    def test_preview_flags_invalid_rows(self):
        preview = self.call(lesson.preview_student_import, rows=[GOOD_ROW, row_with(student_id="", name="X")])
        self.assertEqual((preview.valid_rows, preview.invalid_rows, preview.can_apply), (1, 1, False))

    def test_apply_requires_confirmation(self):
        self.seed()
        with self.assertRaises(HTTPException) as caught:
            self.call(lesson.apply_student_import, confirmed=False)
        self.assertEqual(caught.exception.detail["code"], "CONFIRMATION_REQUIRED")
        self.assertEqual([s.student_id for s in self.repo.list_students()], ["OLD1"])

    def test_apply_replaces_student_source_data_and_preserves_old_placement_for_review(self):
        _, instructor, tag = self.seed()
        before = self.revision()
        result = self.call(lesson.apply_student_import, confirmed=True)
        self.assertEqual(
            (
                result.students_removed, result.placements_removed,
                result.placements_preserved, result.students_created,
            ),
            (1, 0, 1, 1),
        )
        students = self.repo.list_students()
        self.assertEqual([(s.student_id, s.name) for s in students], [("S100", "Synthetic Student")])
        self.assertEqual(len(self.repo.student_availability(students[0].id)), 2)
        self.assertEqual(self.repo.student_tag_states(), {})
        self.assertEqual(len(self.repo.list_placements()), 1)
        self.assertEqual([i.name for i in self.repo.list_instructors()], ["Old Teacher"])
        self.assertEqual(len(self.repo.instructor_availability(instructor.id)), 1)
        self.assertEqual(sorted(t.name for t in self.repo.list_tags()), ["ORPHAN", "VIOLIN"])
        self.assertEqual([t.name for t in self.repo.get_instructor(instructor.id).tags], ["VIOLIN"])
        self.assertEqual(self.revision(), before + 1)
        self.assertTrue(self.repo.get_state().is_schedule_stale)

    def test_apply_reparses_the_file_and_rejects_invalid_rows_atomically(self):
        self.seed()
        before = self.revision()
        with self.assertRaises(HTTPException) as caught:
            self.call(lesson.apply_student_import, rows=[GOOD_ROW, row_with("S100", "Dup")], confirmed=True)
        self.assertEqual(caught.exception.detail["code"], "VALIDATION_FAILED")
        self.assertEqual([s.student_id for s in self.repo.list_students()], ["OLD1"])
        self.assertEqual(len(self.repo.list_placements()), 1)
        self.assertEqual(self.revision(), before)

    def test_apply_may_reuse_a_student_id_from_the_replaced_roster(self):
        self.seed()
        result = self.call(lesson.apply_student_import, rows=[row_with("old1", "Same ID New Row")], confirmed=True)
        self.assertEqual(result.students_created, 1)
        self.assertEqual([s.student_id for s in self.repo.list_students()], ["old1"])

    def test_inspect_and_bad_mapping(self):
        inspection = asyncio.run(lesson.inspect_student_import(file=upload(csv_bytes([GOOD_ROW])), sheet_name=None))
        self.assertEqual(inspection.suggested_mapping.student_id, "Student ID")
        with self.assertRaises(HTTPException) as caught:
            asyncio.run(lesson.preview_student_import(
                file=upload(csv_bytes([GOOD_ROW])), mapping="not json", sheet_name=None, db=self.db
            ))
        self.assertEqual(caught.exception.detail["code"], "INVALID_MAPPING")

    def test_blank_availability_clears_to_none(self):
        self.call(lesson.apply_student_import, rows=[row_with()], confirmed=True)
        student = self.repo.list_students()[0]
        self.assertEqual(self.repo.student_availability(student.id), [])


class ClinicalIsolationTests(LessonTestCase):
    def test_student_workflows_never_touch_clinical(self):
        clinical = ClinicalRepository(self.db)
        clinical.create_student(student_id="S100", name="Synthetic Student", email="")
        clinical.create_tag("violin")
        self.db.commit()
        clinical_revision = clinical.get_state().source_data_revision

        lesson.create_student(schemas.LessonStudentIn(student_id="S100", name="Synthetic Student"), db=self.db)
        asyncio.run(lesson.apply_student_import(
            file=upload(csv_bytes([GOOD_ROW])), mapping=json.dumps(MAPPING), sheet_name=None,
            confirmed=True, db=self.db,
        ))
        self.assertEqual([s.student_id for s in clinical.list_students()], ["S100"])
        self.assertEqual([t.name for t in clinical.list_tags()], ["VIOLIN"])
        self.assertEqual(clinical.get_state().source_data_revision, clinical_revision)
        self.assertEqual([t.name for t in self.repo.list_tags()], [])
        self.assertEqual(self.repo.list_instructors(), [])


if __name__ == "__main__":
    unittest.main()
