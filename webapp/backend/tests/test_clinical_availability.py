"""Synthetic-data tests for Clinical Student availability persistence and API."""

from fastapi import HTTPException

from app.clinical_models import ClinicalStudentAvailabilityWindow, PlacementType
from app.routers import clinical
from app.services.availability import AvailabilityWindow
from app.services.clinical_repository import ClinicalValidationError
from tests.test_clinical_foundation import ClinicalTestCase, opportunity_input

IO = clinical.schemas.ClinicalAvailabilityWindowIO
Payload = clinical.schemas.ClinicalStudentAvailabilityIO


class StudentAvailabilityRepositoryTests(ClinicalTestCase):
    def setUp(self):
        super().setUp()
        self.student = self.repo.create_student("S1", "Pat", require_name=True)

    def test_replacement_stores_recurring_windows_not_cells(self):
        self.repo.set_student_availability(self.student.id, [
            AvailabilityWindow("Monday", 540, 570, "Available"),
            AvailabilityWindow("Monday", 570, 600, "Available"),
            AvailabilityWindow("Monday", 600, 660, "Tentative"),
            AvailabilityWindow("Tuesday", 420, 1260, "Unavailable"),
        ])
        rows = self.repo.student_availability(self.student.id)
        self.assertEqual(
            [(w.day, w.start_minute, w.end_minute, w.status) for w in rows],
            [
                ("Monday", 540, 600, "Available"),
                ("Monday", 600, 660, "Tentative"),
                ("Tuesday", 420, 1260, "Unavailable"),
            ],
        )
        self.repo.set_student_availability(self.student.id, [AvailabilityWindow("Friday", 480, 600, "Available")])
        self.assertEqual(self.db.query(ClinicalStudentAvailabilityWindow).count(), 1)
        self.assertEqual(self.repo.student_availability_counts(), {self.student.id: 1})

    def test_empty_replacement_clears_availability(self):
        self.repo.set_student_availability(self.student.id, [AvailabilityWindow("Monday", 480, 540, "Available")])
        self.repo.set_student_availability(self.student.id, [])
        self.assertEqual(self.repo.student_availability_counts(), {})

    def test_revision_increments_once_per_edit(self):
        before = self.repo.get_state().source_data_revision
        self.repo.set_student_availability(self.student.id, [
            AvailabilityWindow("Monday", 480, 540, "Available"),
            AvailabilityWindow("Tuesday", 480, 540, "Tentative"),
        ])
        self.assertEqual(self.repo.get_state().source_data_revision, before + 1)

    def test_stale_only_when_placements_exist(self):
        self.repo.set_student_availability(self.student.id, [AvailabilityWindow("Monday", 480, 540, "Available")])
        self.assertFalse(self.repo.get_state().is_schedule_stale)
        site = self.repo.create_opportunity(opportunity_input())
        self.repo.set_placement(self.student.id, PlacementType.MANUAL, site.id, 480)
        self.repo.set_student_availability(self.student.id, [AvailabilityWindow("Monday", 540, 600, "Available")])
        self.assertTrue(self.repo.get_state().is_schedule_stale)

    def test_invalid_and_conflicting_windows_store_nothing(self):
        self.repo.set_student_availability(self.student.id, [AvailabilityWindow("Monday", 480, 540, "Available")])
        before = self.repo.get_state().source_data_revision
        for bad in (
            [AvailabilityWindow("Monday", 540, 480, "Available")],
            [AvailabilityWindow("Funday", 480, 540, "Available")],
            [AvailabilityWindow("Monday", 480, 600, "Available"), AvailabilityWindow("Monday", 540, 660, "Tentative")],
        ):
            with self.assertRaises(ClinicalValidationError) as caught:
                self.repo.set_student_availability(self.student.id, bad)
            self.assertEqual(caught.exception.code, "INVALID_AVAILABILITY_WINDOW")
        self.assertEqual(self.repo.student_availability(self.student.id)[0].start_minute, 480)
        self.assertEqual(self.repo.get_state().source_data_revision, before)

    def test_unknown_student(self):
        with self.assertRaises(ClinicalValidationError) as caught:
            self.repo.set_student_availability("missing", [])
        self.assertEqual(caught.exception.code, "STUDENT_NOT_FOUND")


class StudentAvailabilityApiTests(ClinicalTestCase):
    def test_roundtrip_counts_and_errors(self):
        student = clinical.create_student(clinical.schemas.ClinicalStudentFields(student_id="S1", name="Pat"), self.db)
        self.assertEqual(clinical.get_student_availability(student.id, self.db).windows, [])
        saved = clinical.set_student_availability(student.id, Payload(windows=[
            IO(day="Monday", start_minute=420, end_minute=480, status="Unavailable"),
            IO(day="mon", start_minute=480, end_minute=600, status="Available"),
        ]), self.db)
        self.assertEqual([(w.day, w.start_minute, w.end_minute, w.status) for w in saved.windows], [
            ("Monday", 420, 480, "Unavailable"), ("Monday", 480, 600, "Available"),
        ])
        self.assertEqual(clinical.get_student_availability(student.id, self.db).windows, saved.windows)
        self.assertEqual(clinical.list_students(self.db)[0].availability_window_count, 2)
        with self.assertRaises(HTTPException) as caught:
            clinical.set_student_availability(
                student.id, Payload(windows=[IO(day="Monday", start_minute=600, end_minute=500, status="Available")]), self.db
            )
        self.assertEqual(caught.exception.status_code, 422)
        self.assertEqual(clinical.list_students(self.db)[0].availability_window_count, 2)
        with self.assertRaises(HTTPException) as caught:
            clinical.get_student_availability("missing", self.db)
        self.assertEqual(caught.exception.status_code, 404)
