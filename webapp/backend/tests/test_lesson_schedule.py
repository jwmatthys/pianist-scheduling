"""Synthetic tests for Lesson schedule persistence and API orchestration."""

import unittest

from fastapi import HTTPException

from app.lesson_models import LessonPlacementType, LessonSchedulePlacement, LessonUnplacedStudent
from app import lesson_schemas
from app.routers import lesson
from tests.test_lesson_foundation import LessonTestCase, window


class LessonSchedulePersistenceTests(LessonTestCase):
    def add_ready_inputs(self):
        student = self.repo.create_student(
            "S-1", "Student", availability=[window(start=540, end=660)]
        )
        instructor = self.repo.create_instructor(
            "Instructor", capacity=2, availability=[window(start=540, end=660)]
        )
        return student, instructor

    def test_generation_persists_assignments_result_and_revision(self):
        self.add_ready_inputs()
        initial_revision = self.repo.get_state().schedule_revision

        generated = self.repo.generate_schedule()

        self.assertTrue(generated.has_schedule)
        self.assertFalse(generated.is_schedule_stale)
        self.assertEqual(generated.source_data_revision, self.repo.get_state().source_data_revision)
        self.assertEqual(generated.schedule_revision, initial_revision + 1)
        self.assertEqual(generated.statistics.students_scheduled, 1)
        self.assertEqual(len(generated.placements), 1)
        placement = self.db.query(LessonSchedulePlacement).one()
        self.assertEqual(
            (placement.day, placement.start_minute, placement.end_minute),
            (generated.placements[0].day, generated.placements[0].start_minute, generated.placements[0].end_minute),
        )

    def test_generation_records_explicit_unplaced_student(self):
        self.repo.create_instructor(
            "Instructor", availability=[window(start=540, end=600)]
        )
        for number in range(2):
            self.repo.create_student(
                f"S-{number}", f"Student {number}", availability=[window(start=540, end=600)]
            )

        generated = self.repo.generate_schedule()

        self.assertEqual(generated.statistics.students_scheduled, 1)
        self.assertEqual(generated.statistics.students_unscheduled, 1)
        self.assertEqual(len(generated.unplaced_students), 1)
        self.assertEqual(generated.unplaced_students[0].reason.value, "CAPACITY_LIMIT")

    def test_manually_unplacing_an_automatically_unplaced_student_replaces_its_row(self):
        self.repo.create_instructor("Instructor", availability=[window(start=540, end=600)])
        for number in range(2):
            self.repo.create_student(
                f"S-{number}", f"Student {number}", availability=[window(start=540, end=600)]
            )
        generated = self.repo.generate_schedule()
        unplaced_id = generated.unplaced_students[0].student_id

        self.repo.set_manually_unplaced(unplaced_id)

        row = self.db.query(LessonUnplacedStudent).filter_by(student_id=unplaced_id).one()
        self.assertEqual(row.placement_type, "manuallyUnplaced")

    def test_source_change_marks_persisted_schedule_stale(self):
        self.add_ready_inputs()
        self.repo.generate_schedule()

        self.repo.create_student("S-2", "Second", availability=[window()])

        schedule = self.repo.get_schedule()
        self.assertTrue(schedule.has_schedule)
        self.assertTrue(schedule.is_schedule_stale)
        self.assertLess(schedule.source_data_revision, self.repo.get_state().source_data_revision)

    def test_stale_schedule_snapshot_survives_source_roster_deletion(self):
        student, _ = self.add_ready_inputs()
        self.repo.generate_schedule()

        self.repo.delete_student(student.id)

        stale = self.repo.get_schedule()
        self.assertTrue(stale.is_schedule_stale)
        self.assertEqual(len(stale.placements), 1)
        self.assertEqual(stale.placements[0].student_id, student.id)

    def test_regeneration_replaces_snapshot_and_advances_schedule_revision(self):
        self.add_ready_inputs()
        first = self.repo.generate_schedule()

        second = self.repo.generate_schedule()

        self.assertEqual(second.schedule_revision, first.schedule_revision + 1)
        self.assertEqual(second.placements, first.placements)
        self.assertEqual(self.db.query(LessonSchedulePlacement).count(), 1)

    def test_clear_removes_generated_schedule_and_increments_schedule_revision(self):
        self.add_ready_inputs()
        generated = self.repo.generate_schedule()

        cleared = self.repo.clear_schedule()

        self.assertFalse(cleared.has_schedule)
        self.assertEqual(cleared.placements, ())
        self.assertEqual(cleared.unplaced_students, ())
        self.assertEqual(cleared.schedule_revision, generated.schedule_revision + 1)
        self.assertEqual(self.db.query(LessonSchedulePlacement).count(), 0)
        self.assertFalse(self.repo.get_state().is_schedule_stale)

    def test_not_ready_inputs_block_generation(self):
        self.add_ready_inputs()
        self.repo.set_student_availability(self.repo.list_students()[0].id, [])

        with self.assertRaisesRegex(ValueError, "not ready"):
            self.repo.generate_schedule()
        self.assertFalse(self.repo.get_schedule().has_schedule)

    def test_generation_preserves_existing_manual_decisions(self):
        student, instructor = self.add_ready_inputs()
        manual = self.repo.set_placement(
            student.id, LessonPlacementType.MANUAL, instructor.id, 600
        )

        generated = self.repo.generate_schedule()

        self.assertEqual(self.repo.get_placement(student.id).id, manual.id)
        self.assertEqual(generated.manual_count, 1)
        self.assertEqual(generated.automatic_count, 0)
        self.assertEqual(generated.placements[0].start_minute, 600)
        self.assertEqual(generated.placements[0].instructor_id, instructor.id)

    def test_manually_unplaced_decision_survives_regeneration_and_is_excluded(self):
        student, _ = self.add_ready_inputs()
        another = self.repo.create_student(
            "S-2", "Second", availability=[window(start=540, end=660)]
        )
        self.repo.set_manually_unplaced(student.id)

        generated = self.repo.generate_schedule()

        self.assertEqual(self.repo.get_placement(student.id).placement_type, "manuallyUnplaced")
        self.assertEqual(generated.manually_unplaced_count, 1)
        self.assertEqual(generated.manual_count, 0)
        self.assertEqual(
            [placement.student_id for placement in generated.placements],
            [another.id],
        )

    def test_unassigned_students_are_eligible_but_manually_unplaced_are_not(self):
        first, instructor = self.add_ready_inputs()
        second = self.repo.create_student("S-2", "Second", availability=[window(start=540, end=660)])
        self.repo.generate_schedule()
        placed = next(item for item in self.repo.list_placements() if item.placement_type == "automatic")
        other = second.id if placed.lesson_student_id == first.id else first.id

        self.repo.remove_placement(placed.lesson_student_id)
        self.assertIsNone(self.repo.get_placement(placed.lesson_student_id))
        self.repo.set_manually_unplaced(other)
        generated = self.repo.generate_schedule()

        self.assertEqual([item.student_id for item in generated.placements], [placed.lesson_student_id])
        self.assertEqual(self.repo.get_placement(other).placement_type, "manuallyUnplaced")
        states = {item.student_id: item.placement_state for item in lesson._schedule_out(generated).entries}
        self.assertEqual(states[placed.lesson_student_id], "assigned")
        self.assertEqual(states[other], "manuallyUnplaced")

    def test_removing_a_placement_creates_unassigned_state(self):
        student, _ = self.add_ready_inputs()
        self.repo.generate_schedule()

        out = lesson.remove_placement_decision(student.id, self.db)

        self.assertIsNone(self.repo.get_placement(student.id))
        self.assertEqual(out.entries[0].placement_state, "unassigned")
        with self.assertRaises(HTTPException) as raised:
            lesson.remove_placement_decision(student.id, self.db)
        self.assertEqual(raised.exception.detail["code"], "PLACEMENT_NOT_FOUND")

    def test_removing_override_returns_student_to_pool_without_generating(self):
        student, instructor = self.add_ready_inputs()
        self.repo.set_manual_placement(student.id, instructor.id, "Monday", 600)
        before_revision = self.repo.get_state().schedule_revision

        self.repo.remove_manual_override(student.id)

        self.assertIsNone(self.repo.get_placement(student.id))
        self.assertEqual(self.repo.get_state().schedule_revision, before_revision + 1)
        self.assertFalse(self.repo.get_schedule().has_schedule)
        self.assertEqual(self.db.query(LessonSchedulePlacement).count(), 0)

    def test_source_change_preserves_assignments_and_verification_marks_invalid(self):
        student, _ = self.add_ready_inputs()
        generated = self.repo.generate_schedule()
        self.repo.set_student_availability(student.id, [])

        stale = self.repo.get_schedule()

        self.assertTrue(stale.is_schedule_stale)
        self.assertEqual(stale.placements, generated.placements)
        self.assertEqual(stale.unverified_count, 1)
        self.assertEqual(stale.verification[0].warnings[0].code.value, "OUTSIDE_STUDENT_AVAILABILITY")
        self.assertEqual(stale.schedule_revision, generated.schedule_revision)

    def test_regeneration_replaces_only_automatic_and_keeps_manual_fixed(self):
        student, instructor = self.add_ready_inputs()
        second = self.repo.create_student(
            "S-2", "Second", availability=[window(start=540, end=660)]
        )
        self.repo.set_manual_placement(student.id, instructor.id, "Monday", 540)

        generated = self.repo.generate_schedule()

        self.assertEqual(generated.manual_count, 1)
        self.assertEqual(generated.automatic_count, 1)
        self.assertEqual(
            next(item for item in generated.placements if item.student_id == student.id).start_minute,
            540,
        )
        first_automatic = next(item for item in generated.placements if item.student_id == second.id)
        regenerated = self.repo.generate_schedule()
        self.assertEqual(regenerated.schedule_revision, generated.schedule_revision + 1)
        self.assertEqual(
            next(item for item in regenerated.placements if item.student_id == student.id).start_minute,
            540,
        )
        self.assertEqual(
            next(item for item in regenerated.placements if item.student_id == second.id),
            first_automatic,
        )

    def test_source_changes_preserve_deleted_roster_assignments(self):
        student, instructor = self.add_ready_inputs()
        self.repo.set_manual_placement(student.id, instructor.id, "Monday", 540)

        self.repo.delete_student(student.id)
        self.repo.delete_instructor(instructor.id)

        schedule = self.repo.get_schedule()
        self.assertTrue(schedule.is_schedule_stale)
        self.assertEqual(len(schedule.placements), 1)
        self.assertEqual(schedule.entries[0].student_name, "Student")
        self.assertEqual(schedule.entries[0].instructor_name, "Instructor")
        self.assertEqual(
            {warning.code.value for warning in schedule.verification[0].warnings},
            {"STUDENT_REMOVED", "INSTRUCTOR_REMOVED"},
        )


class LessonScheduleApiTests(LessonTestCase):
    def test_schedule_api_get_generate_and_clear(self):
        student = self.repo.create_student(
            "S-1", "Student", availability=[window(start=540, end=660)]
        )
        self.repo.create_instructor(
            "Instructor", availability=[window(start=540, end=660)]
        )

        initial = lesson.get_schedule(self.db)
        generated = lesson.generate_schedule(self.db)
        retrieved = lesson.get_schedule(self.db)
        cleared = lesson.clear_schedule(self.db)

        self.assertFalse(initial.has_schedule)
        self.assertEqual(generated.students_scheduled, 1)
        self.assertTrue(retrieved.has_schedule)
        self.assertEqual(retrieved.source_data_revision, self.repo.get_state().source_data_revision)
        self.assertFalse(cleared.has_schedule)
        self.assertEqual(cleared.students_scheduled, 0)
        self.assertTrue(student.id)

    def test_schedule_api_returns_structured_not_ready_error(self):
        with self.assertRaises(HTTPException) as raised:
            lesson.generate_schedule(self.db)

        self.assertEqual(raised.exception.status_code, 409)
        self.assertEqual(raised.exception.detail["code"], "NOT_READY")

    def test_manual_unplaced_remove_override_and_verification_api(self):
        student = self.repo.create_student(
            "S-1", "Student", availability=[window(start=540, end=660)]
        )
        instructor = self.repo.create_instructor(
            "Instructor", availability=[window(start=540, end=660)]
        )
        payload = lesson_schemas.LessonManualPlacementIn(
            student_id=student.id,
            instructor_id=instructor.id,
            day="Monday",
            start_minute=545,
        )

        assigned = lesson.create_manual_placement(payload, self.db)

        self.assertEqual(assigned.manual_count, 1)
        self.assertEqual(assigned.entries[0].placement_type, "manual")
        self.assertEqual(lesson.get_schedule_verification(self.db), [])
        lesson.manually_unplace_student(student.id, self.db)
        self.assertEqual(self.repo.get_placement(student.id).placement_type, "manuallyUnplaced")
        self.assertEqual(self.repo.get_schedule().manually_unplaced_count, 1)

        revision_before_remove = self.repo.get_state().schedule_revision
        removed = lesson.remove_manual_placement_override(student.id, self.db)

        self.assertIsNone(self.repo.get_placement(student.id))
        self.assertEqual(removed.schedule_revision, revision_before_remove + 1)
        self.assertFalse(removed.has_schedule)
        self.assertEqual(removed.entries[0].placement_type, None)
        self.assertEqual(removed.entries[0].placement_state, "unassigned")

    def test_verification_endpoint_reports_assignments_invalidated_by_source_change(self):
        student = self.repo.create_student(
            "S-1", "Student", availability=[window(start=540, end=660)]
        )
        instructor = self.repo.create_instructor(
            "Instructor", availability=[window(start=540, end=660)]
        )
        self.repo.generate_schedule()
        self.repo.set_student_availability(student.id, [])

        verification = lesson.get_schedule_verification(self.db)

        self.assertEqual(len(verification), 1)
        self.assertEqual(verification[0].student_id, student.id)
        self.assertIn(
            "OUTSIDE_STUDENT_AVAILABILITY",
            {item.code for item in verification[0].warnings},
        )


if __name__ == "__main__":
    unittest.main()
