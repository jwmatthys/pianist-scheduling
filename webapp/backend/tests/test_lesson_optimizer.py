"""Synthetic tests for the Lesson-specific exact optimizer."""

from dataclasses import replace
import unittest

from app.lesson_models import LessonTagState
from app.services.lesson_candidates import generate_lesson_candidates
from app.services.lesson_optimizer import optimize_lesson_schedule
from app.services.lesson_scheduling_models import LessonCandidateSet
from tests.test_lesson_foundation import LessonTestCase, window


class LessonOptimizerTests(LessonTestCase):
    def optimize(self):
        scheduling_input = self.repo.get_scheduling_input()
        candidates = generate_lesson_candidates(scheduling_input)
        return scheduling_input, candidates, optimize_lesson_schedule(scheduling_input, candidates)

    def test_maximizes_students_and_uses_stable_student_tie_break(self):
        instructor = self.repo.create_instructor(
            "Instructor", capacity=5, availability=[window(start=540, end=600)]
        )
        self.repo.create_student("S-2", "Bea", availability=[window(start=540, end=600)])
        self.repo.create_student("S-1", "Ada", availability=[window(start=540, end=600)])

        _, _, result = self.optimize()

        self.assertEqual(result.statistics.students_scheduled, 1)
        self.assertEqual(result.placements[0].student_id, self.repo.list_students()[0].id)
        self.assertEqual(result.placements[0].instructor_id, instructor.id)
        self.assertEqual(result.unplaced_students[0].reason.value, "CAPACITY_LIMIT")

    def test_instructor_capacity_is_enforced(self):
        instructor = self.repo.create_instructor(
            "Instructor", capacity=1, availability=[window(start=540, end=660)]
        )
        for index, start in enumerate((540, 600), 1):
            self.repo.create_student(
                f"S-{index}", f"Student {index}", availability=[window(start=start, end=start + 60)]
            )

        _, _, result = self.optimize()

        self.assertEqual(result.statistics.students_scheduled, 1)
        self.assertLessEqual(
            sum(item.instructor_id == instructor.id for item in result.placements), instructor.capacity
        )

    def test_non_overlapping_assignments_are_enforced_even_with_available_capacity(self):
        self.repo.create_instructor(
            "Instructor", capacity=4, availability=[window(start=540, end=660)]
        )
        for index in range(2):
            self.repo.create_student(
                f"S-{index}", f"Student {index}", availability=[window(start=540, end=600)]
            )

        _, _, result = self.optimize()

        self.assertEqual(result.statistics.students_scheduled, 1)
        self.assertEqual(result.unplaced_students[0].reason.value, "INSTRUCTOR_OVERLAP")

    def test_effective_capacity_is_limited_by_schedulable_lesson_time(self):
        self.repo.create_instructor(
            "Instructor", capacity=5, availability=[window(start=540, end=660)]
        )
        for index in range(3):
            self.repo.create_student(
                f"S-{index}", f"Student {index}", availability=[window(start=540, end=660)]
            )

        _, _, result = self.optimize()

        self.assertEqual(result.statistics.students_scheduled, 2)
        self.assertEqual(result.statistics.instructor_utilization, 1.0)

    def test_preferred_matches_are_optimized_before_tentative_time(self):
        tag = self.repo.create_tag("Advanced")
        student = self.repo.create_student(
            "S-1", "Student", availability=[window(start=540, end=600)],
            tag_states={tag.id: LessonTagState.PREFERRED},
        )
        self.repo.create_instructor(
            "A Unmatched", availability=[window(start=540, end=600)]
        )
        preferred = self.repo.create_instructor(
            "Z Preferred",
            availability=[window(start=540, end=600, status="Tentative")],
            tag_ids=[tag.id],
        )

        _, _, result = self.optimize()

        self.assertEqual(result.placements[0].student_id, student.id)
        self.assertEqual(result.placements[0].instructor_id, preferred.id)
        self.assertEqual(result.statistics.preferred_tag_matches, 1)
        self.assertEqual(result.statistics.tentative_minutes, 60)

    def test_tentative_minutes_are_taken_from_candidate_values(self):
        student = self.repo.create_student(
            "S-1", "Student", availability=[window(start=540, end=660)]
        )
        available = self.repo.create_instructor(
            "Available", availability=[window(start=540, end=600)]
        )
        self.repo.create_instructor(
            "Tentative",
            availability=[window(start=600, end=660, status="Tentative")],
        )

        scheduling_input, candidates, result = self.optimize()

        self.assertEqual(result.placements[0].student_id, student.id)
        self.assertEqual(result.placements[0].instructor_id, available.id)
        self.assertEqual(result.statistics.tentative_minutes, 0)
        self.assertEqual(result.placements[0].tentative_minutes, 0)
        self.assertEqual(candidates.source_data_revision, scheduling_input.source_data_revision)

    def test_gap_cost_breaks_ties_after_tentative_minutes(self):
        self.repo.create_instructor(
            "Instructor", capacity=2, availability=[window(start=540, end=780)]
        )
        self.repo.create_student(
            "S-1", "Flexible",
            availability=[
                window(start=540, end=600),
                window(start=720, end=780),
            ],
        )
        self.repo.create_student(
            "S-2", "Fixed", availability=[window(start=600, end=660)]
        )

        _, _, result = self.optimize()

        flexible = next(item for item in self.repo.list_students() if item.name == "Flexible")
        flexible_placement = next(item for item in result.placements if item.student_id == flexible.id)
        self.assertEqual(flexible_placement.start_minute, 540)
        self.assertEqual(result.statistics.gap_cost_minutes, 0)

    def test_gap_between_separate_availability_windows_is_not_counted(self):
        self.repo.create_instructor(
            "Instructor",
            capacity=2,
            availability=[window(start=540, end=600), window(start=660, end=720)],
        )
        self.repo.create_student("S-1", "First", availability=[window(start=540, end=600)])
        self.repo.create_student("S-2", "Second", availability=[window(start=660, end=720)])

        _, _, result = self.optimize()

        self.assertEqual(result.statistics.students_scheduled, 2)
        self.assertEqual(result.statistics.gap_cost_minutes, 0)

    def test_candidate_and_input_revision_mismatch_is_rejected(self):
        self.repo.create_student("S-1", "Student", availability=[window()])
        self.repo.create_instructor("Instructor", availability=[window()])
        scheduling_input = self.repo.get_scheduling_input()
        candidates = generate_lesson_candidates(scheduling_input)

        with self.assertRaisesRegex(ValueError, "different source-data revision"):
            optimize_lesson_schedule(
                replace(scheduling_input, source_data_revision=scheduling_input.source_data_revision + 1),
                candidates,
            )

    def test_infeasible_required_tag_candidate_cannot_be_promoted_by_optimizer(self):
        tag = self.repo.create_tag("Required")
        self.repo.create_student(
            "S-1", "Student",
            availability=[window(start=540, end=600)],
            tag_states={tag.id: LessonTagState.REQUIRED},
        )
        self.repo.create_instructor("Instructor", availability=[window(start=540, end=600)])
        scheduling_input = self.repo.get_scheduling_input()
        generated = generate_lesson_candidates(scheduling_input)
        forged = LessonCandidateSet(
            generated.source_data_revision,
            tuple(replace(item, is_feasible=True, issues=()) for item in generated.candidates),
        )

        result = optimize_lesson_schedule(scheduling_input, forged)

        self.assertEqual(result.placements, ())
        self.assertEqual(result.unplaced_students[0].reason.value, "NO_FEASIBLE_CANDIDATE")

    def test_student_without_any_candidate_is_reported_unplaced(self):
        self.repo.create_instructor(
            "Instructor", availability=[window(start=540, end=600)]
        )
        self.repo.create_student("S-1", "Placed", availability=[window(start=540, end=600)])
        self.repo.create_student(
            "S-2", "No Match", availability=[window(day="Tuesday", start=540, end=600)]
        )

        _, _, result = self.optimize()

        no_match = next(item for item in self.repo.list_students() if item.name == "No Match")
        self.assertEqual(
            next(item.reason.value for item in result.unplaced_students if item.student_id == no_match.id),
            "NO_FEASIBLE_CANDIDATE",
        )

    def test_identical_logical_inputs_produce_deterministic_results(self):
        self.repo.create_student("S-2", "Bea", availability=[window()])
        self.repo.create_student("S-1", "Ada", availability=[window()])
        self.repo.create_instructor("Instructor", capacity=2, availability=[window()])
        scheduling_input = self.repo.get_scheduling_input()
        candidates = generate_lesson_candidates(scheduling_input)
        reordered_input = replace(
            scheduling_input,
            students=tuple(reversed(scheduling_input.students)),
            instructors=tuple(reversed(scheduling_input.instructors)),
        )
        reordered_candidates = generate_lesson_candidates(reordered_input)

        first = optimize_lesson_schedule(scheduling_input, candidates)
        second = optimize_lesson_schedule(reordered_input, reordered_candidates)

        self.assertEqual(first, second)


    def test_wide_availability_with_many_equivalent_optima_finishes_quickly(self):
        import time

        days = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
        wide = [window(day, 540, 1020) for day in days]
        for index in range(5):
            self.repo.create_instructor(f"Instructor {index}", capacity=8, availability=wide)
            self.repo.create_student(f"S-{index}", f"Student {index}", availability=wide)

        started = time.monotonic()
        _, _, result = self.optimize()

        self.assertLess(time.monotonic() - started, 5)
        self.assertEqual(result.statistics.students_scheduled, 5)
        self.assertEqual(result.statistics.gap_cost_minutes, 0)


if __name__ == "__main__":
    unittest.main()
