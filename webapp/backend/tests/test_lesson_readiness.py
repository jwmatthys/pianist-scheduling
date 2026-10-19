"""Synthetic-data tests for Lesson Scheduling readiness analysis."""

from dataclasses import replace
import unittest

from app.lesson_models import LessonTagState
from app.routers import lesson
from app.services.availability import AvailabilityWindow
from app.services.lesson_candidates import generate_lesson_candidates
from app.services.lesson_readiness import analyze_lesson_readiness
from app.services.lesson_scheduling_models import (
    LessonReadinessCode,
    LessonReadinessSeverity,
)
from tests.test_lesson_foundation import LessonTestCase, window


class LessonReadinessTests(LessonTestCase):
    def report(self):
        return self.repo.get_readiness_report()

    def warning(self, report, code):
        return next(item for item in report.warnings if item.code is code)

    def test_empty_system_reports_no_students_and_no_instructors(self):
        report = self.report()
        self.assertFalse(report.is_ready)
        self.assertEqual((report.student_count, report.instructor_count), (0, 0))
        self.assertEqual(
            {item.code for item in report.warnings},
            {LessonReadinessCode.NO_STUDENTS, LessonReadinessCode.NO_INSTRUCTORS},
        )

    def test_no_students_with_instructors_reports_only_missing_students(self):
        self.repo.create_instructor("Instructor")
        report = self.report()
        self.assertFalse(report.is_ready)
        self.assertIn(LessonReadinessCode.NO_STUDENTS, {item.code for item in report.warnings})
        self.assertNotIn(LessonReadinessCode.NO_FEASIBLE_CANDIDATES, {item.code for item in report.warnings})

    def test_no_instructors_reports_student_without_feasible_instructor(self):
        self.repo.create_student("S-1", "Student", availability=[window()])
        report = self.report()
        codes = {item.code for item in report.warnings}
        self.assertIn(LessonReadinessCode.NO_INSTRUCTORS, codes)
        self.assertNotIn(LessonReadinessCode.STUDENT_WITHOUT_CANDIDATES, codes)
        self.assertFalse(report.is_ready)

    def test_student_without_availability_is_grouped_with_subject_detail(self):
        students = [self.repo.create_student(f"S-{i}", f"Student {i}") for i in range(3)]
        self.repo.create_instructor("Instructor", availability=[window()])
        report = self.report()
        warning = self.warning(report, LessonReadinessCode.STUDENT_WITHOUT_AVAILABILITY)
        self.assertEqual(len(report.warnings), len({item.code for item in report.warnings}))
        self.assertEqual({item.id for item in warning.subjects}, {student.id for student in students})
        self.assertEqual(warning.severity, LessonReadinessSeverity.WARNING)

    def test_instructor_without_availability_is_reported(self):
        self.repo.create_student("S-1", "Student", availability=[window()])
        instructor = self.repo.create_instructor("Instructor")
        warning = self.warning(self.report(), LessonReadinessCode.INSTRUCTOR_WITHOUT_AVAILABILITY)
        self.assertEqual([item.id for item in warning.subjects], [instructor.id])

    def test_required_tag_without_matching_instructor_includes_tag_name(self):
        tag = self.repo.create_tag("Violin")
        self.repo.create_student(
            "S-1", "Student", availability=[window()], tag_states={tag.id: LessonTagState.REQUIRED}
        )
        self.repo.create_instructor("Instructor", availability=[window()])
        warning = self.warning(self.report(), LessonReadinessCode.REQUIRED_TAG_WITHOUT_MATCH)
        self.assertEqual([(subject.id, subject.label) for subject in warning.subjects], [(tag.id, "VIOLIN")])

    def test_student_with_availability_but_no_feasible_candidate_is_reported(self):
        restricted = self.repo.create_tag("Restricted")
        student = self.repo.create_student("S-1", "Student", availability=[window()])
        self.repo.set_student_tag_states(student.id, {restricted.id: LessonTagState.RESTRICTED})
        self.repo.create_instructor(
            "Instructor",
            availability=[window()],
            tag_ids=[restricted.id],
        )
        report = self.report()
        warning = self.warning(report, LessonReadinessCode.STUDENT_WITHOUT_CANDIDATES)
        self.assertEqual([subject.id for subject in warning.subjects], [student.id])
        self.assertIn(LessonReadinessCode.NO_FEASIBLE_CANDIDATES, {item.code for item in report.warnings})

    def test_instructor_without_schedulable_lesson_is_reported(self):
        self.repo.create_instructor(
            "Instructor", lesson_length_minutes=60, availability=[window(start=480, end=510)]
        )
        report = self.report()
        warning = self.warning(report, LessonReadinessCode.INSTRUCTOR_WITHOUT_SCHEDULABLE_LESSONS)
        self.assertEqual(len(warning.subjects), 1)

    def test_adjacent_available_and_tentative_windows_can_fit_one_lesson(self):
        instructor = self.repo.create_instructor(
            "Instructor",
            lesson_length_minutes=60,
            availability=[
                window(start=480, end=510),
                window(start=510, end=540, status="Tentative"),
            ],
        )
        report = self.report()
        self.assertNotIn(
            instructor.id,
            {subject.id for warning in report.warnings
             if warning.code is LessonReadinessCode.INSTRUCTOR_WITHOUT_SCHEDULABLE_LESSONS
             for subject in warning.subjects},
        )

    def test_invalid_instructor_window_is_blocking(self):
        self.repo.create_instructor("Instructor", availability=[window()])
        scheduling_input = self.repo.get_scheduling_input()
        instructor = replace(
            scheduling_input.instructors[0],
            availability=(AvailabilityWindow("Funday", 600, 700, "Available"),),
        )
        invalid_input = replace(scheduling_input, instructors=(instructor,))
        candidates = generate_lesson_candidates(invalid_input)
        report = analyze_lesson_readiness(invalid_input, candidates)
        warning = self.warning(report, LessonReadinessCode.INVALID_INSTRUCTOR_WINDOW)
        self.assertEqual(warning.severity, LessonReadinessSeverity.ERROR)
        self.assertFalse(report.is_ready)

    def test_zero_capacity_is_blocking_readiness_input(self):
        student = self.repo.create_student("S-1", "Student", availability=[window()])
        self.repo.create_instructor("Instructor", availability=[window()])
        scheduling_input = self.repo.get_scheduling_input()
        instructors = (replace(scheduling_input.instructors[0], capacity=0),)
        invalid_input = replace(scheduling_input, instructors=instructors)
        candidates = generate_lesson_candidates(invalid_input)
        report = analyze_lesson_readiness(invalid_input, candidates)
        warning = self.warning(report, LessonReadinessCode.ZERO_CAPACITY)
        self.assertEqual(warning.subjects[0].kind.value, "instructor")
        self.assertFalse(report.is_ready)
        self.assertEqual(report.student_count, 1)
        self.assertEqual(report.students_with_candidates, 1)
        self.assertEqual(student.name, "Student")

    def test_ready_state_and_revision_stamping(self):
        student = self.repo.create_student("S-1", "Student", availability=[window()])
        instructor = self.repo.create_instructor(
            "Instructor", lesson_length_minutes=60, availability=[window()]
        )
        report = self.report()
        self.assertTrue(report.is_ready)
        self.assertEqual(report.source_data_revision, self.revision())
        self.assertEqual(report.feasible_candidate_count, 9)
        self.assertEqual(report.students_with_candidates, 1)
        self.assertEqual(report.instructors_with_schedulable_lessons, 1)
        self.assertEqual((report.student_count, report.instructor_count), (1, 1))
        self.assertTrue(student.id)
        self.assertTrue(instructor.id)

    def test_revision_mismatch_is_rejected(self):
        self.repo.create_student("S-1", "Student", availability=[window()])
        self.repo.create_instructor("Instructor", availability=[window()])
        scheduling_input = self.repo.get_scheduling_input()
        candidates = generate_lesson_candidates(scheduling_input)
        with self.assertRaisesRegex(ValueError, "revisions do not match"):
            analyze_lesson_readiness(
                scheduling_input,
                replace(candidates, source_data_revision=candidates.source_data_revision + 1),
            )

    def test_readiness_api_returns_grouped_warnings_counts_and_revision(self):
        tag = self.repo.create_tag("Violin")
        self.repo.create_student(
            "S-1", "Student", availability=[window()], tag_states={tag.id: LessonTagState.REQUIRED}
        )
        self.repo.create_instructor("Instructor", availability=[window()])
        self.db.commit()
        response = lesson.get_readiness(self.db)
        self.assertEqual(response.source_data_revision, self.revision())
        self.assertEqual((response.student_count, response.instructor_count), (1, 1))
        required_warning = next(
            warning for warning in response.warnings if warning.code == "REQUIRED_TAG_WITHOUT_MATCH"
        )
        self.assertEqual(required_warning.subjects[0].label, "VIOLIN")
        self.assertEqual(required_warning.subjects[0].kind, "tag")
        self.assertFalse(response.is_ready)


if __name__ == "__main__":
    unittest.main()
