"""Synthetic-data tests for Tutor Scheduling candidate generation."""

import unittest

from app.tutor_models import TutorTagState
from app.services.tutor_candidates import generate_tutor_candidates
from app.services.tutor_scheduling_models import (
    TutorCandidateIssueCode,
    TutorTutorInput,
    TutorSchedulingInput,
)
from tests.test_tutor_foundation import TutorTestCase, window


class TutorCandidateGenerationTests(TutorTestCase):
    def make_pair(self, *, length=30, capacity=1, tutor_windows=(), student_windows=(), tag_ids=(), states=None):
        tutor = self.repo.create_tutor(
            "Tutor", lesson_length_minutes=length, capacity=capacity, tag_ids=tag_ids
        )
        student = self.repo.create_student("S-1", "Student", tag_states=states or {})
        self.repo.set_tutor_availability(tutor.id, tutor_windows)
        self.repo.set_student_availability(student.id, student_windows)
        return student, tutor

    def test_candidates_require_intersection_and_exact_tutor_lesson_length(self):
        student, tutor = self.make_pair(
            length=30,
            tutor_windows=[window(start=540, end=630)],
            student_windows=[window(start=555, end=615)],
        )
        result = self.repo.generate_candidate_set()
        candidates = result.feasible_for_student(student.id)
        self.assertEqual(
            [(item.tutor_start_minute, item.tutor_end_minute) for item in candidates],
            [(555, 585), (570, 600), (585, 615)],
        )
        self.assertTrue(all(item.tutor_id == tutor.id for item in candidates))

    def test_required_tags_must_be_present_and_restricted_tags_must_be_absent(self):
        required = self.repo.create_tag("Required")
        restricted = self.repo.create_tag("Restricted")
        tutor = self.repo.create_tutor(
            "Tutor",
            lesson_length_minutes=30,
            tag_ids=[required.id, restricted.id],
            availability=[window(start=540, end=570)],
        )
        student = self.repo.create_student(
            "S-1",
            "Student",
            availability=[window(start=540, end=570)],
            tag_states={
                required.id: TutorTagState.REQUIRED,
                restricted.id: TutorTagState.RESTRICTED,
            },
        )
        candidates = self.repo.generate_candidate_set().for_student(student.id)
        self.assertEqual(len(candidates), 1)
        candidate = candidates[0]
        self.assertFalse(candidate.is_feasible)
        self.assertEqual(
            {issue.code for issue in candidate.issues},
            {TutorCandidateIssueCode.RESTRICTED_TAG_PRESENT},
        )
        self.assertEqual(candidate.issues[0].tag_ids, (restricted.id,))
        self.assertEqual(candidate.tutor_id, tutor.id)

        self.repo.set_tutor_tags(tutor.id, [restricted.id])
        updated = self.repo.generate_candidate_set().for_student(student.id)[0]
        self.assertEqual(
            {issue.code for issue in updated.issues},
            {
                TutorCandidateIssueCode.REQUIRED_TAG_MISSING,
                TutorCandidateIssueCode.RESTRICTED_TAG_PRESENT,
            },
        )

    def test_preferred_tags_count_without_affecting_feasibility(self):
        preferred = self.repo.create_tag("Preferred")
        tutor = self.repo.create_tutor(
            "Tutor",
            lesson_length_minutes=30,
            tag_ids=[preferred.id],
            availability=[window(start=540, end=570)],
        )
        student = self.repo.create_student(
            "S-1",
            "Student",
            availability=[window(start=540, end=570)],
            tag_states={preferred.id: TutorTagState.PREFERRED},
        )
        candidate = self.repo.generate_candidate_set().for_student(student.id)[0]
        self.assertTrue(candidate.is_feasible)
        self.assertEqual(candidate.preferred_tag_match_count, 1)
        self.assertEqual(candidate.issues, ())
        self.assertEqual(candidate.tutor_id, tutor.id)

    def test_tentative_minutes_sum_student_and_tutor_minutes(self):
        student, _ = self.make_pair(
            length=60,
            tutor_windows=[
                window(start=540, end=570, status="Tentative"),
                window(start=570, end=600),
            ],
            student_windows=[
                window(start=540, end=555, status="Tentative"),
                window(start=555, end=600),
            ],
        )
        candidate = self.repo.generate_candidate_set().feasible_for_student(student.id)[0]
        self.assertEqual(candidate.tentative_minutes, 45)

    def test_starts_are_generated_on_15_minute_boundaries(self):
        student, _ = self.make_pair(
            length=20,
            tutor_windows=[window(start=540, end=600)],
            student_windows=[window(start=540, end=600)],
        )
        candidates = self.repo.generate_candidate_set().feasible_for_student(student.id)
        self.assertEqual(
            [candidate.tutor_start_minute for candidate in candidates],
            [540, 555, 570],
        )
        self.assertTrue(all(candidate.tutor_start_minute % 15 == 0 for candidate in candidates))

    def test_tutor_without_a_legal_tutor_window_has_no_candidates(self):
        student, _ = self.make_pair(
            length=30,
            tutor_windows=[window(start=540, end=569)],
            student_windows=[window(start=540, end=600)],
        )
        self.assertEqual(self.repo.generate_candidate_set().for_student(student.id), ())

    def test_student_without_availability_has_no_legal_tutors(self):
        student, _ = self.make_pair(
            length=30,
            tutor_windows=[window(start=540, end=600)],
        )
        self.assertEqual(self.repo.generate_candidate_set().feasible_for_student(student.id), ())

    def test_capacity_and_overlapping_candidate_times_are_not_enforced(self):
        tutor = self.repo.create_tutor(
            "Tutor",
            lesson_length_minutes=60,
            capacity=1,
            availability=[window(start=540, end=600)],
        )
        students = [
            self.repo.create_student(
                student_id, name, availability=[window(start=540, end=600)]
            )
            for student_id, name in (("S-1", "One"), ("S-2", "Two"))
        ]
        candidates = self.repo.generate_candidate_set().feasible
        self.assertEqual(len(candidates), 2)
        self.assertEqual({item.tutor_id for item in candidates}, {tutor.id})
        self.assertEqual({item.tutor_start_minute for item in candidates}, {540})
        self.assertEqual({item.tutor_capacity for item in candidates}, {1})
        self.assertEqual({item.student_id for item in candidates}, {item.id for item in students})

    def test_candidate_order_is_deterministic(self):
        self.make_pair(
            length=30,
            tutor_windows=[window(start=540, end=600)],
            student_windows=[window(start=540, end=600)],
        )
        self.repo.create_tutor(
            "Second Tutor",
            lesson_length_minutes=30,
            availability=[window(start=540, end=600)],
        )
        self.repo.create_student(
            "S-2",
            "Second Student",
            availability=[window(start=540, end=600)],
        )
        first = self.repo.generate_candidate_set().candidates
        second = generate_tutor_candidates(self.repo.get_scheduling_input()).candidates
        self.assertEqual(first, second)
        self.assertEqual(
            [(item.student_id, item.tutor_id, item.day, item.tutor_start_minute) for item in first],
            sorted((item.student_id, item.tutor_id, item.day, item.tutor_start_minute) for item in first),
        )

    def test_repository_snapshot_and_candidate_set_are_revision_stamped(self):
        student, tutor = self.make_pair(
            length=30,
            tutor_windows=[window(start=540, end=600)],
            student_windows=[window(start=540, end=600)],
        )
        snapshot = self.repo.get_scheduling_input()
        self.assertEqual(snapshot.source_data_revision, self.revision())
        self.assertEqual({item.id for item in snapshot.students}, {student.id})
        self.assertEqual({item.id for item in snapshot.tutors}, {tutor.id})
        before = self.revision()
        candidate_set = self.repo.generate_candidate_set()
        self.assertEqual(candidate_set.source_data_revision, before)
        self.assertEqual(self.revision(), before)
        self.repo.set_student_availability(student.id, [window(start=540, end=570)])
        self.assertGreater(self.repo.generate_candidate_set().source_data_revision, before)

    def test_generator_is_pure_and_empty_students_produce_no_candidates(self):
        tutor = TutorTutorInput(
            id="tutor",
            name="Tutor",
            lesson_length_minutes=30,
            capacity=1,
            tag_ids=frozenset(),
            availability=(window(start=540, end=600),),
        )
        scheduling_input = TutorSchedulingInput(
            source_data_revision=17,
            students=(),
            tutors=(tutor,),
        )
        self.assertEqual(generate_tutor_candidates(scheduling_input).candidates, ())
        self.assertEqual(scheduling_input.source_data_revision, 17)


if __name__ == "__main__":
    unittest.main()
