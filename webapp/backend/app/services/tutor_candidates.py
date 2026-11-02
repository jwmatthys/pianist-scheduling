"""Pure candidate generation for Tutor Scheduling."""

from __future__ import annotations

from .availability import AvailabilityWindow, WEEKDAYS
from .tutor_scheduling_models import (
    MINUTES_PER_DAY,
    TutorCandidate,
    TutorCandidateIssue,
    TutorCandidateIssueCode,
    TutorCandidateSet,
    TutorTutorInput,
    TutorSchedulingInput,
    TutorStudentInput,
)
from ..tutor_domain import TutorTagState


class _DayCoverage:
    """Per-minute weekly availability coverage with constant-time interval queries."""

    def __init__(self, windows: tuple[AvailabilityWindow, ...]):
        available = [False] * MINUTES_PER_DAY
        tentative = [False] * MINUTES_PER_DAY
        blocked = [False] * MINUTES_PER_DAY
        for window in windows:
            marks = {
                "Available": available,
                "Tentative": tentative,
                "Unavailable": blocked,
            }.get(window.status)
            if marks is None or window.day not in WEEKDAYS:
                continue
            start = max(0, window.start_minute)
            end = min(MINUTES_PER_DAY, window.end_minute)
            for minute in range(start, max(start, end)):
                marks[minute] = True

        self._usable_prefix = [0]
        self._tentative_prefix = [0]
        for minute in range(MINUTES_PER_DAY):
            usable = (available[minute] or tentative[minute]) and not blocked[minute]
            tentative_minute = usable and not available[minute]
            self._usable_prefix.append(self._usable_prefix[-1] + int(usable))
            self._tentative_prefix.append(self._tentative_prefix[-1] + int(tentative_minute))

    def tentative_minutes_if_covered(self, start: int, end: int) -> int | None:
        if start < 0 or end > MINUTES_PER_DAY or end <= start:
            return None
        covered = self._usable_prefix[end] - self._usable_prefix[start]
        if covered != end - start:
            return None
        return self._tentative_prefix[end] - self._tentative_prefix[start]


def _coverage_by_day(windows: tuple[AvailabilityWindow, ...]) -> dict[str, _DayCoverage]:
    by_day: dict[str, list[AvailabilityWindow]] = {}
    for window in windows:
        by_day.setdefault(window.day, []).append(window)
    return {day: _DayCoverage(tuple(day_windows)) for day, day_windows in by_day.items()}


def _tag_issues(
    student: TutorStudentInput, tutor: TutorTutorInput
) -> tuple[tuple[TutorCandidateIssue, ...], int]:
    required_missing = tuple(sorted(
        tag_id
        for tag_id, state in student.tag_states
        if state is TutorTagState.REQUIRED and tag_id not in tutor.tag_ids
    ))
    restricted_present = tuple(sorted(
        tag_id
        for tag_id, state in student.tag_states
        if state is TutorTagState.RESTRICTED and tag_id in tutor.tag_ids
    ))
    preferred_matches = sum(
        1
        for tag_id, state in student.tag_states
        if state is TutorTagState.PREFERRED and tag_id in tutor.tag_ids
    )

    issues: list[TutorCandidateIssue] = []
    if required_missing:
        issues.append(TutorCandidateIssue(
            TutorCandidateIssueCode.REQUIRED_TAG_MISSING, required_missing
        ))
    if restricted_present:
        issues.append(TutorCandidateIssue(
            TutorCandidateIssueCode.RESTRICTED_TAG_PRESENT, restricted_present
        ))
    return tuple(issues), preferred_matches


def generate_tutor_candidates(scheduling_input: TutorSchedulingInput) -> TutorCandidateSet:
    """Return all student/tutor/start combinations covered by both Availability Windows."""
    students = sorted(scheduling_input.students, key=lambda item: item.id)
    tutors = sorted(scheduling_input.tutors, key=lambda item: item.id)
    student_coverage = {student.id: _coverage_by_day(student.availability) for student in students}
    tutor_coverage = {tutor.id: _coverage_by_day(tutor.availability) for tutor in tutors}
    candidates: list[TutorCandidate] = []

    for student in students:
        for tutor in tutors:
            length = tutor.lesson_length_minutes
            if not 0 < length <= MINUTES_PER_DAY:
                continue
            tag_issues, preferred_matches = _tag_issues(student, tutor)
            student_days = student_coverage[student.id]
            tutor_days = tutor_coverage[tutor.id]
            for day in WEEKDAYS:
                student_day = student_days.get(day)
                tutor_day = tutor_days.get(day)
                if student_day is None or tutor_day is None:
                    continue
                last_start = MINUTES_PER_DAY - length
                for start in range(0, last_start + 1, 15):
                    end = start + length
                    student_tentative = student_day.tentative_minutes_if_covered(start, end)
                    if student_tentative is None:
                        continue
                    tutor_tentative = tutor_day.tentative_minutes_if_covered(start, end)
                    if tutor_tentative is None:
                        continue
                    candidates.append(TutorCandidate(
                        student_id=student.id,
                        tutor_id=tutor.id,
                        day=day,
                        tutor_start_minute=start,
                        tutor_end_minute=end,
                        preferred_tag_match_count=preferred_matches,
                        tentative_minutes=student_tentative + tutor_tentative,
                        issues=tag_issues,
                        is_feasible=not tag_issues,
                        tutor_capacity=tutor.capacity,
                    ))

    return TutorCandidateSet(
        source_data_revision=scheduling_input.source_data_revision,
        candidates=tuple(candidates),
    )
