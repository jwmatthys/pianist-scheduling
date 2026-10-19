"""Pure readiness analysis for Lesson Scheduling."""

from __future__ import annotations

from collections import defaultdict

from .availability import AvailabilityWindow, WEEKDAYS, normalize_windows, validate_window
from .lesson_scheduling_models import (
    MINUTES_PER_DAY,
    LessonCandidateSet,
    LessonInstructorInput,
    LessonReadinessCategory,
    LessonReadinessCode,
    LessonReadinessReport,
    LessonReadinessSeverity,
    LessonReadinessSubject,
    LessonReadinessSubjectKind,
    LessonReadinessWarning,
    LessonSchedulingInput,
)
from ..lesson_models import LessonTagState

_ERROR = LessonReadinessSeverity.ERROR
_WARNING = LessonReadinessSeverity.WARNING


def _availability_is_invalid(windows: tuple[AvailabilityWindow, ...]) -> bool:
    if any(validate_window(window) is not None for window in windows):
        return True
    _, issues = normalize_windows(list(windows))
    return any(issue.severity == "error" for issue in issues)


def _instructor_has_schedulable_lesson(instructor: LessonInstructorInput) -> bool:
    length = instructor.lesson_length_minutes
    if length <= 0 or length > MINUTES_PER_DAY:
        return False

    available: dict[str, list[bool]] = {
        day: [False] * MINUTES_PER_DAY for day in WEEKDAYS
    }
    tentative: dict[str, list[bool]] = {
        day: [False] * MINUTES_PER_DAY for day in WEEKDAYS
    }
    blocked: dict[str, list[bool]] = {
        day: [False] * MINUTES_PER_DAY for day in WEEKDAYS
    }
    for window in instructor.availability:
        if validate_window(window) is not None:
            continue
        marks = {
            "Available": available,
            "Tentative": tentative,
            "Unavailable": blocked,
        }[window.status]
        for minute in range(window.start_minute, window.end_minute):
            marks[window.day][minute] = True

    for day in WEEKDAYS:
        legal = [
            (available[day][minute] or tentative[day][minute]) and not blocked[day][minute]
            for minute in range(MINUTES_PER_DAY)
        ]
        prefix = [0]
        for minute_is_legal in legal:
            prefix.append(prefix[-1] + int(minute_is_legal))
        for start in range(0, MINUTES_PER_DAY - length + 1, 15):
            if prefix[start + length] - prefix[start] == length:
                return True
    return False


def _required_tag_subjects(
    scheduling_input: LessonSchedulingInput,
) -> tuple[LessonReadinessSubject, ...]:
    tag_names = {tag.id: tag.name for tag in scheduling_input.tags}
    instructor_tags = {tag_id for instructor in scheduling_input.instructors for tag_id in instructor.tag_ids}
    required = {
        tag_id
        for student in scheduling_input.students
        for tag_id, state in student.tag_states
        if state is LessonTagState.REQUIRED
    }
    return tuple(
        LessonReadinessSubject(
            kind=LessonReadinessSubjectKind.TAG,
            id=tag_id,
            label=tag_names.get(tag_id, tag_id),
        )
        for tag_id in sorted(required)
        if tag_id not in instructor_tags
    )


def analyze_lesson_readiness(
    scheduling_input: LessonSchedulingInput,
    candidates: LessonCandidateSet,
) -> LessonReadinessReport:
    """Evaluate readiness using the supplied snapshot and candidates without regenerating them."""
    if scheduling_input.source_data_revision != candidates.source_data_revision:
        raise ValueError("Lesson scheduling input and candidate revisions do not match.")

    grouped: dict[
        tuple[LessonReadinessCode, LessonReadinessSeverity, LessonReadinessCategory],
        list[LessonReadinessSubject],
    ] = defaultdict(list)

    def add(
        code: LessonReadinessCode,
        severity: LessonReadinessSeverity,
        category: LessonReadinessCategory,
        subject: LessonReadinessSubject | None = None,
    ) -> None:
        if subject is not None:
            grouped[(code, severity, category)].append(subject)
        else:
            grouped[(code, severity, category)]

    students = scheduling_input.students
    instructors = scheduling_input.instructors
    if not students:
        add(LessonReadinessCode.NO_STUDENTS, _ERROR, LessonReadinessCategory.DATA)
    if not instructors:
        add(LessonReadinessCode.NO_INSTRUCTORS, _ERROR, LessonReadinessCategory.DATA)

    student_candidates: dict[str, list] = defaultdict(list)
    for candidate in candidates.candidates:
        student_candidates[candidate.student_id].append(candidate)

    for student in students:
        if not student.availability:
            add(
                LessonReadinessCode.STUDENT_WITHOUT_AVAILABILITY,
                _WARNING,
                LessonReadinessCategory.STUDENTS,
                LessonReadinessSubject(LessonReadinessSubjectKind.STUDENT, student.id, student.name),
            )
        elif instructors and not any(
            candidate.is_feasible for candidate in student_candidates.get(student.id, ())
        ):
            add(
                LessonReadinessCode.STUDENT_WITHOUT_CANDIDATES,
                _WARNING,
                LessonReadinessCategory.STUDENTS,
                LessonReadinessSubject(LessonReadinessSubjectKind.STUDENT, student.id, student.name),
            )

    schedulable_instructors: set[str] = set()
    for instructor in instructors:
        subject = LessonReadinessSubject(
            LessonReadinessSubjectKind.INSTRUCTOR, instructor.id, instructor.name
        )
        if not instructor.availability:
            add(
                LessonReadinessCode.INSTRUCTOR_WITHOUT_AVAILABILITY,
                _WARNING,
                LessonReadinessCategory.INSTRUCTORS,
                subject,
            )
        if instructor.capacity <= 0:
            add(
                LessonReadinessCode.ZERO_CAPACITY,
                _ERROR,
                LessonReadinessCategory.INSTRUCTORS,
                subject,
            )
        if _availability_is_invalid(instructor.availability):
            add(
                LessonReadinessCode.INVALID_INSTRUCTOR_WINDOW,
                _ERROR,
                LessonReadinessCategory.INSTRUCTORS,
                subject,
            )
        if _instructor_has_schedulable_lesson(instructor):
            schedulable_instructors.add(instructor.id)
        elif instructor.availability:
            add(
                LessonReadinessCode.INSTRUCTOR_WITHOUT_SCHEDULABLE_LESSONS,
                _WARNING,
                LessonReadinessCategory.INSTRUCTORS,
                subject,
            )

    required_tags_without_match = _required_tag_subjects(scheduling_input)
    if required_tags_without_match:
        grouped[
            (
                LessonReadinessCode.REQUIRED_TAG_WITHOUT_MATCH,
                _WARNING,
                LessonReadinessCategory.TAGS,
            )
        ].extend(required_tags_without_match)

    feasible_candidate_count = sum(candidate.is_feasible for candidate in candidates.candidates)
    if students and instructors and feasible_candidate_count == 0:
        add(
            LessonReadinessCode.NO_FEASIBLE_CANDIDATES,
            _ERROR,
            LessonReadinessCategory.DATA,
        )

    ordered_groups = sorted(
        grouped.items(),
        key=lambda item: (
            item[0][2].value,
            item[0][0].value,
        ),
    )
    warnings = tuple(
        LessonReadinessWarning(
            code=code,
            severity=severity,
            category=category,
            subjects=tuple(sorted(subjects, key=lambda item: (item.label.casefold(), item.id))),
        )
        for (code, severity, category), subjects in ordered_groups
    )

    return LessonReadinessReport(
        source_data_revision=scheduling_input.source_data_revision,
        is_ready=(
            feasible_candidate_count > 0
            and not any(warning.severity is LessonReadinessSeverity.ERROR for warning in warnings)
        ),
        student_count=len(students),
        instructor_count=len(instructors),
        feasible_candidate_count=feasible_candidate_count,
        total_candidate_count=len(candidates.candidates),
        students_with_candidates=len({
            candidate.student_id for candidate in candidates.candidates if candidate.is_feasible
        }),
        instructors_with_schedulable_lessons=len(schedulable_instructors),
        warnings=warnings,
    )
