"""Pure readiness analysis for Tutor Scheduling."""

from __future__ import annotations

from collections import defaultdict

from .availability import AvailabilityWindow, WEEKDAYS, normalize_windows, validate_window
from .tutor_scheduling_models import (
    MINUTES_PER_DAY,
    TutorCandidateSet,
    TutorTutorInput,
    TutorReadinessCategory,
    TutorReadinessCode,
    TutorReadinessReport,
    TutorReadinessSeverity,
    TutorReadinessSubject,
    TutorReadinessSubjectKind,
    TutorReadinessWarning,
    TutorSchedulingInput,
)
from ..tutor_domain import TutorTagState

_ERROR = TutorReadinessSeverity.ERROR
_WARNING = TutorReadinessSeverity.WARNING


def _availability_is_invalid(windows: tuple[AvailabilityWindow, ...]) -> bool:
    if any(validate_window(window) is not None for window in windows):
        return True
    _, issues = normalize_windows(list(windows))
    return any(issue.severity == "error" for issue in issues)


def _tutor_has_schedulable_tutor(tutor: TutorTutorInput) -> bool:
    length = tutor.lesson_length_minutes
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
    for window in tutor.availability:
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
    scheduling_input: TutorSchedulingInput,
) -> tuple[TutorReadinessSubject, ...]:
    tag_names = {tag.id: tag.name for tag in scheduling_input.tags}
    tutor_tags = {tag_id for tutor in scheduling_input.tutors for tag_id in tutor.tag_ids}
    required = {
        tag_id
        for student in scheduling_input.students
        for tag_id, state in student.tag_states
        if state is TutorTagState.REQUIRED
    }
    return tuple(
        TutorReadinessSubject(
            kind=TutorReadinessSubjectKind.TAG,
            id=tag_id,
            label=tag_names.get(tag_id, tag_id),
        )
        for tag_id in sorted(required)
        if tag_id not in tutor_tags
    )


def analyze_tutor_readiness(
    scheduling_input: TutorSchedulingInput,
    candidates: TutorCandidateSet,
) -> TutorReadinessReport:
    """Evaluate readiness using the supplied snapshot and candidates without regenerating them."""
    if scheduling_input.source_data_revision != candidates.source_data_revision:
        raise ValueError("Tutor scheduling input and candidate revisions do not match.")

    grouped: dict[
        tuple[TutorReadinessCode, TutorReadinessSeverity, TutorReadinessCategory],
        list[TutorReadinessSubject],
    ] = defaultdict(list)

    def add(
        code: TutorReadinessCode,
        severity: TutorReadinessSeverity,
        category: TutorReadinessCategory,
        subject: TutorReadinessSubject | None = None,
    ) -> None:
        if subject is not None:
            grouped[(code, severity, category)].append(subject)
        else:
            grouped[(code, severity, category)]

    students = scheduling_input.students
    tutors = scheduling_input.tutors
    if not students:
        add(TutorReadinessCode.NO_STUDENTS, _ERROR, TutorReadinessCategory.DATA)
    if not tutors:
        add(TutorReadinessCode.NO_TUTORS, _ERROR, TutorReadinessCategory.DATA)

    student_candidates: dict[str, list] = defaultdict(list)
    for candidate in candidates.candidates:
        student_candidates[candidate.student_id].append(candidate)

    for student in students:
        if not student.availability:
            add(
                TutorReadinessCode.STUDENT_WITHOUT_AVAILABILITY,
                _WARNING,
                TutorReadinessCategory.STUDENTS,
                TutorReadinessSubject(TutorReadinessSubjectKind.STUDENT, student.id, student.name),
            )
        elif tutors and not any(
            candidate.is_feasible for candidate in student_candidates.get(student.id, ())
        ):
            add(
                TutorReadinessCode.STUDENT_WITHOUT_CANDIDATES,
                _WARNING,
                TutorReadinessCategory.STUDENTS,
                TutorReadinessSubject(TutorReadinessSubjectKind.STUDENT, student.id, student.name),
            )

    schedulable_tutors: set[str] = set()
    for tutor in tutors:
        subject = TutorReadinessSubject(
            TutorReadinessSubjectKind.TUTOR, tutor.id, tutor.name
        )
        if not tutor.availability:
            add(
                TutorReadinessCode.TUTOR_WITHOUT_AVAILABILITY,
                _WARNING,
                TutorReadinessCategory.TUTORS,
                subject,
            )
        if tutor.capacity <= 0:
            add(
                TutorReadinessCode.ZERO_CAPACITY,
                _ERROR,
                TutorReadinessCategory.TUTORS,
                subject,
            )
        if _availability_is_invalid(tutor.availability):
            add(
                TutorReadinessCode.INVALID_TUTOR_WINDOW,
                _ERROR,
                TutorReadinessCategory.TUTORS,
                subject,
            )
        if _tutor_has_schedulable_tutor(tutor):
            schedulable_tutors.add(tutor.id)
        elif tutor.availability:
            add(
                TutorReadinessCode.TUTOR_WITHOUT_SCHEDULABLE_TUTORS,
                _WARNING,
                TutorReadinessCategory.TUTORS,
                subject,
            )

    required_tags_without_match = _required_tag_subjects(scheduling_input)
    if required_tags_without_match:
        grouped[
            (
                TutorReadinessCode.REQUIRED_TAG_WITHOUT_MATCH,
                _WARNING,
                TutorReadinessCategory.TAGS,
            )
        ].extend(required_tags_without_match)

    feasible_candidate_count = sum(candidate.is_feasible for candidate in candidates.candidates)
    if students and tutors and feasible_candidate_count == 0:
        add(
            TutorReadinessCode.NO_FEASIBLE_CANDIDATES,
            _ERROR,
            TutorReadinessCategory.DATA,
        )

    ordered_groups = sorted(
        grouped.items(),
        key=lambda item: (
            item[0][2].value,
            item[0][0].value,
        ),
    )
    warnings = tuple(
        TutorReadinessWarning(
            code=code,
            severity=severity,
            category=category,
            subjects=tuple(sorted(subjects, key=lambda item: (item.label.casefold(), item.id))),
        )
        for (code, severity, category), subjects in ordered_groups
    )

    return TutorReadinessReport(
        source_data_revision=scheduling_input.source_data_revision,
        is_ready=(
            feasible_candidate_count > 0
            and not any(warning.severity is TutorReadinessSeverity.ERROR for warning in warnings)
        ),
        student_count=len(students),
        tutor_count=len(tutors),
        feasible_candidate_count=feasible_candidate_count,
        total_candidate_count=len(candidates.candidates),
        students_with_candidates=len({
            candidate.student_id for candidate in candidates.candidates if candidate.is_feasible
        }),
        tutors_with_schedulable_tutors=len(schedulable_tutors),
        warnings=warnings,
    )
