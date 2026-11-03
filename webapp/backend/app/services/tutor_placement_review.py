"""Current warnings and placement verification for the Tutor schedule."""

from __future__ import annotations

from collections import Counter, defaultdict
from typing import Sequence

from ..tutor_domain import TutorPlacementType, TutorTagState, TutorUnplacedReason
from .availability import WEEKDAYS, validate_window
from .tutor_scheduling_models import (
    TutorCandidateSet,
    TutorPlacementVerification,
    TutorPlacementWarning,
    TutorPlacementWarningCode,
    TutorScheduleEntry,
    TutorSchedulingInput,
    TutorStoredAssignment,
    TutorStoredUnplaced,
)
from .tutor_optimizer import _schedulable_tutor_capacity


def _window_status(windows, day: str, start: int, end: int) -> tuple[bool, int]:
    covered_status: list[str | None] = []
    for minute in range(start, end):
        status = None
        for window in windows:
            if (
                validate_window(window) is None
                and window.day == day
                and window.start_minute <= minute < window.end_minute
                and window.status in {"Available", "Tentative"}
            ):
                status = window.status
                break
        covered_status.append(status)
    tentative = sum(status == "Tentative" for status in covered_status)
    return all(status is not None for status in covered_status), tentative


def review_tutor_schedule(
    scheduling_input: TutorSchedulingInput,
    candidates: TutorCandidateSet,
    assignments: Sequence[TutorStoredAssignment],
    unplaced: Sequence[TutorStoredUnplaced],
) -> tuple[tuple[TutorScheduleEntry, ...], tuple[TutorPlacementVerification, ...]]:
    """Build current schedule rows and feasibility warnings without changing persistence."""
    students = {item.id: item for item in scheduling_input.students}
    tutors = {item.id: item for item in scheduling_input.tutors}
    tags = {item.id: item.name for item in scheduling_input.tags}
    assignment_by_student = {item.student_id: item for item in assignments}
    unplaced_by_student = {item.student_id: item for item in unplaced}
    tutor_counts = Counter(item.tutor_id for item in assignments)
    by_tutor_day: dict[tuple[str, str], list[TutorStoredAssignment]] = defaultdict(list)
    for assignment in assignments:
        by_tutor_day[(assignment.tutor_id, assignment.day)].append(assignment)

    overlap_students: set[str] = set()
    for items in by_tutor_day.values():
        ordered = sorted(items, key=lambda item: (item.start_minute, item.end_minute, item.student_id))
        for index, current in enumerate(ordered):
            for following in ordered[index + 1:]:
                if following.start_minute >= current.end_minute:
                    break
                overlap_students.update((current.student_id, following.student_id))

    by_person_day: dict[tuple[str, str], list[TutorStoredAssignment]] = defaultdict(list)
    for assignment in assignments:
        known = students.get(assignment.student_id)
        by_person_day[((known.person_key if known else "") or assignment.student_id, assignment.day)].append(assignment)
    person_overlap_students: set[str] = set()
    for items in by_person_day.values():
        ordered = sorted(items, key=lambda item: (item.start_minute, item.end_minute, item.student_id))
        for index, current in enumerate(ordered):
            for following in ordered[index + 1:]:
                if following.start_minute >= current.end_minute:
                    break
                person_overlap_students.update((current.student_id, following.student_id))

    feasible_starts = {
        (item.student_id, item.tutor_id, item.day, item.tutor_start_minute)
        for item in candidates.feasible
    }
    entries: list[TutorScheduleEntry] = []
    verification: list[TutorPlacementVerification] = []
    all_student_ids = set(students) | set(assignment_by_student) | set(unplaced_by_student)
    ordered_student_ids = sorted(
        all_student_ids,
        key=lambda key: (
            (students[key].name if key in students else assignment_by_student.get(key, None).student_name
             if key in assignment_by_student else unplaced_by_student[key].student_name).casefold(),
            key,
        ),
    )

    for student_id in ordered_student_ids:
        student = students.get(student_id)
        assignment = assignment_by_student.get(student_id)
        unplaced_item = unplaced_by_student.get(student_id)
        student_name = (
            student.name if student is not None
            else assignment.student_name if assignment is not None
            else unplaced_item.student_name if unplaced_item is not None
            else ""
        )
        if assignment is None:
            kind = unplaced_item.placement_type if unplaced_item is not None else None
            reason = (
                unplaced_item.reason if unplaced_item is not None
                else TutorUnplacedReason.NO_FEASIBLE_CANDIDATE.value
            )
            entries.append(TutorScheduleEntry(
                student_id=student_id,
                student_name=student_name,
                placement_type=kind,
                tutor_id=None,
                tutor_name=None,
                day=None,
                start_minute=None,
                end_minute=None,
                preferred_tag_match_count=0,
                tentative_minutes=0,
                verified=None,
                warnings=(),
                unplaced_reason=reason,
            ))
            continue

        tutor = tutors.get(assignment.tutor_id)
        tutor_name = (
            tutor.name if tutor is not None else assignment.tutor_name
        )
        warnings: list[TutorPlacementWarning] = []
        if student is None:
            warnings.append(TutorPlacementWarning(
                TutorPlacementWarningCode.STUDENT_REMOVED,
                "The assigned student is no longer in the current roster.",
            ))
        if tutor is None:
            warnings.append(TutorPlacementWarning(
                TutorPlacementWarningCode.TUTOR_REMOVED,
                "The assigned tutor is no longer in the current roster.",
            ))
        elif assignment.end_minute - assignment.start_minute != tutor.lesson_length_minutes:
            warnings.append(TutorPlacementWarning(
                TutorPlacementWarningCode.LESSON_LENGTH_CHANGED,
                "The assigned session duration differs from the tutor's current lesson length.",
            ))
        if (
            assignment.day not in WEEKDAYS
            or not 0 <= assignment.start_minute < assignment.end_minute <= 24 * 60
        ):
            warnings.append(TutorPlacementWarning(
                TutorPlacementWarningCode.INVALID_TIME,
                "The placement has an invalid day or time.",
            ))
        else:
            if student is not None:
                covered, tentative = _window_status(
                    student.availability, assignment.day, assignment.start_minute, assignment.end_minute
                )
                if not covered:
                    warnings.append(TutorPlacementWarning(
                        TutorPlacementWarningCode.OUTSIDE_STUDENT_AVAILABILITY,
                        "The session is outside the student's current Availability Windows.",
                    ))
                elif tentative:
                    warnings.append(TutorPlacementWarning(
                        TutorPlacementWarningCode.TENTATIVE_AVAILABILITY,
                        f"The session overlaps {tentative} tentative student availability minutes.",
                    ))
            else:
                tentative = 0
            if tutor is not None:
                covered, tutor_tentative = _window_status(
                    tutor.availability, assignment.day, assignment.start_minute, assignment.end_minute
                )
                if not covered:
                    warnings.append(TutorPlacementWarning(
                        TutorPlacementWarningCode.OUTSIDE_TUTOR_AVAILABILITY,
                        "The session is outside the tutor's current Availability Windows.",
                    ))
                elif tutor_tentative:
                    warnings.append(TutorPlacementWarning(
                        TutorPlacementWarningCode.TENTATIVE_AVAILABILITY,
                        f"The session overlaps {tutor_tentative} tentative tutor availability minutes.",
                    ))

        preferred_count = assignment.preferred_tag_match_count
        tentative_minutes = assignment.tentative_minutes
        if student is not None and tutor is not None:
            student_tag_states = dict(student.tag_states)
            required_missing = sorted(
                tags[tag_id]
                for tag_id, state in student_tag_states.items()
                if state is TutorTagState.REQUIRED
                and tag_id in tags
                and tag_id not in tutor.tag_ids
            )
            restricted_present = sorted(
                tags[tag_id]
                for tag_id, state in student_tag_states.items()
                if state is TutorTagState.RESTRICTED
                and tag_id in tags
                and tag_id in tutor.tag_ids
            )
            if required_missing:
                warnings.append(TutorPlacementWarning(
                    TutorPlacementWarningCode.REQUIRED_TAG_MISSING,
                    "Missing Required tags: " + ", ".join(required_missing) + ".",
                ))
            if restricted_present:
                warnings.append(TutorPlacementWarning(
                    TutorPlacementWarningCode.RESTRICTED_TAG_CONFLICT,
                    "Restricted tag conflict: " + ", ".join(restricted_present) + ".",
                ))
            preferred_count = sum(
                state is TutorTagState.PREFERRED and tag_id in tutor.tag_ids
                for tag_id, state in student_tag_states.items()
            )
            if assignment.placement_type == TutorPlacementType.AUTOMATIC.value and (
                student.id, tutor.id, assignment.day, assignment.start_minute
            ) not in feasible_starts:
                warnings.append(TutorPlacementWarning(
                    TutorPlacementWarningCode.NOT_FEASIBLE_CANDIDATE,
                    "The automatic placement is no longer a feasible candidate.",
                ))

        tutor_assignment_count = tutor_counts[assignment.tutor_id]
        if tutor is not None and tutor_assignment_count > _schedulable_tutor_capacity(tutor):
            warnings.append(TutorPlacementWarning(
                TutorPlacementWarningCode.TUTOR_CAPACITY_EXCEEDED,
                f"{tutor_assignment_count} sessions exceed the tutor's effective capacity "
                f"of {_schedulable_tutor_capacity(tutor)}.",
            ))
        if assignment.student_id in overlap_students:
            warnings.append(TutorPlacementWarning(
                TutorPlacementWarningCode.TUTOR_OVERLAP,
                "This session overlaps another session assigned to the tutor.",
            ))
        if assignment.student_id in person_overlap_students:
            warnings.append(TutorPlacementWarning(
                TutorPlacementWarningCode.STUDENT_OVERLAP,
                "This session overlaps another session for the same student.",
            ))
        warning_tuple = tuple(warnings)
        blocking = any(
            item.code is not TutorPlacementWarningCode.TENTATIVE_AVAILABILITY
            for item in warning_tuple
        )
        entries.append(TutorScheduleEntry(
            student_id=student_id,
            student_name=student_name,
            placement_type=assignment.placement_type,
            tutor_id=assignment.tutor_id,
            tutor_name=tutor_name,
            day=assignment.day,
            start_minute=assignment.start_minute,
            end_minute=assignment.end_minute,
            preferred_tag_match_count=preferred_count,
            tentative_minutes=tentative_minutes,
            verified=not blocking,
            warnings=warning_tuple,
            unplaced_reason=None,
        ))
        if blocking:
            verification.append(TutorPlacementVerification(
                student_id=student_id,
                student_name=student_name,
                tutor_id=assignment.tutor_id,
                tutor_name=tutor_name,
                day=assignment.day,
                start_minute=assignment.start_minute,
                warnings=warning_tuple,
            ))

    return tuple(entries), tuple(verification)
