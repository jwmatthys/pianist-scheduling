"""Current warnings and placement verification for the Lesson schedule."""

from __future__ import annotations

from collections import Counter, defaultdict
from typing import Sequence

from ..lesson_models import LessonPlacementType, LessonTagState, LessonUnplacedReason
from .availability import WEEKDAYS, validate_window
from .lesson_scheduling_models import (
    LessonCandidateSet,
    LessonPlacementVerification,
    LessonPlacementWarning,
    LessonPlacementWarningCode,
    LessonScheduleEntry,
    LessonSchedulingInput,
    LessonStoredAssignment,
    LessonStoredUnplaced,
)
from .lesson_optimizer import _schedulable_lesson_capacity


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


def review_lesson_schedule(
    scheduling_input: LessonSchedulingInput,
    candidates: LessonCandidateSet,
    assignments: Sequence[LessonStoredAssignment],
    unplaced: Sequence[LessonStoredUnplaced],
) -> tuple[tuple[LessonScheduleEntry, ...], tuple[LessonPlacementVerification, ...]]:
    """Build current schedule rows and feasibility warnings without changing persistence."""
    students = {item.id: item for item in scheduling_input.students}
    instructors = {item.id: item for item in scheduling_input.instructors}
    tags = {item.id: item.name for item in scheduling_input.tags}
    assignment_by_student = {item.student_id: item for item in assignments}
    unplaced_by_student = {item.student_id: item for item in unplaced}
    instructor_counts = Counter(item.instructor_id for item in assignments)
    by_instructor_day: dict[tuple[str, str], list[LessonStoredAssignment]] = defaultdict(list)
    for assignment in assignments:
        by_instructor_day[(assignment.instructor_id, assignment.day)].append(assignment)

    overlap_students: set[str] = set()
    for items in by_instructor_day.values():
        ordered = sorted(items, key=lambda item: (item.start_minute, item.end_minute, item.student_id))
        for index, current in enumerate(ordered):
            for following in ordered[index + 1:]:
                if following.start_minute >= current.end_minute:
                    break
                overlap_students.update((current.student_id, following.student_id))

    feasible_starts = {
        (item.student_id, item.instructor_id, item.day, item.lesson_start_minute)
        for item in candidates.feasible
    }
    entries: list[LessonScheduleEntry] = []
    verification: list[LessonPlacementVerification] = []
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
                else LessonUnplacedReason.NO_FEASIBLE_CANDIDATE.value
            )
            entries.append(LessonScheduleEntry(
                student_id=student_id,
                student_name=student_name,
                placement_type=kind,
                instructor_id=None,
                instructor_name=None,
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

        instructor = instructors.get(assignment.instructor_id)
        instructor_name = (
            instructor.name if instructor is not None else assignment.instructor_name
        )
        warnings: list[LessonPlacementWarning] = []
        if student is None:
            warnings.append(LessonPlacementWarning(
                LessonPlacementWarningCode.STUDENT_REMOVED,
                "The assigned student is no longer in the current roster.",
            ))
        if instructor is None:
            warnings.append(LessonPlacementWarning(
                LessonPlacementWarningCode.INSTRUCTOR_REMOVED,
                "The assigned instructor is no longer in the current roster.",
            ))
        elif assignment.end_minute - assignment.start_minute != instructor.lesson_length_minutes:
            warnings.append(LessonPlacementWarning(
                LessonPlacementWarningCode.LESSON_LENGTH_CHANGED,
                "The assigned lesson duration differs from the instructor's current lesson length.",
            ))
        if (
            assignment.day not in WEEKDAYS
            or not 0 <= assignment.start_minute < assignment.end_minute <= 24 * 60
        ):
            warnings.append(LessonPlacementWarning(
                LessonPlacementWarningCode.INVALID_TIME,
                "The placement has an invalid day or time.",
            ))
        else:
            if student is not None:
                covered, tentative = _window_status(
                    student.availability, assignment.day, assignment.start_minute, assignment.end_minute
                )
                if not covered:
                    warnings.append(LessonPlacementWarning(
                        LessonPlacementWarningCode.OUTSIDE_STUDENT_AVAILABILITY,
                        "The lesson is outside the student's current Availability Windows.",
                    ))
                elif tentative:
                    warnings.append(LessonPlacementWarning(
                        LessonPlacementWarningCode.TENTATIVE_AVAILABILITY,
                        f"The lesson overlaps {tentative} tentative student availability minutes.",
                    ))
            else:
                tentative = 0
            if instructor is not None:
                covered, instructor_tentative = _window_status(
                    instructor.availability, assignment.day, assignment.start_minute, assignment.end_minute
                )
                if not covered:
                    warnings.append(LessonPlacementWarning(
                        LessonPlacementWarningCode.OUTSIDE_INSTRUCTOR_AVAILABILITY,
                        "The lesson is outside the instructor's current Availability Windows.",
                    ))
                elif instructor_tentative:
                    warnings.append(LessonPlacementWarning(
                        LessonPlacementWarningCode.TENTATIVE_AVAILABILITY,
                        f"The lesson overlaps {instructor_tentative} tentative instructor availability minutes.",
                    ))

        preferred_count = assignment.preferred_tag_match_count
        tentative_minutes = assignment.tentative_minutes
        if student is not None and instructor is not None:
            student_tag_states = dict(student.tag_states)
            required_missing = sorted(
                tags[tag_id]
                for tag_id, state in student_tag_states.items()
                if state is LessonTagState.REQUIRED
                and tag_id in tags
                and tag_id not in instructor.tag_ids
            )
            restricted_present = sorted(
                tags[tag_id]
                for tag_id, state in student_tag_states.items()
                if state is LessonTagState.RESTRICTED
                and tag_id in tags
                and tag_id in instructor.tag_ids
            )
            if required_missing:
                warnings.append(LessonPlacementWarning(
                    LessonPlacementWarningCode.REQUIRED_TAG_MISSING,
                    "Missing Required tags: " + ", ".join(required_missing) + ".",
                ))
            if restricted_present:
                warnings.append(LessonPlacementWarning(
                    LessonPlacementWarningCode.RESTRICTED_TAG_CONFLICT,
                    "Restricted tag conflict: " + ", ".join(restricted_present) + ".",
                ))
            preferred_count = sum(
                state is LessonTagState.PREFERRED and tag_id in instructor.tag_ids
                for tag_id, state in student_tag_states.items()
            )
            if assignment.placement_type == LessonPlacementType.AUTOMATIC.value and (
                student.id, instructor.id, assignment.day, assignment.start_minute
            ) not in feasible_starts:
                warnings.append(LessonPlacementWarning(
                    LessonPlacementWarningCode.NOT_FEASIBLE_CANDIDATE,
                    "The automatic placement is no longer a feasible candidate.",
                ))

        instructor_assignment_count = instructor_counts[assignment.instructor_id]
        if instructor is not None and instructor_assignment_count > _schedulable_lesson_capacity(instructor):
            warnings.append(LessonPlacementWarning(
                LessonPlacementWarningCode.INSTRUCTOR_CAPACITY_EXCEEDED,
                f"{instructor_assignment_count} lessons exceed the instructor's effective capacity "
                f"of {_schedulable_lesson_capacity(instructor)}.",
            ))
        if assignment.student_id in overlap_students:
            warnings.append(LessonPlacementWarning(
                LessonPlacementWarningCode.INSTRUCTOR_OVERLAP,
                "This lesson overlaps another lesson assigned to the instructor.",
            ))
        warning_tuple = tuple(warnings)
        blocking = any(
            item.code is not LessonPlacementWarningCode.TENTATIVE_AVAILABILITY
            for item in warning_tuple
        )
        entries.append(LessonScheduleEntry(
            student_id=student_id,
            student_name=student_name,
            placement_type=assignment.placement_type,
            instructor_id=assignment.instructor_id,
            instructor_name=instructor_name,
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
            verification.append(LessonPlacementVerification(
                student_id=student_id,
                student_name=student_name,
                instructor_id=assignment.instructor_id,
                instructor_name=instructor_name,
                day=assignment.day,
                start_minute=assignment.start_minute,
                warnings=warning_tuple,
            ))

    return tuple(entries), tuple(verification)
