"""Exact, database-free optimizer for Lesson Scheduling."""

from __future__ import annotations

from collections import defaultdict
from typing import Sequence

from ..lesson_models import LessonUnplacedReason
from .availability import AvailabilityWindow, WEEKDAYS, validate_window
from .lesson_scheduling_models import (
    MINUTES_PER_DAY,
    LessonCandidate,
    LessonCandidateSet,
    LessonInstructorInput,
    LessonOptimizationResult,
    LessonScheduleStatistics,
    LessonScheduledPlacement,
    LessonSchedulingInput,
    LessonUnplacedStudentResult,
)


def _usable_coverage(
    windows: tuple[AvailabilityWindow, ...],
) -> dict[str, list[bool]]:
    available = {day: [False] * MINUTES_PER_DAY for day in WEEKDAYS}
    tentative = {day: [False] * MINUTES_PER_DAY for day in WEEKDAYS}
    blocked = {day: [False] * MINUTES_PER_DAY for day in WEEKDAYS}
    for window in windows:
        if validate_window(window) is not None:
            continue
        coverage = {
            "Available": available,
            "Tentative": tentative,
            "Unavailable": blocked,
        }[window.status]
        for minute in range(window.start_minute, window.end_minute):
            coverage[window.day][minute] = True
    return {
        day: [
            (available[day][minute] or tentative[day][minute]) and not blocked[day][minute]
            for minute in range(MINUTES_PER_DAY)
        ]
        for day in WEEKDAYS
    }


def _coverage_segments(coverage: list[bool]) -> tuple[tuple[int, int], ...]:
    segments: list[tuple[int, int]] = []
    start: int | None = None
    for minute, covered in enumerate(coverage + [False]):
        if covered and start is None:
            start = minute
        elif not covered and start is not None:
            segments.append((start, minute))
            start = None
    return tuple(segments)


def _schedulable_lesson_capacity(instructor: LessonInstructorInput) -> int:
    if instructor.capacity <= 0:
        return 0
    length = instructor.lesson_length_minutes
    if not 0 < length <= MINUTES_PER_DAY:
        return 0
    coverage = _usable_coverage(instructor.availability)
    possible = 0
    for day in WEEKDAYS:
        segments = _coverage_segments(coverage[day])
        for segment_start, segment_end in segments:
            start = ((segment_start + 14) // 15) * 15
            while start + length <= segment_end:
                possible += 1
                start += ((length + 14) // 15) * 15
    return min(instructor.capacity, possible)


def _gap_cost(
    placements: list[LessonCandidate],
    instructor_coverage: dict[str, dict[str, list[bool]]],
) -> int:
    cost = 0
    by_instructor_day: dict[tuple[str, str], list[LessonCandidate]] = defaultdict(list)
    for item in placements:
        by_instructor_day[(item.instructor_id, item.day)].append(item)
    for (instructor_id, day), lessons in by_instructor_day.items():
        segments = _coverage_segments(instructor_coverage[instructor_id][day])
        ordered = sorted(lessons, key=lambda item: (item.lesson_start_minute, item.lesson_end_minute))
        for segment_start, segment_end in segments:
            inside = [
                item for item in ordered
                if item.lesson_start_minute >= segment_start and item.lesson_end_minute <= segment_end
            ]
            cost += sum(
                max(0, next_item.lesson_start_minute - current.lesson_end_minute)
                for current, next_item in zip(inside, inside[1:])
            )
    return cost


def optimize_lesson_schedule(
    scheduling_input: LessonSchedulingInput,
    candidates: LessonCandidateSet,
    fixed_placements: Sequence[LessonScheduledPlacement] = (),
    excluded_student_ids: frozenset[str] = frozenset(),
) -> LessonOptimizationResult:
    """Choose the exact lexicographic optimum over feasible candidates.

    Tie-breaking uses student and instructor names/IDs, weekday, then start time;
    it is independent of input and import order.
    """
    if candidates.source_data_revision != scheduling_input.source_data_revision:
        raise ValueError("The candidate set was generated from a different source-data revision.")

    students = sorted(scheduling_input.students, key=lambda item: (item.name.casefold(), item.id))
    instructors = sorted(scheduling_input.instructors, key=lambda item: (item.name.casefold(), item.id))
    student_by_id = {item.id: item for item in students}
    instructor_by_id = {item.id: item for item in instructors}
    weekday_order = {day: index for index, day in enumerate(WEEKDAYS)}
    effective_capacity = {
        instructor.id: _schedulable_lesson_capacity(instructor) for instructor in instructors
    }
    instructor_coverage = {
        instructor.id: _usable_coverage(instructor.availability) for instructor in instructors
    }
    student_coverage = {
        student.id: _usable_coverage(student.availability) for student in students
    }
    student_tags = {
        student.id: dict(student.tag_states) for student in students
    }

    fixed_student_ids = {item.student_id for item in fixed_placements}
    excluded_student_ids = excluded_student_ids | fixed_student_ids
    options_by_student: dict[str, list[LessonCandidate]] = {
        item.id: [] for item in students if item.id not in excluded_student_ids
    }
    seen: set[tuple[str, str, str, int, int]] = set()
    for candidate in candidates.candidates:
        if not candidate.is_feasible:
            continue
        student = student_by_id.get(candidate.student_id)
        instructor = instructor_by_id.get(candidate.instructor_id)
        if student is None or instructor is None:
            continue
        if candidate.day not in weekday_order:
            continue
        if (
            candidate.lesson_end_minute - candidate.lesson_start_minute
            != instructor.lesson_length_minutes
            or not 0 <= candidate.lesson_start_minute < candidate.lesson_end_minute <= MINUTES_PER_DAY
            or candidate.lesson_start_minute % 15 != 0
            or candidate.issues
        ):
            continue
        if not all(
            student_coverage[student.id][candidate.day][minute]
            and instructor_coverage[instructor.id][candidate.day][minute]
            for minute in range(candidate.lesson_start_minute, candidate.lesson_end_minute)
        ):
            continue
        instructor_tags = instructor.tag_ids
        required_tags = {
            tag_id for tag_id, state in student_tags[student.id].items()
            if state.value == "required"
        }
        restricted_tags = {
            tag_id for tag_id, state in student_tags[student.id].items()
            if state.value == "restricted"
        }
        if not required_tags.issubset(instructor_tags) or restricted_tags.intersection(instructor_tags):
            continue
        key = (
            candidate.student_id,
            candidate.instructor_id,
            candidate.day,
            candidate.lesson_start_minute,
            candidate.lesson_end_minute,
        )
        if key in seen:
            continue
        seen.add(key)
        if student.id in options_by_student:
            options_by_student[student.id].append(candidate)

    def option_key(candidate: LessonCandidate) -> tuple:
        instructor = instructor_by_id[candidate.instructor_id]
        return (
            -candidate.preferred_tag_match_count,
            candidate.tentative_minutes,
            instructor.name.casefold(),
            instructor.id,
            weekday_order[candidate.day],
            candidate.lesson_start_minute,
        )

    for options in options_by_student.values():
        options.sort(key=option_key)
    ordered_students = sorted(
        (item for item in students if item.id in options_by_student),
        key=lambda item: (
            len(options_by_student[item.id]),
            item.name.casefold(),
            item.id,
        ),
    )
    max_preferred = {
        student.id: max(
            (candidate.preferred_tag_match_count for candidate in options_by_student[student.id]),
            default=0,
        )
        for student in ordered_students
    }
    min_tentative = {
        student.id: min(
            (candidate.tentative_minutes for candidate in options_by_student[student.id]),
            default=0,
        )
        for student in ordered_students
    }

    assigned: list[LessonCandidate] = []
    instructor_assignments: dict[str, list[LessonCandidate]] = defaultdict(list)
    instructor_counts: dict[str, int] = defaultdict(int)
    for placement in fixed_placements:
        instructor = instructor_by_id.get(placement.instructor_id)
        if instructor is None:
            continue
        fixed_candidate = LessonCandidate(
            student_id=placement.student_id,
            instructor_id=placement.instructor_id,
            day=placement.day,
            lesson_start_minute=placement.start_minute,
            lesson_end_minute=placement.end_minute,
            preferred_tag_match_count=placement.preferred_tag_match_count,
            tentative_minutes=placement.tentative_minutes,
            issues=(),
            is_feasible=True,
            instructor_capacity=instructor.capacity,
        )
        assigned.append(fixed_candidate)
        instructor_assignments[instructor.id].append(fixed_candidate)
        instructor_counts[instructor.id] += 1
    # Students in canonical order; options within a student in canonical order.
    # The canonical tuple compares students first, so a depth-first search in this
    # order that tries "skip student" last meets equal-objective solutions in
    # canonical order and the first one found is the canonical minimum.
    canonical_students = sorted(ordered_students, key=lambda item: (item.name.casefold(), item.id))
    canonical_options = {
        student.id: sorted(
            options_by_student[student.id],
            key=lambda c: (
                instructor_by_id[c.instructor_id].name.casefold(),
                c.instructor_id,
                weekday_order[c.day],
                c.lesson_start_minute,
            ),
        )
        for student in canonical_students
    }
    options_by_student_instructor: dict[str, dict[str, list[LessonCandidate]]] = {}
    for student_id, student_options in options_by_student.items():
        grouped: dict[str, list[LessonCandidate]] = defaultdict(list)
        for candidate in student_options:
            grouped[candidate.instructor_id].append(candidate)
        options_by_student_instructor[student_id] = dict(grouped)
    segments_cache: dict[tuple[str, str], tuple[tuple[int, int], ...]] = {}

    def segments_for(instructor_id: str, day: str) -> tuple[tuple[int, int], ...]:
        key = (instructor_id, day)
        if key not in segments_cache:
            segments_cache[key] = _coverage_segments(instructor_coverage[instructor_id][day])
        return segments_cache[key]

    def gap_lower_bound() -> int:
        """Gap that cannot be removed by inserting more same-length lessons."""
        total = 0
        for instructor_id, lessons in instructor_assignments.items():
            length = instructor_by_id[instructor_id].lesson_length_minutes
            if length <= 0:
                continue
            by_day: dict[str, list[LessonCandidate]] = defaultdict(list)
            for lesson in lessons:
                by_day[lesson.day].append(lesson)
            for day, day_lessons in by_day.items():
                ordered = sorted(day_lessons, key=lambda item: item.lesson_start_minute)
                for segment_start, segment_end in segments_for(instructor_id, day):
                    inside = [
                        item for item in ordered
                        if item.lesson_start_minute >= segment_start and item.lesson_end_minute <= segment_end
                    ]
                    for current, following in zip(inside, inside[1:]):
                        total += max(0, following.lesson_start_minute - current.lesson_end_minute) % length
        return total

    def can_place(candidate: LessonCandidate) -> bool:
        instructor = instructor_by_id[candidate.instructor_id]
        if instructor_counts[instructor.id] >= effective_capacity[instructor.id]:
            return False
        return not any(
            existing.day == candidate.day
            and candidate.lesson_start_minute < existing.lesson_end_minute
            and candidate.lesson_end_minute > existing.lesson_start_minute
            for existing in instructor_assignments[instructor.id]
        )

    def node_bound(index: int, preferred_total: int, tentative_total: int, order: list) -> tuple:
        possible = [
            student for student in order[index:]
            if any(
                instructor_counts[instructor_id] < effective_capacity[instructor_id]
                and any(can_place(candidate) for candidate in instructor_options)
                for instructor_id, instructor_options in options_by_student_instructor[student.id].items()
            )
        ]
        return (
            -(len(assigned) + len(possible)),
            -(preferred_total + sum(max_preferred[item.id] for item in possible)),
            tentative_total + sum(min_tentative[item.id] for item in possible),
            gap_lower_bound(),
        )

    def leaf_value() -> tuple:
        return (
            -len(assigned),
            -sum(item.preferred_tag_match_count for item in assigned),
            sum(item.tentative_minutes for item in assigned),
            _gap_cost(assigned, instructor_coverage),
        )

    def run_search(order: list, option_lists, prune, on_leaf, option_sort=None) -> None:
        def push(candidate: LessonCandidate) -> None:
            assigned.append(candidate)
            instructor_assignments[candidate.instructor_id].append(candidate)
            instructor_counts[candidate.instructor_id] += 1

        def pop(candidate: LessonCandidate) -> None:
            instructor_counts[candidate.instructor_id] -= 1
            instructor_assignments[candidate.instructor_id].pop()
            assigned.pop()

        def search(index: int, preferred_total: int, tentative_total: int) -> bool:
            if prune(node_bound(index, preferred_total, tentative_total, order)):
                return False
            if index == len(order):
                return on_leaf()
            student = order[index]
            for candidate in option_lists[student.id]:
                if not can_place(candidate):
                    continue
                push(candidate)
                done = search(
                    index + 1,
                    preferred_total + candidate.preferred_tag_match_count,
                    tentative_total + candidate.tentative_minutes,
                )
                pop(candidate)
                if done:
                    return True
            return search(index + 1, preferred_total, tentative_total)

        search(0, 0, 0)

    # Phase A: find the optimal (count, preferred, tentative, gap) value. Only
    # strictly better solutions matter, so ties are pruned.
    best_value: list[tuple | None] = [None]

    def prune_a(bound: tuple) -> bool:
        return best_value[0] is not None and bound >= best_value[0]

    def leaf_a() -> bool:
        value = leaf_value()
        if best_value[0] is None or value < best_value[0]:
            best_value[0] = value
        return False

    run_search(ordered_students, options_by_student, prune_a, leaf_a)

    # Phase B: the first solution reaching that value, in canonical order, is the
    # deterministic tie-break winner.
    best_placements: tuple[LessonCandidate, ...] = ()
    target = best_value[0]

    def prune_b(bound: tuple) -> bool:
        return target is not None and bound > target

    def leaf_b() -> bool:
        nonlocal best_placements
        if leaf_value() == target:
            best_placements = tuple(assigned)
            return True
        return False

    if target is not None:
        run_search(canonical_students, canonical_options, prune_b, leaf_b)

    best_student_ids = {item.student_id for item in best_placements}
    feasible_by_student = {student.id: options_by_student[student.id] for student in ordered_students}
    unplaced = []
    for student in ordered_students:
        if student.id in best_student_ids:
            continue
        remaining_capacity_options = [
            candidate
            for candidate in feasible_by_student[student.id]
            if sum(1 for item in best_placements if item.instructor_id == candidate.instructor_id)
            < effective_capacity[candidate.instructor_id]
        ]
        has_non_overlapping_option = any(
            not any(
                existing.instructor_id == candidate.instructor_id
                and existing.day == candidate.day
                and candidate.lesson_start_minute < existing.lesson_end_minute
                and candidate.lesson_end_minute > existing.lesson_start_minute
                for existing in best_placements
            )
            for candidate in remaining_capacity_options
        )
        unplaced.append(LessonUnplacedStudentResult(
            student_id=student.id,
            reason=(
                LessonUnplacedReason.NO_FEASIBLE_CANDIDATE
                if not feasible_by_student[student.id]
                else LessonUnplacedReason.INSTRUCTOR_OVERLAP
                if remaining_capacity_options and not has_non_overlapping_option
                else LessonUnplacedReason.CAPACITY_LIMIT
            ),
        ))

    ordered_placements = tuple(
        LessonScheduledPlacement(
            student_id=item.student_id,
            instructor_id=item.instructor_id,
            day=item.day,
            start_minute=item.lesson_start_minute,
            end_minute=item.lesson_end_minute,
            preferred_tag_match_count=item.preferred_tag_match_count,
            tentative_minutes=item.tentative_minutes,
        )
        for item in sorted(best_placements, key=lambda item: (
            student_by_id[item.student_id].name.casefold()
            if item.student_id in student_by_id else item.student_id,
            item.student_id,
        ))
    )
    denominator = sum(effective_capacity.values())
    scheduled = len(ordered_placements)
    return LessonOptimizationResult(
        source_data_revision=scheduling_input.source_data_revision,
        placements=ordered_placements,
        unplaced_students=tuple(unplaced),
        statistics=LessonScheduleStatistics(
            students_scheduled=scheduled,
            students_unscheduled=len(unplaced),
            instructor_utilization=scheduled / denominator if denominator else 0.0,
            preferred_tag_matches=sum(item.preferred_tag_match_count for item in ordered_placements),
            tentative_minutes=sum(item.tentative_minutes for item in ordered_placements),
            gap_cost_minutes=_gap_cost(list(best_placements), instructor_coverage),
        ),
    )
