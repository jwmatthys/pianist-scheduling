"""Summary of the persisted Clinical schedule against current source data.

Pure and database-free. Used for display only; it never changes placements.
"""

from __future__ import annotations

from typing import Sequence

from .clinical_candidates import opportunity_window_issue
from .clinical_scheduling_models import (
    CandidateSet,
    PlacementKind,
    PlacementRecord,
    ScheduleSummary,
    SchedulingInput,
    TagPreference,
    UnplacedReason,
    UnplacedStudent,
)


def summarize_schedule(
    scheduling_input: SchedulingInput, candidates: CandidateSet, records: Sequence[PlacementRecord]
) -> ScheduleSummary:
    students = {s.id: s for s in scheduling_input.students}
    opportunities = {o.id: o for o in scheduling_input.opportunities}
    by_record = {item.student_id: item for item in records if item.student_id in students}
    feasible_pairs = {(c.student_id, c.opportunity_id): c for c in candidates.candidates if c.is_feasible}

    preferred = tentative = unverified = used = 0
    placed: set[str] = set()
    for student_id, record in by_record.items():
        if record.kind is PlacementKind.MANUALLY_UNPLACED or record.opportunity_id not in opportunities:
            continue
        placed.add(student_id)
        used += 1
        candidate = feasible_pairs.get((student_id, record.opportunity_id))
        start = None if candidate is None else next(
            (item for item in candidate.feasible_starts if item.start_minute == record.start_minute), None
        )
        if candidate is None or start is None:
            unverified += 1
            continue
        tentative += start.tentative_minutes
        student, opportunity = students[student_id], opportunities[record.opportunity_id]
        preferred += sum(
            1 for tag_id, pref in student.tag_preferences.items()
            if pref is TagPreference.PREFERRED and tag_id in opportunity.tag_ids
        )

    has_feasible = {student_id for student_id, _ in feasible_pairs}
    unplaced = []
    for student in sorted(scheduling_input.students, key=lambda s: (s.label.casefold(), s.id)):
        if student.id in placed:
            continue
        record = by_record.get(student.id)
        if record is not None and record.kind is PlacementKind.MANUALLY_UNPLACED:
            reason = UnplacedReason.MANUALLY_UNPLACED
        elif student.id in has_feasible:
            reason = UnplacedReason.CAPACITY_LIMIT
        else:
            reason = UnplacedReason.NO_FEASIBLE_CANDIDATE
        unplaced.append(UnplacedStudent(student.id, reason))

    total_capacity = sum(o.capacity for o in opportunities.values() if opportunity_window_issue(o) is None)
    return ScheduleSummary(
        student_count=len(students),
        placed_count=len(placed),
        unplaced_count=len(unplaced),
        used_capacity=used,
        total_capacity=total_capacity,
        preferred_tag_matches=preferred,
        tentative_minutes=tentative,
        unverified_placement_count=unverified,
        unplaced=tuple(unplaced),
    )
