"""Summary of the persisted Clinical schedule against current source data.

Pure and database-free. Used for display only; it never changes placements.
"""

from __future__ import annotations

from typing import Sequence

from .clinical_candidates import opportunity_window_issue
from .clinical_placement_review import review_schedule
from .clinical_scheduling_models import (
    CandidateSet,
    PlacementKind,
    PlacementRecord,
    ScheduleSummary,
    SchedulingInput,
    UnplacedReason,
    UnplacedStudent,
)


def summarize_schedule(
    scheduling_input: SchedulingInput, candidates: CandidateSet, records: Sequence[PlacementRecord]
) -> ScheduleSummary:
    students = {s.id: s for s in scheduling_input.students}
    opportunities = {o.id: o for o in scheduling_input.opportunities}
    by_record = {item.student_id: item for item in records if item.student_id in students}
    reviewed_entries, verification = review_schedule(scheduling_input, records, candidates)
    reviewed_by_student = {item.student_id: item for item in reviewed_entries}

    preferred = tentative = used = 0
    placed: set[str] = set()
    for student_id, record in by_record.items():
        if record.kind is PlacementKind.MANUALLY_UNPLACED or record.opportunity_id not in opportunities:
            continue
        placed.add(student_id)
        used += 1
        entry = reviewed_by_student[student_id]
        preferred += entry.preferred_tag_match_count
        if entry.tentative_minutes is not None:
            tentative += entry.tentative_minutes

    has_feasible = {
        candidate.student_id for candidate in candidates.candidates if candidate.is_feasible
    }
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
        unverified_placement_count=len(verification),
        unplaced=tuple(unplaced),
    )
