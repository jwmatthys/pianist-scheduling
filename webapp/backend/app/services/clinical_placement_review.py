"""Current warnings and placement verification for the Clinical schedule."""

from __future__ import annotations

from collections import Counter
from typing import Sequence

from .availability import WEEKDAYS
from .clinical_candidates import DayCoverage, generate_candidates, opportunity_window_issue
from .clinical_scheduling_models import (
    CandidateSet,
    PlacementKind,
    PlacementRecord,
    PlacementVerification,
    PlacementWarning,
    PlacementWarningCode,
    OpportunityInput,
    ScheduleEntry,
    SchedulingInput,
    StudentInput,
    TagPreference,
    UnplacedReason,
)


def _warnings_for(
    student: StudentInput,
    opportunity: OpportunityInput,
    start_minute: int,
    site_assignment_count: int,
    input_data: SchedulingInput,
) -> tuple[PlacementWarning, ...]:
    warnings: list[PlacementWarning] = []
    session_end = start_minute + opportunity.session_length_minutes
    if (
        opportunity_window_issue(opportunity) is not None
        or start_minute < opportunity.start_minute
        or session_end > opportunity.end_minute
    ):
        warnings.append(
            PlacementWarning(
                PlacementWarningCode.OUTSIDE_OPPORTUNITY_WINDOW,
                "Manual session is outside the Placement Opportunity time window.",
            )
        )

    day_is_valid = opportunity.day in WEEKDAYS
    day_windows = [w for w in student.availability if w.day == opportunity.day]
    commitment_start = start_minute - opportunity.travel_minutes
    commitment_end = session_end + opportunity.travel_minutes
    coverage = DayCoverage(day_windows)
    tentative = coverage.tentative_minutes_if_covered(commitment_start, commitment_end)
    if not day_is_valid or tentative is None:
        warnings.append(
            PlacementWarning(
                PlacementWarningCode.OUTSIDE_AVAILABILITY,
                "Placement commitment, including travel before and after, is outside student availability.",
            )
        )

    missing = sorted(
        input_data.tag_names[tag_id]
        for tag_id, state in student.tag_preferences.items()
        if state is TagPreference.REQUIRED and tag_id not in opportunity.tag_ids and tag_id in input_data.tag_names
    )
    if missing:
        warnings.append(
            PlacementWarning(
                PlacementWarningCode.REQUIRED_TAG_MISSING,
                "Missing Required tag" + ("s" if len(missing) != 1 else "") + ": " + ", ".join(missing) + ".",
            )
        )
    restricted = sorted(
        input_data.tag_names[tag_id]
        for tag_id, state in student.tag_preferences.items()
        if state is TagPreference.RESTRICTED and tag_id in opportunity.tag_ids and tag_id in input_data.tag_names
    )
    if restricted:
        warnings.append(
            PlacementWarning(
                PlacementWarningCode.RESTRICTED_TAG_CONFLICT,
                "Restricted tag conflict" + ("s" if len(restricted) != 1 else "") + ": " + ", ".join(restricted) + ".",
            )
        )
    if site_assignment_count > opportunity.capacity:
        warnings.append(
            PlacementWarning(
                PlacementWarningCode.CAPACITY_EXCEEDED,
                f"Capacity exceeded: {site_assignment_count} students assigned; capacity is {opportunity.capacity}.",
            )
        )
    return tuple(warnings)


def review_schedule(
    input_data: SchedulingInput,
    records: Sequence[PlacementRecord],
    candidates: CandidateSet | None = None,
) -> tuple[tuple[ScheduleEntry, ...], tuple[PlacementVerification, ...]]:
    """Build rows and current feasibility diagnostics without changing persistence."""
    candidates = candidates or generate_candidates(input_data)
    students = {s.id: s for s in input_data.students}
    opportunities = {o.id: o for o in input_data.opportunities}
    record_map = {r.student_id: r for r in records if r.student_id in students}
    assigned_counts = Counter(
        r.opportunity_id
        for r in records
        if r.opportunity_id is not None and r.kind is not PlacementKind.MANUALLY_UNPLACED
    )
    feasible = {
        (candidate.student_id, candidate.opportunity_id): candidate
        for candidate in candidates.candidates
        if candidate.is_feasible
    }
    entries: list[ScheduleEntry] = []
    verification: list[PlacementVerification] = []

    has_feasible_candidate = {
        student_id for student_id, _ in feasible
    }
    for student_input in sorted(input_data.students, key=lambda s: (s.label.casefold(), s.id)):
        record = record_map.get(student_input.id)
        kind = record.kind if record else None
        opportunity_id = record.opportunity_id if record else None
        opportunity = opportunities.get(opportunity_id) if opportunity_id is not None else None
        if (
            record is None
            or record.kind is PlacementKind.MANUALLY_UNPLACED
            or opportunity is None
            or record.start_minute is None
        ):
            if record is not None and record.kind is PlacementKind.MANUALLY_UNPLACED:
                reason = UnplacedReason.MANUALLY_UNPLACED
            elif student_input.id in has_feasible_candidate:
                reason = UnplacedReason.CAPACITY_LIMIT
            else:
                reason = UnplacedReason.NO_FEASIBLE_CANDIDATE
            entries.append(
                ScheduleEntry(
                    student_id=student_input.id,
                    student_name=student_input.label,
                    student_number=student_input.student_number,
                    kind=kind,
                    opportunity_id=None,
                    site_name=None,
                    opportunity_label=None,
                    day=None,
                    start_minute=None,
                    end_minute=None,
                    tags=(),
                    preferred_tag_match_count=0,
                    tentative_minutes=None,
                    verified=None,
                    warnings=(),
                    unplaced_reason=reason,
                )
            )
            continue

        warnings = _warnings_for(
            student_input,
            opportunity,
            record.start_minute,
            assigned_counts[opportunity.id],
            input_data,
        )
        tentative_minutes = DayCoverage(
            [w for w in student_input.availability if w.day == opportunity.day]
        ).tentative_minutes_if_covered(
            record.start_minute - opportunity.travel_minutes,
            record.start_minute + opportunity.session_length_minutes + opportunity.travel_minutes,
        )
        candidate = feasible.get((student_input.id, opportunity.id))
        feasible_start = (
            next(
                (item for item in candidate.feasible_starts if item.start_minute == record.start_minute),
                None,
            )
            if candidate is not None
            else None
        )
        if kind is PlacementKind.MANUAL:
            is_verified = not any(
                warning.code is not PlacementWarningCode.CAPACITY_EXCEEDED for warning in warnings
            )
        else:
            is_verified = feasible_start is not None
        tags = tuple(sorted(input_data.tag_names[tag_id] for tag_id in opportunity.tag_ids if tag_id in input_data.tag_names))
        preferred_count = sum(
            1
            for tag_id, preference in student_input.tag_preferences.items()
            if preference is TagPreference.PREFERRED and tag_id in opportunity.tag_ids
        )
        entry_warnings = warnings
        if tentative_minutes:
            # Informational only: tentative time is allowed, so it never affects verification.
            entry_warnings = (
                *warnings,
                PlacementWarning(
                    PlacementWarningCode.TENTATIVE_AVAILABILITY,
                    f"Placement falls during tentative availability ({tentative_minutes} minutes, including travel).",
                ),
            )
        entry = ScheduleEntry(
            student_id=student_input.id,
            student_name=student_input.label,
            student_number=student_input.student_number,
            kind=kind,
            opportunity_id=opportunity.id,
            site_name=opportunity.site_name or opportunity.label,
            opportunity_label=opportunity.label,
            day=opportunity.day,
            start_minute=record.start_minute,
            end_minute=record.start_minute + opportunity.session_length_minutes,
            tags=tags,
            preferred_tag_match_count=preferred_count,
            tentative_minutes=tentative_minutes,
            verified=is_verified,
            warnings=entry_warnings,
            unplaced_reason=None,
        )
        entries.append(entry)
        if not is_verified:
            verification.append(
                PlacementVerification(
                    student_id=student_input.id,
                    student_name=student_input.label,
                    opportunity_id=opportunity.id,
                    site_name=entry.site_name,
                    day=opportunity.day,
                    start_minute=record.start_minute,
                    warnings=warnings,
                )
            )

    return tuple(entries), tuple(verification)
