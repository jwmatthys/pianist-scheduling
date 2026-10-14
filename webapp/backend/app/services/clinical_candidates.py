"""Clinical candidate generation.

Pure and database-free. Determines which student/opportunity pairs are legal and at
which 15-minute starts. It never assigns a student and never enforces capacity;
selection belongs to the (future) optimizer.
"""

from __future__ import annotations

from .availability import AvailabilityWindow, WEEKDAYS
from .clinical_scheduling_models import (
    START_INCREMENT_MINUTES,
    Candidate,
    CandidateIssue,
    CandidateIssueCode,
    CandidateSet,
    FeasibleStart,
    OpportunityInput,
    SchedulingInput,
    StudentInput,
    TagPreference,
)

MINUTES_PER_DAY = 24 * 60


def opportunity_window_issue(opportunity: OpportunityInput) -> CandidateIssueCode | None:
    """Return the structural problem that stops an opportunity from hosting anyone, if any."""
    window = opportunity.end_minute - opportunity.start_minute
    if (
        opportunity.day not in WEEKDAYS
        or not 0 <= opportunity.start_minute < opportunity.end_minute <= MINUTES_PER_DAY
        or opportunity.session_length_minutes <= 0
        or opportunity.session_length_minutes > window
        or opportunity.travel_minutes < 0
    ):
        return CandidateIssueCode.INVALID_SITE_WINDOW
    if opportunity.capacity <= 0:
        return CandidateIssueCode.ZERO_CAPACITY
    return None


class DayCoverage:
    """Per-minute coverage of one recurring weekday with O(1) interval queries.

    A minute is usable when covered by Available or Tentative time and by no
    Unavailable window. Minutes with no window are not usable: missing is not Available.
    """

    def __init__(self, windows: list[AvailabilityWindow]):
        available = [False] * MINUTES_PER_DAY
        tentative = [False] * MINUTES_PER_DAY
        blocked = [False] * MINUTES_PER_DAY
        for window in windows:
            marks = {"Available": available, "Tentative": tentative, "Unavailable": blocked}.get(window.status)
            if marks is None:
                continue
            for minute in range(max(0, window.start_minute), min(MINUTES_PER_DAY, window.end_minute)):
                marks[minute] = True
        self._usable = [0]
        self._tentative = [0]
        for minute in range(MINUTES_PER_DAY):
            is_usable = (available[minute] or tentative[minute]) and not blocked[minute]
            # Minutes that are both Available and Tentative count as Available.
            is_tentative = is_usable and not available[minute]
            self._usable.append(self._usable[-1] + int(is_usable))
            self._tentative.append(self._tentative[-1] + int(is_tentative))

    def tentative_minutes_if_covered(self, start: int, end: int) -> int | None:
        """Tentative minutes in [start, end) if every minute is usable, otherwise None."""
        if start < 0 or end > MINUTES_PER_DAY or end <= start:
            return None
        if self._usable[end] - self._usable[start] != end - start:
            return None
        return self._tentative[end] - self._tentative[start]


def _coverage_by_day(student: StudentInput) -> dict[str, DayCoverage]:
    by_day: dict[str, list[AvailabilityWindow]] = {}
    for window in student.availability:
        by_day.setdefault(window.day, []).append(window)
    return {day: DayCoverage(windows) for day, windows in by_day.items()}


def feasible_starts(opportunity: OpportunityInput, coverage: DayCoverage | None) -> tuple[FeasibleStart, ...]:
    """Legal session starts on 15-minute clock boundaries.

    The session must lie within the site window; travel before + session + travel
    after must be continuously covered by Available and/or Tentative time.
    """
    if coverage is None:
        return ()
    first = -(-opportunity.start_minute // START_INCREMENT_MINUTES) * START_INCREMENT_MINUTES
    last = opportunity.end_minute - opportunity.session_length_minutes
    starts: list[FeasibleStart] = []
    for start in range(first, last + 1, START_INCREMENT_MINUTES):
        tentative = coverage.tentative_minutes_if_covered(
            start - opportunity.travel_minutes,
            start + opportunity.session_length_minutes + opportunity.travel_minutes,
        )
        if tentative is not None:
            starts.append(FeasibleStart(start_minute=start, tentative_minutes=tentative))
    return tuple(starts)


def _evaluate(
    student: StudentInput, opportunity: OpportunityInput, coverage: dict[str, DayCoverage]
) -> Candidate:
    def result(issues: list[CandidateIssue], starts=(), preferred=0) -> Candidate:
        best = min(starts, key=lambda item: (item.tentative_minutes, item.start_minute)) if starts else None
        return Candidate(
            student_id=student.id,
            opportunity_id=opportunity.id,
            is_feasible=not issues,
            issues=tuple(issues),
            preferred_tag_match_count=preferred if not issues else 0,
            tentative_minutes=best.tentative_minutes if best and not issues else None,
            opportunity_start_minute=best.start_minute if best and not issues else None,
            feasible_starts=tuple(starts) if not issues else (),
            opportunity_capacity=opportunity.capacity,
        )

    structural = opportunity_window_issue(opportunity)
    if structural:
        return result([CandidateIssue(structural)])

    issues: list[CandidateIssue] = []
    missing = sorted(
        tag_id for tag_id, pref in student.tag_preferences.items()
        if pref is TagPreference.REQUIRED and tag_id not in opportunity.tag_ids
    )
    if missing:
        issues.append(CandidateIssue(CandidateIssueCode.REQUIRED_TAG_MISSING, tuple(missing)))
    present = sorted(
        tag_id for tag_id, pref in student.tag_preferences.items()
        if pref is TagPreference.RESTRICTED and tag_id in opportunity.tag_ids
    )
    if present:
        issues.append(CandidateIssue(CandidateIssueCode.RESTRICTED_TAG_PRESENT, tuple(present)))

    starts: tuple[FeasibleStart, ...] = ()
    if not student.availability:
        issues.append(CandidateIssue(CandidateIssueCode.STUDENT_WITHOUT_AVAILABILITY))
    else:
        starts = feasible_starts(opportunity, coverage.get(opportunity.day))
        if not starts:
            issues.append(CandidateIssue(CandidateIssueCode.NO_AVAILABILITY_FIT))

    preferred = sum(
        1 for tag_id, pref in student.tag_preferences.items()
        if pref is TagPreference.PREFERRED and tag_id in opportunity.tag_ids
    )
    return result(issues, starts, preferred)


def generate_candidates(scheduling_input: SchedulingInput) -> CandidateSet:
    """Evaluate every student/opportunity pair in stable (input) order."""
    candidates: list[Candidate] = []
    for student in scheduling_input.students:
        coverage = _coverage_by_day(student)
        for opportunity in scheduling_input.opportunities:
            candidates.append(_evaluate(student, opportunity, coverage))
    return CandidateSet(scheduling_input.source_data_revision, tuple(candidates))
