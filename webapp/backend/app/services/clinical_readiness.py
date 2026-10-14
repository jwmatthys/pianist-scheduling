"""Clinical readiness analysis.

Pure and database-free. Consumes a scheduling input snapshot and its candidate set
and returns typed, machine-readable warnings. It does not generate candidates or
schedules.
"""

from __future__ import annotations

from .clinical_candidates import opportunity_window_issue
from .clinical_scheduling_models import (
    CandidateIssueCode,
    CandidateSet,
    ReadinessCategory,
    ReadinessCode,
    ReadinessReport,
    ReadinessSeverity,
    ReadinessSubject,
    ReadinessWarning,
    SchedulingInput,
    SubjectKind,
    TagPreference,
)

_ERROR = ReadinessSeverity.ERROR
_WARNING = ReadinessSeverity.WARNING
_INFO = ReadinessSeverity.INFO


def analyze_readiness(scheduling_input: SchedulingInput, candidates: CandidateSet) -> ReadinessReport:
    students = scheduling_input.students
    opportunities = scheduling_input.opportunities
    warnings: list[ReadinessWarning] = []

    def add(code, severity, category, subject=None):
        warnings.append(ReadinessWarning(ReadinessCode(code), severity, category, subject))

    def student_subject(student):
        return ReadinessSubject(SubjectKind.STUDENT, student.id, student.label)

    def opportunity_subject(opportunity):
        return ReadinessSubject(SubjectKind.OPPORTUNITY, opportunity.id, opportunity.label)

    if not students:
        add(ReadinessCode.NO_STUDENTS, _ERROR, ReadinessCategory.DATA)
    if not opportunities:
        add(ReadinessCode.NO_PLACEMENT_OPPORTUNITIES, _ERROR, ReadinessCategory.DATA)

    by_student: dict[str, list] = {}
    by_opportunity: dict[str, list] = {}
    for item in candidates.candidates:
        by_student.setdefault(item.student_id, []).append(item)
        by_opportunity.setdefault(item.opportunity_id, []).append(item)

    # Opportunities: structural problems first; they cannot host anyone.
    usable_opportunities = []
    for opportunity in opportunities:
        problem = opportunity_window_issue(opportunity)
        if problem is CandidateIssueCode.INVALID_SITE_WINDOW:
            add(ReadinessCode.INVALID_SITE_WINDOW, _ERROR, ReadinessCategory.OPPORTUNITIES, opportunity_subject(opportunity))
        elif problem is CandidateIssueCode.ZERO_CAPACITY:
            add(ReadinessCode.ZERO_CAPACITY, _ERROR, ReadinessCategory.OPPORTUNITIES, opportunity_subject(opportunity))
        else:
            usable_opportunities.append(opportunity)

    # Students without any availability are reported separately from students whose
    # availability simply does not fit any opportunity.
    no_availability = {student.id for student in students if not student.availability}
    for student in students:
        if student.id in no_availability:
            add(ReadinessCode.STUDENT_WITHOUT_AVAILABILITY, _WARNING, ReadinessCategory.STUDENTS, student_subject(student))
    if usable_opportunities:
        for student in students:
            if student.id in no_availability:
                continue
            if not any(item.is_feasible for item in by_student.get(student.id, [])):
                add(ReadinessCode.STUDENT_WITHOUT_CANDIDATES, _WARNING, ReadinessCategory.STUDENTS, student_subject(student))

    # An opportunity has an availability match when some student's availability can
    # hold its full commitment, regardless of tags.
    if students:
        for opportunity in usable_opportunities:
            fits = any(
                item.opportunity_id == opportunity.id
                and CandidateIssueCode.NO_AVAILABILITY_FIT not in item.issue_codes
                and CandidateIssueCode.STUDENT_WITHOUT_AVAILABILITY not in item.issue_codes
                for item in by_opportunity.get(opportunity.id, [])
            )
            if not fits:
                add(ReadinessCode.OPPORTUNITY_WITHOUT_AVAILABILITY_MATCH, _WARNING,
                    ReadinessCategory.OPPORTUNITIES, opportunity_subject(opportunity))

    # A Required tag that no opportunity carries cannot be satisfied for anyone requiring it.
    required_tags = sorted({
        tag_id for student in students for tag_id, pref in student.tag_preferences.items()
        if pref is TagPreference.REQUIRED
    }, key=lambda tag_id: scheduling_input.tag_names.get(tag_id, "").casefold())
    carried = {tag_id for opportunity in opportunities for tag_id in opportunity.tag_ids}
    for tag_id in required_tags:
        if tag_id not in carried:
            add(ReadinessCode.REQUIRED_TAG_WITHOUT_MATCH, _WARNING, ReadinessCategory.TAGS,
                ReadinessSubject(SubjectKind.TAG, tag_id, scheduling_input.tag_names.get(tag_id, "")))

    feasible_count = len(candidates.feasible)
    if students and usable_opportunities and feasible_count == 0:
        add(ReadinessCode.NO_FEASIBLE_CANDIDATES, _ERROR, ReadinessCategory.DATA)

    # Informational: more students than total automatic capacity guarantees unplaced students.
    if students and usable_opportunities:
        total_capacity = sum(opportunity.capacity for opportunity in usable_opportunities)
        if total_capacity < len(students):
            add(ReadinessCode.CAPACITY_BELOW_STUDENT_COUNT, _INFO, ReadinessCategory.DATA)

    return ReadinessReport(
        source_data_revision=scheduling_input.source_data_revision,
        student_count=len(students),
        opportunity_count=len(opportunities),
        feasible_candidate_count=feasible_count,
        total_candidate_count=len(candidates.candidates),
        warnings=tuple(warnings),
    )
