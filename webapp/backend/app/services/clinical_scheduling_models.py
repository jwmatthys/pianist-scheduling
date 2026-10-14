"""Typed, database-free structures for Clinical validation and candidate generation.

These are Clinical-owned and deliberately independent of persistence, HTTP, and the
Accompanist/Jury modules. Nothing here assigns a student to an opportunity.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from .availability import AvailabilityWindow

# Automatic start times fall on 15-minute clock boundaries (spec 8.3).
START_INCREMENT_MINUTES = 15


class TagPreference(str, Enum):
    REQUIRED = "required"
    PREFERRED = "preferred"
    RESTRICTED = "restricted"


@dataclass(frozen=True)
class OpportunityInput:
    id: str
    label: str
    day: str
    start_minute: int
    end_minute: int
    session_length_minutes: int
    travel_minutes: int
    capacity: int
    tag_ids: frozenset[str] = frozenset()


@dataclass(frozen=True)
class StudentInput:
    id: str
    label: str
    availability: tuple[AvailabilityWindow, ...] = ()
    # Required/Preferred/Restricted only; Neutral is the absence of an entry.
    tag_preferences: dict[str, TagPreference] = field(default_factory=dict)


@dataclass(frozen=True)
class SchedulingInput:
    """Immutable snapshot of Clinical scheduling inputs at one source-data revision."""

    source_data_revision: int
    students: tuple[StudentInput, ...]
    opportunities: tuple[OpportunityInput, ...]
    tag_names: dict[str, str]


class CandidateIssueCode(str, Enum):
    INVALID_SITE_WINDOW = "INVALID_SITE_WINDOW"
    ZERO_CAPACITY = "ZERO_CAPACITY"
    STUDENT_WITHOUT_AVAILABILITY = "STUDENT_WITHOUT_AVAILABILITY"
    NO_AVAILABILITY_FIT = "NO_AVAILABILITY_FIT"
    REQUIRED_TAG_MISSING = "REQUIRED_TAG_MISSING"
    RESTRICTED_TAG_PRESENT = "RESTRICTED_TAG_PRESENT"


@dataclass(frozen=True)
class CandidateIssue:
    """A machine-readable reason a student/opportunity pair is not a legal candidate."""

    code: CandidateIssueCode
    tag_ids: tuple[str, ...] = ()


@dataclass(frozen=True)
class FeasibleStart:
    """One legal session start, with Tentative minutes across travel + session + travel."""

    start_minute: int
    tentative_minutes: int


@dataclass(frozen=True)
class Candidate:
    student_id: str
    opportunity_id: str
    is_feasible: bool
    issues: tuple[CandidateIssue, ...]
    # Preference metrics. For an infeasible candidate the metrics are 0 / None.
    preferred_tag_match_count: int
    tentative_minutes: int | None
    opportunity_start_minute: int | None
    feasible_starts: tuple[FeasibleStart, ...]
    # Exposed for the optimizer; never enforced during candidate generation.
    opportunity_capacity: int

    @property
    def issue_codes(self) -> frozenset[CandidateIssueCode]:
        return frozenset(issue.code for issue in self.issues)


@dataclass(frozen=True)
class CandidateSet:
    source_data_revision: int
    candidates: tuple[Candidate, ...]

    @property
    def feasible(self) -> tuple[Candidate, ...]:
        return tuple(item for item in self.candidates if item.is_feasible)

    def for_student(self, student_id: str) -> tuple[Candidate, ...]:
        return tuple(item for item in self.candidates if item.student_id == student_id)

    def feasible_for_student(self, student_id: str) -> tuple[Candidate, ...]:
        return tuple(item for item in self.for_student(student_id) if item.is_feasible)


class ReadinessSeverity(str, Enum):
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


class ReadinessCode(str, Enum):
    NO_STUDENTS = "NO_STUDENTS"
    NO_PLACEMENT_OPPORTUNITIES = "NO_PLACEMENT_OPPORTUNITIES"
    NO_FEASIBLE_CANDIDATES = "NO_FEASIBLE_CANDIDATES"
    STUDENT_WITHOUT_AVAILABILITY = "STUDENT_WITHOUT_AVAILABILITY"
    STUDENT_WITHOUT_CANDIDATES = "STUDENT_WITHOUT_CANDIDATES"
    OPPORTUNITY_WITHOUT_AVAILABILITY_MATCH = "OPPORTUNITY_WITHOUT_AVAILABILITY_MATCH"
    REQUIRED_TAG_WITHOUT_MATCH = "REQUIRED_TAG_WITHOUT_MATCH"
    ZERO_CAPACITY = "ZERO_CAPACITY"
    INVALID_SITE_WINDOW = "INVALID_SITE_WINDOW"
    CAPACITY_BELOW_STUDENT_COUNT = "CAPACITY_BELOW_STUDENT_COUNT"


class ReadinessCategory(str, Enum):
    DATA = "data"
    STUDENTS = "students"
    OPPORTUNITIES = "opportunities"
    TAGS = "tags"


class SubjectKind(str, Enum):
    STUDENT = "student"
    OPPORTUNITY = "opportunity"
    TAG = "tag"


@dataclass(frozen=True)
class ReadinessSubject:
    kind: SubjectKind
    id: str
    label: str


@dataclass(frozen=True)
class ReadinessWarning:
    code: ReadinessCode
    severity: ReadinessSeverity
    category: ReadinessCategory
    subject: ReadinessSubject | None = None


@dataclass(frozen=True)
class ReadinessReport:
    source_data_revision: int
    student_count: int
    opportunity_count: int
    feasible_candidate_count: int
    total_candidate_count: int
    warnings: tuple[ReadinessWarning, ...]

    @property
    def is_ready(self) -> bool:
        """Ready means at least one legal candidate exists and nothing blocks generation."""
        return not any(item.severity is ReadinessSeverity.ERROR for item in self.warnings)

    def count(self, code: ReadinessCode) -> int:
        return sum(1 for item in self.warnings if item.code is code)
