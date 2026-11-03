"""Typed Tutor Scheduling inputs, candidates, readiness, and schedule results."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from ..tutor_domain import TutorTagState, TutorUnplacedReason
from .availability import AvailabilityWindow, WEEKDAYS

TUTOR_START_INCREMENT_MINUTES = 15
MINUTES_PER_DAY = 24 * 60


@dataclass(frozen=True)
class TutorTutorInput:
    id: str
    name: str
    lesson_length_minutes: int
    capacity: int
    tag_ids: frozenset[str]
    availability: tuple[AvailabilityWindow, ...]


@dataclass(frozen=True)
class TutorStudentInput:
    id: str
    name: str
    availability: tuple[AvailabilityWindow, ...]
    # Neutral is represented by an absent entry.
    tag_states: tuple[tuple[str, TutorTagState], ...]
    # Records that share a person key are one person and are never scheduled at overlapping times.
    person_key: str = ""


@dataclass(frozen=True)
class TutorTagInput:
    id: str
    name: str


@dataclass(frozen=True)
class TutorSchedulingInput:
    """Tutor-owned scheduling inputs captured at one source-data revision."""

    source_data_revision: int
    students: tuple[TutorStudentInput, ...]
    tutors: tuple[TutorTutorInput, ...]
    tags: tuple[TutorTagInput, ...] = ()


class TutorCandidateIssueCode(str, Enum):
    REQUIRED_TAG_MISSING = "REQUIRED_TAG_MISSING"
    RESTRICTED_TAG_PRESENT = "RESTRICTED_TAG_PRESENT"


@dataclass(frozen=True)
class TutorCandidateIssue:
    code: TutorCandidateIssueCode
    tag_ids: tuple[str, ...] = ()


@dataclass(frozen=True)
class TutorCandidate:
    student_id: str
    tutor_id: str
    day: str
    tutor_start_minute: int
    tutor_end_minute: int
    preferred_tag_match_count: int
    tentative_minutes: int
    issues: tuple[TutorCandidateIssue, ...]
    is_feasible: bool
    # Exposed to future scheduling layers; candidate generation never enforces it.
    tutor_capacity: int


@dataclass(frozen=True)
class TutorCandidateSet:
    source_data_revision: int
    candidates: tuple[TutorCandidate, ...]

    @property
    def feasible(self) -> tuple[TutorCandidate, ...]:
        return tuple(candidate for candidate in self.candidates if candidate.is_feasible)

    def for_student(self, student_id: str) -> tuple[TutorCandidate, ...]:
        return tuple(candidate for candidate in self.candidates if candidate.student_id == student_id)

    def feasible_for_student(self, student_id: str) -> tuple[TutorCandidate, ...]:
        return tuple(candidate for candidate in self.for_student(student_id) if candidate.is_feasible)


class TutorReadinessSeverity(str, Enum):
    ERROR = "error"
    WARNING = "warning"


class TutorReadinessCategory(str, Enum):
    DATA = "data"
    STUDENTS = "students"
    TUTORS = "tutors"
    TAGS = "tags"


class TutorReadinessSubjectKind(str, Enum):
    STUDENT = "student"
    TUTOR = "tutor"
    TAG = "tag"


class TutorReadinessCode(str, Enum):
    NO_STUDENTS = "NO_STUDENTS"
    NO_TUTORS = "NO_TUTORS"
    NO_FEASIBLE_CANDIDATES = "NO_FEASIBLE_CANDIDATES"
    STUDENT_WITHOUT_AVAILABILITY = "STUDENT_WITHOUT_AVAILABILITY"
    TUTOR_WITHOUT_AVAILABILITY = "TUTOR_WITHOUT_AVAILABILITY"
    STUDENT_WITHOUT_CANDIDATES = "STUDENT_WITHOUT_CANDIDATES"
    REQUIRED_TAG_WITHOUT_MATCH = "REQUIRED_TAG_WITHOUT_MATCH"
    ZERO_CAPACITY = "ZERO_CAPACITY"
    INVALID_TUTOR_WINDOW = "INVALID_TUTOR_WINDOW"
    TUTOR_WITHOUT_SCHEDULABLE_TUTORS = "TUTOR_WITHOUT_SCHEDULABLE_TUTORS"


@dataclass(frozen=True)
class TutorReadinessSubject:
    kind: TutorReadinessSubjectKind
    id: str
    label: str


@dataclass(frozen=True)
class TutorReadinessWarning:
    code: TutorReadinessCode
    severity: TutorReadinessSeverity
    category: TutorReadinessCategory
    subjects: tuple[TutorReadinessSubject, ...] = ()


@dataclass(frozen=True)
class TutorReadinessReport:
    source_data_revision: int
    is_ready: bool
    student_count: int
    tutor_count: int
    feasible_candidate_count: int
    total_candidate_count: int
    students_with_candidates: int
    tutors_with_schedulable_tutors: int
    warnings: tuple[TutorReadinessWarning, ...]


@dataclass(frozen=True)
class TutorScheduledPlacement:
    student_id: str
    tutor_id: str
    day: str
    start_minute: int
    end_minute: int
    preferred_tag_match_count: int
    tentative_minutes: int


@dataclass(frozen=True)
class TutorUnplacedStudentResult:
    student_id: str
    reason: TutorUnplacedReason


class TutorPlacementWarningCode(str, Enum):
    STUDENT_REMOVED = "STUDENT_REMOVED"
    TUTOR_REMOVED = "TUTOR_REMOVED"
    INVALID_TIME = "INVALID_TIME"
    LESSON_LENGTH_CHANGED = "LESSON_LENGTH_CHANGED"
    OUTSIDE_STUDENT_AVAILABILITY = "OUTSIDE_STUDENT_AVAILABILITY"
    OUTSIDE_TUTOR_AVAILABILITY = "OUTSIDE_TUTOR_AVAILABILITY"
    REQUIRED_TAG_MISSING = "REQUIRED_TAG_MISSING"
    RESTRICTED_TAG_CONFLICT = "RESTRICTED_TAG_CONFLICT"
    TUTOR_CAPACITY_EXCEEDED = "TUTOR_CAPACITY_EXCEEDED"
    TUTOR_OVERLAP = "TUTOR_OVERLAP"
    STUDENT_OVERLAP = "STUDENT_OVERLAP"
    NOT_FEASIBLE_CANDIDATE = "NOT_FEASIBLE_CANDIDATE"
    TENTATIVE_AVAILABILITY = "TENTATIVE_AVAILABILITY"


@dataclass(frozen=True)
class TutorPlacementWarning:
    code: TutorPlacementWarningCode
    message: str


@dataclass(frozen=True)
class TutorScheduleEntry:
    student_id: str
    student_name: str
    placement_type: str | None
    tutor_id: str | None
    tutor_name: str | None
    day: str | None
    start_minute: int | None
    end_minute: int | None
    preferred_tag_match_count: int
    tentative_minutes: int
    verified: bool | None
    warnings: tuple[TutorPlacementWarning, ...]
    unplaced_reason: str | None


@dataclass(frozen=True)
class TutorPlacementVerification:
    student_id: str
    student_name: str
    tutor_id: str | None
    tutor_name: str | None
    day: str | None
    start_minute: int | None
    warnings: tuple[TutorPlacementWarning, ...]


@dataclass(frozen=True)
class TutorStoredAssignment:
    student_id: str
    placement_type: str
    tutor_id: str
    tutor_name: str
    student_name: str
    day: str
    start_minute: int
    end_minute: int
    preferred_tag_match_count: int
    tentative_minutes: int


@dataclass(frozen=True)
class TutorStoredUnplaced:
    student_id: str
    student_name: str
    placement_type: str
    reason: str


@dataclass(frozen=True)
class TutorScheduleStatistics:
    students_scheduled: int
    students_unscheduled: int
    tutor_utilization: float
    preferred_tag_matches: int
    tentative_minutes: int
    gap_cost_minutes: int


@dataclass(frozen=True)
class TutorOptimizationResult:
    source_data_revision: int
    placements: tuple[TutorScheduledPlacement, ...]
    unplaced_students: tuple[TutorUnplacedStudentResult, ...]
    statistics: TutorScheduleStatistics


@dataclass(frozen=True)
class TutorSchedule:
    schedule_revision: int
    source_data_revision: int
    is_schedule_stale: bool
    has_schedule: bool
    placements: tuple[TutorScheduledPlacement, ...]
    unplaced_students: tuple[TutorUnplacedStudentResult, ...]
    statistics: TutorScheduleStatistics
    entries: tuple[TutorScheduleEntry, ...] = ()
    verification: tuple[TutorPlacementVerification, ...] = ()
    automatic_count: int = 0
    manual_count: int = 0
    manually_unplaced_count: int = 0
    verified_count: int = 0
    unverified_count: int = 0
