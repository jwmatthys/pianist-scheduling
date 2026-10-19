"""Typed Lesson Scheduling inputs, candidates, readiness, and schedule results."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from ..lesson_models import LessonTagState, LessonUnplacedReason
from .availability import AvailabilityWindow, WEEKDAYS

LESSON_START_INCREMENT_MINUTES = 15
MINUTES_PER_DAY = 24 * 60


@dataclass(frozen=True)
class LessonInstructorInput:
    id: str
    name: str
    lesson_length_minutes: int
    capacity: int
    tag_ids: frozenset[str]
    availability: tuple[AvailabilityWindow, ...]


@dataclass(frozen=True)
class LessonStudentInput:
    id: str
    name: str
    availability: tuple[AvailabilityWindow, ...]
    # Neutral is represented by an absent entry.
    tag_states: tuple[tuple[str, LessonTagState], ...]


@dataclass(frozen=True)
class LessonTagInput:
    id: str
    name: str


@dataclass(frozen=True)
class LessonSchedulingInput:
    """Lesson-owned scheduling inputs captured at one source-data revision."""

    source_data_revision: int
    students: tuple[LessonStudentInput, ...]
    instructors: tuple[LessonInstructorInput, ...]
    tags: tuple[LessonTagInput, ...] = ()


class LessonCandidateIssueCode(str, Enum):
    REQUIRED_TAG_MISSING = "REQUIRED_TAG_MISSING"
    RESTRICTED_TAG_PRESENT = "RESTRICTED_TAG_PRESENT"


@dataclass(frozen=True)
class LessonCandidateIssue:
    code: LessonCandidateIssueCode
    tag_ids: tuple[str, ...] = ()


@dataclass(frozen=True)
class LessonCandidate:
    student_id: str
    instructor_id: str
    day: str
    lesson_start_minute: int
    lesson_end_minute: int
    preferred_tag_match_count: int
    tentative_minutes: int
    issues: tuple[LessonCandidateIssue, ...]
    is_feasible: bool
    # Exposed to future scheduling layers; candidate generation never enforces it.
    instructor_capacity: int


@dataclass(frozen=True)
class LessonCandidateSet:
    source_data_revision: int
    candidates: tuple[LessonCandidate, ...]

    @property
    def feasible(self) -> tuple[LessonCandidate, ...]:
        return tuple(candidate for candidate in self.candidates if candidate.is_feasible)

    def for_student(self, student_id: str) -> tuple[LessonCandidate, ...]:
        return tuple(candidate for candidate in self.candidates if candidate.student_id == student_id)

    def feasible_for_student(self, student_id: str) -> tuple[LessonCandidate, ...]:
        return tuple(candidate for candidate in self.for_student(student_id) if candidate.is_feasible)


class LessonReadinessSeverity(str, Enum):
    ERROR = "error"
    WARNING = "warning"


class LessonReadinessCategory(str, Enum):
    DATA = "data"
    STUDENTS = "students"
    INSTRUCTORS = "instructors"
    TAGS = "tags"


class LessonReadinessSubjectKind(str, Enum):
    STUDENT = "student"
    INSTRUCTOR = "instructor"
    TAG = "tag"


class LessonReadinessCode(str, Enum):
    NO_STUDENTS = "NO_STUDENTS"
    NO_INSTRUCTORS = "NO_INSTRUCTORS"
    NO_FEASIBLE_CANDIDATES = "NO_FEASIBLE_CANDIDATES"
    STUDENT_WITHOUT_AVAILABILITY = "STUDENT_WITHOUT_AVAILABILITY"
    INSTRUCTOR_WITHOUT_AVAILABILITY = "INSTRUCTOR_WITHOUT_AVAILABILITY"
    STUDENT_WITHOUT_CANDIDATES = "STUDENT_WITHOUT_CANDIDATES"
    REQUIRED_TAG_WITHOUT_MATCH = "REQUIRED_TAG_WITHOUT_MATCH"
    ZERO_CAPACITY = "ZERO_CAPACITY"
    INVALID_INSTRUCTOR_WINDOW = "INVALID_INSTRUCTOR_WINDOW"
    INSTRUCTOR_WITHOUT_SCHEDULABLE_LESSONS = "INSTRUCTOR_WITHOUT_SCHEDULABLE_LESSONS"


@dataclass(frozen=True)
class LessonReadinessSubject:
    kind: LessonReadinessSubjectKind
    id: str
    label: str


@dataclass(frozen=True)
class LessonReadinessWarning:
    code: LessonReadinessCode
    severity: LessonReadinessSeverity
    category: LessonReadinessCategory
    subjects: tuple[LessonReadinessSubject, ...] = ()


@dataclass(frozen=True)
class LessonReadinessReport:
    source_data_revision: int
    is_ready: bool
    student_count: int
    instructor_count: int
    feasible_candidate_count: int
    total_candidate_count: int
    students_with_candidates: int
    instructors_with_schedulable_lessons: int
    warnings: tuple[LessonReadinessWarning, ...]


@dataclass(frozen=True)
class LessonScheduledPlacement:
    student_id: str
    instructor_id: str
    day: str
    start_minute: int
    end_minute: int
    preferred_tag_match_count: int
    tentative_minutes: int


@dataclass(frozen=True)
class LessonUnplacedStudentResult:
    student_id: str
    reason: LessonUnplacedReason


class LessonPlacementWarningCode(str, Enum):
    STUDENT_REMOVED = "STUDENT_REMOVED"
    INSTRUCTOR_REMOVED = "INSTRUCTOR_REMOVED"
    INVALID_TIME = "INVALID_TIME"
    LESSON_LENGTH_CHANGED = "LESSON_LENGTH_CHANGED"
    OUTSIDE_STUDENT_AVAILABILITY = "OUTSIDE_STUDENT_AVAILABILITY"
    OUTSIDE_INSTRUCTOR_AVAILABILITY = "OUTSIDE_INSTRUCTOR_AVAILABILITY"
    REQUIRED_TAG_MISSING = "REQUIRED_TAG_MISSING"
    RESTRICTED_TAG_CONFLICT = "RESTRICTED_TAG_CONFLICT"
    INSTRUCTOR_CAPACITY_EXCEEDED = "INSTRUCTOR_CAPACITY_EXCEEDED"
    INSTRUCTOR_OVERLAP = "INSTRUCTOR_OVERLAP"
    NOT_FEASIBLE_CANDIDATE = "NOT_FEASIBLE_CANDIDATE"
    TENTATIVE_AVAILABILITY = "TENTATIVE_AVAILABILITY"


@dataclass(frozen=True)
class LessonPlacementWarning:
    code: LessonPlacementWarningCode
    message: str


@dataclass(frozen=True)
class LessonScheduleEntry:
    student_id: str
    student_name: str
    placement_type: str | None
    instructor_id: str | None
    instructor_name: str | None
    day: str | None
    start_minute: int | None
    end_minute: int | None
    preferred_tag_match_count: int
    tentative_minutes: int
    verified: bool | None
    warnings: tuple[LessonPlacementWarning, ...]
    unplaced_reason: str | None


@dataclass(frozen=True)
class LessonPlacementVerification:
    student_id: str
    student_name: str
    instructor_id: str | None
    instructor_name: str | None
    day: str | None
    start_minute: int | None
    warnings: tuple[LessonPlacementWarning, ...]


@dataclass(frozen=True)
class LessonStoredAssignment:
    student_id: str
    placement_type: str
    instructor_id: str
    instructor_name: str
    student_name: str
    day: str
    start_minute: int
    end_minute: int
    preferred_tag_match_count: int
    tentative_minutes: int


@dataclass(frozen=True)
class LessonStoredUnplaced:
    student_id: str
    student_name: str
    placement_type: str
    reason: str


@dataclass(frozen=True)
class LessonScheduleStatistics:
    students_scheduled: int
    students_unscheduled: int
    instructor_utilization: float
    preferred_tag_matches: int
    tentative_minutes: int
    gap_cost_minutes: int


@dataclass(frozen=True)
class LessonOptimizationResult:
    source_data_revision: int
    placements: tuple[LessonScheduledPlacement, ...]
    unplaced_students: tuple[LessonUnplacedStudentResult, ...]
    statistics: LessonScheduleStatistics


@dataclass(frozen=True)
class LessonSchedule:
    schedule_revision: int
    source_data_revision: int
    is_schedule_stale: bool
    has_schedule: bool
    placements: tuple[LessonScheduledPlacement, ...]
    unplaced_students: tuple[LessonUnplacedStudentResult, ...]
    statistics: LessonScheduleStatistics
    entries: tuple[LessonScheduleEntry, ...] = ()
    verification: tuple[LessonPlacementVerification, ...] = ()
    automatic_count: int = 0
    manual_count: int = 0
    manually_unplaced_count: int = 0
    verified_count: int = 0
    unverified_count: int = 0
