"""Lesson Scheduling value types shared by persistence and the pure solver layer.

This module must stay free of SQLAlchemy and other infrastructure imports so the
candidate, readiness, optimizer, and placement-review code can depend on it
without depending on persistence.
"""

from enum import Enum


class LessonTagState(str, Enum):
    NEUTRAL = "neutral"
    REQUIRED = "required"
    PREFERRED = "preferred"
    RESTRICTED = "restricted"


class LessonPlacementType(str, Enum):
    AUTOMATIC = "automatic"
    MANUAL = "manual"
    MANUALLY_UNPLACED = "manuallyUnplaced"


class LessonPlacementState(str, Enum):
    """Coordinator-facing state of a student: no placement record is UNASSIGNED."""

    UNASSIGNED = "unassigned"
    ASSIGNED = "assigned"
    MANUALLY_UNPLACED = "manuallyUnplaced"

    @classmethod
    def from_placement_type(cls, placement_type: str | None) -> "LessonPlacementState":
        if placement_type is None:
            return cls.UNASSIGNED
        if placement_type == LessonPlacementType.MANUALLY_UNPLACED.value:
            return cls.MANUALLY_UNPLACED
        return cls.ASSIGNED


class LessonUnplacedReason(str, Enum):
    NO_FEASIBLE_CANDIDATE = "NO_FEASIBLE_CANDIDATE"
    CAPACITY_LIMIT = "CAPACITY_LIMIT"
    INSTRUCTOR_OVERLAP = "INSTRUCTOR_OVERLAP"
    MANUALLY_UNPLACED = "MANUALLY_UNPLACED"
