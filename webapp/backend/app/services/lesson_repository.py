"""Lesson Scheduling persistence and typed orchestration repository.

The repository owns roster, availability, placement, and schedule persistence,
revision/staleness state, and orchestration over pure candidate, readiness, and
optimizer services.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Iterable, Mapping, Sequence

from sqlalchemy import func
from sqlalchemy.orm import Session

from .. import models
from ..lesson_models import (
    DEFAULT_INSTRUCTOR_CAPACITY,
    DEFAULT_LESSON_LENGTH_MINUTES,
    LessonInstructor,
    LessonInstructorAvailabilityWindow,
    LessonInstructorTag,
    LessonModuleState,
    LessonPlacement,
    LessonPlacementType,
    LessonSchedulePlacement,
    LessonScheduleResult,
    LessonStudent,
    LessonStudentAvailabilityWindow,
    LessonStudentTagState,
    LessonTag,
    LessonTagState,
    LessonUnplacedReason,
    LessonUnplacedStudent,
    lesson_tag_identity_key,
    normalize_lesson_tag_name,
)
from .availability import AvailabilityWindow, WEEKDAYS, normalize_day, normalize_windows
from .lesson_candidates import generate_lesson_candidates
from .lesson_scheduling_models import (
    LessonCandidate,
    LessonCandidateSet,
    LessonInstructorInput,
    LessonOptimizationResult,
    LessonPlacementVerification,
    LessonReadinessReport,
    LessonSchedule,
    LessonScheduleEntry,
    LessonScheduleStatistics,
    LessonScheduledPlacement,
    LessonSchedulingInput,
    LessonStoredAssignment,
    LessonStoredUnplaced,
    LessonStudentInput,
    LessonTagInput,
    LessonUnplacedStudentResult,
)
from .lesson_readiness import analyze_lesson_readiness
from .lesson_optimizer import (
    _gap_cost,
    _schedulable_lesson_capacity,
    _usable_coverage,
    optimize_lesson_schedule,
)
from .lesson_placement_review import review_lesson_schedule


class LessonValidationError(ValueError):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


@dataclass(frozen=True)
class InstructorInput:
    name: str
    email: str = ""
    room: str = ""
    lesson_length_minutes: int = DEFAULT_LESSON_LENGTH_MINUTES
    capacity: int = DEFAULT_INSTRUCTOR_CAPACITY


@dataclass(frozen=True)
class LessonTagSummary:
    tag: LessonTag
    instructor_count: int
    student_count: int

    @property
    def is_unused(self) -> bool:
        return self.instructor_count == 0 and self.student_count == 0


@dataclass(frozen=True)
class StudentInput:
    student_id: str
    name: str
    email: str = ""


@dataclass(frozen=True)
class LessonStudentReplaceResult:
    students_removed: int
    placements_removed: int
    placements_preserved: int
    students_created: int


@dataclass(frozen=True)
class LessonReplaceResult:
    instructors_removed: int
    placements_removed: int
    placements_preserved: int
    instructors_created: int
    tags_created: int


def _now() -> datetime:
    return datetime.utcnow()


def _is_int(value: object) -> bool:
    return isinstance(value, int) and not isinstance(value, bool)


class LessonRepository:
    def __init__(self, db: Session):
        self.db = db

    # --- module state -------------------------------------------------

    def _session_uuid(self) -> str:
        rows = self.db.query(models.SchedulingSession.session_uuid).all()
        if len(rows) != 1:
            raise ValueError("The active database must contain exactly one Scheduling Session.")
        return rows[0][0]

    def get_state(self) -> LessonModuleState:
        session_uuid = self._session_uuid()
        state = self.db.get(LessonModuleState, session_uuid)
        if state is None:
            state = LessonModuleState(
                session_uuid=session_uuid,
                schedule_revision=0,
                source_data_revision=0,
                is_schedule_stale=False,
                modified_at=_now(),
            )
            self.db.add(state)
            self.db.flush()
        return state

    def record_source_change(self, *, schedule_exists: bool | None = None) -> LessonModuleState:
        """Bump the source revision; flag an existing schedule stale without altering it."""
        state = self.get_state()
        state.source_data_revision += 1
        if schedule_exists is None:
            schedule_exists = self._has_schedule_data()
        if schedule_exists:
            state.is_schedule_stale = True
        state.modified_at = _now()
        self.db.flush()
        return state

    def _record_schedule_edit(self) -> None:
        state = self.get_state()
        state.schedule_revision += 1
        state.modified_at = _now()
        self.db.flush()

    def _has_schedule_data(self) -> bool:
        return (
            self.db.query(LessonPlacement.id).first() is not None
            or self.db.query(LessonScheduleResult.session_uuid).first() is not None
        )

    # --- tags ---------------------------------------------------------

    def list_tags(self) -> list[LessonTag]:
        return self.db.query(LessonTag).order_by(LessonTag.normalized_name).all()

    def find_tag(self, name: str) -> LessonTag | None:
        key = lesson_tag_identity_key(name)
        if not key:
            return None
        return self.db.query(LessonTag).filter(LessonTag.normalized_name == key).one_or_none()

    def _clean_tag_name(self, name: str) -> str:
        cleaned = normalize_lesson_tag_name(name)
        if not cleaned:
            raise LessonValidationError("INVALID_TAG", "A tag name is required.")
        return cleaned

    def create_tag(self, name: str) -> LessonTag:
        cleaned = self._clean_tag_name(name)
        if self.find_tag(cleaned) is not None:
            raise LessonValidationError("DUPLICATE_TAG", f"The tag '{cleaned}' already exists.")
        tag = self._insert_tag(cleaned)
        self.record_source_change()
        return tag

    def _insert_tag(self, cleaned_name: str) -> LessonTag:
        tag = LessonTag(name=cleaned_name, normalized_name=lesson_tag_identity_key(cleaned_name))
        self.db.add(tag)
        self.db.flush()
        return tag

    def tag_summaries(self) -> list["LessonTagSummary"]:
        """Every registered tag with current-use information for the Tag Manager."""
        instructor_counts = dict(
            self.db.query(LessonInstructorTag.tag_id, func.count(LessonInstructorTag.instructor_id))
            .group_by(LessonInstructorTag.tag_id)
            .all()
        )
        student_counts = dict(
            self.db.query(LessonStudentTagState.tag_id, func.count(func.distinct(LessonStudentTagState.student_id)))
            .filter(LessonStudentTagState.state != LessonTagState.NEUTRAL.value)
            .group_by(LessonStudentTagState.tag_id)
            .all()
        )
        return [
            LessonTagSummary(
                tag=tag,
                instructor_count=instructor_counts.get(tag.id, 0),
                student_count=student_counts.get(tag.id, 0),
            )
            for tag in self.list_tags()
        ]

    def rename_tag(self, tag_id: str, name: str) -> LessonTag:
        tag = self._require(LessonTag, tag_id, "TAG_NOT_FOUND", "The tag no longer exists.")
        cleaned = self._clean_tag_name(name)
        existing = self.find_tag(cleaned)
        if existing is not None and existing.id != tag.id:
            raise LessonValidationError("DUPLICATE_TAG", f"The tag '{cleaned}' already exists.")
        tag.name = cleaned
        tag.normalized_name = lesson_tag_identity_key(cleaned)
        tag.updated_at = _now()
        self.db.flush()
        self.record_source_change()
        return tag

    def delete_tag(self, tag_id: str) -> None:
        """Delete a tag and its Lesson Scheduling associations."""
        tag = self._require(LessonTag, tag_id, "TAG_NOT_FOUND", "The tag no longer exists.")
        self.db.query(LessonInstructorTag).filter(LessonInstructorTag.tag_id == tag.id).delete()
        self.db.query(LessonStudentTagState).filter(LessonStudentTagState.tag_id == tag.id).delete()
        self.db.delete(tag)
        self.db.flush()
        self.record_source_change()

    # --- instructors --------------------------------------------------

    def _clean_instructor_values(
        self, name: str, email: str, room: str, lesson_length_minutes: int, capacity: int
    ) -> dict:
        cleaned_name = (name or "").strip()
        if not cleaned_name:
            raise LessonValidationError("INVALID_INSTRUCTOR_NAME", "Instructor name is required.")
        if not _is_int(lesson_length_minutes) or lesson_length_minutes <= 0:
            raise LessonValidationError("INVALID_LESSON_LENGTH", "Lesson length must be a positive number of minutes.")
        if not _is_int(capacity) or capacity < 1:
            raise LessonValidationError("INVALID_CAPACITY", "Capacity must be at least 1.")
        return {
            "name": cleaned_name,
            "email": (email or "").strip(),
            "room": (room or "").strip(),
            "lesson_length_minutes": lesson_length_minutes,
            "capacity": capacity,
        }

    def create_instructor(
        self,
        name: str,
        email: str = "",
        room: str = "",
        lesson_length_minutes: int = DEFAULT_LESSON_LENGTH_MINUTES,
        capacity: int = DEFAULT_INSTRUCTOR_CAPACITY,
        availability: Sequence[AvailabilityWindow] = (),
        tag_ids: Iterable[str] = (),
    ) -> LessonInstructor:
        values = self._clean_instructor_values(name, email, room, lesson_length_minutes, capacity)
        instructor = LessonInstructor(**values)
        self.db.add(instructor)
        self.db.flush()
        self._replace_instructor_availability(instructor.id, availability)
        self._replace_instructor_tags(instructor, tag_ids)
        self.record_source_change()
        return instructor

    def get_instructor(self, instructor_id: str) -> LessonInstructor:
        return self._require(LessonInstructor, instructor_id, "INSTRUCTOR_NOT_FOUND", "The instructor no longer exists.")

    def list_instructors(self) -> list[LessonInstructor]:
        return self.db.query(LessonInstructor).order_by(func.lower(LessonInstructor.name), LessonInstructor.id).all()

    def update_instructor(
        self,
        instructor_id: str,
        name: str,
        email: str = "",
        room: str = "",
        lesson_length_minutes: int = DEFAULT_LESSON_LENGTH_MINUTES,
        capacity: int = DEFAULT_INSTRUCTOR_CAPACITY,
        tag_ids: Iterable[str] | None = None,
    ) -> LessonInstructor:
        """Update fields (and tags when given) as a single source-data change."""
        instructor = self.get_instructor(instructor_id)
        values = self._clean_instructor_values(name, email, room, lesson_length_minutes, capacity)
        if tag_ids is not None:
            self._replace_instructor_tags(instructor, tag_ids)
        for field, value in values.items():
            setattr(instructor, field, value)
        instructor.updated_at = _now()
        self.db.flush()
        self.record_source_change()
        return instructor

    def set_instructor_tags(self, instructor_id: str, tag_ids: Iterable[str]) -> LessonInstructor:
        instructor = self.get_instructor(instructor_id)
        self._replace_instructor_tags(instructor, tag_ids)
        self.record_source_change()
        return instructor

    def delete_instructor(self, instructor_id: str) -> int:
        """Delete an instructor without deleting schedule decisions that reference it."""
        instructor = self.get_instructor(instructor_id)
        self.db.query(LessonInstructorTag).filter(LessonInstructorTag.instructor_id == instructor.id).delete()
        self.db.query(LessonInstructorAvailabilityWindow).filter(
            LessonInstructorAvailabilityWindow.instructor_id == instructor.id
        ).delete()
        self.db.delete(instructor)
        self.db.flush()
        self.record_source_change()
        return 0

    def _replace_instructor_tags(self, instructor: LessonInstructor, tag_ids: Iterable[str]) -> None:
        wanted = list(dict.fromkeys(tag_ids))
        for tag_id in wanted:
            self._require(LessonTag, tag_id, "TAG_NOT_FOUND", "A selected tag no longer exists.")
        self.db.query(LessonInstructorTag).filter(LessonInstructorTag.instructor_id == instructor.id).delete()
        for tag_id in wanted:
            self.db.add(LessonInstructorTag(instructor_id=instructor.id, tag_id=tag_id))
        self.db.flush()
        self.db.expire(instructor, ["tags"])

    def replace_instructors(
        self, instructors: Iterable[tuple[InstructorInput, Sequence[AvailabilityWindow], Iterable[str]]]
    ) -> "LessonReplaceResult":
        """Replace instructors without deleting schedule decisions."""
        prepared = [
            (
                self._clean_instructor_values(
                    values.name, values.email, values.room, values.lesson_length_minutes, values.capacity
                ),
                self._validated_windows(windows),
                list(names),
            )
            for values, windows, names in instructors
        ]
        instructors_removed = self.db.query(func.count(LessonInstructor.id)).scalar() or 0
        placements_removed = 0
        placements_preserved = self.db.query(func.count(LessonPlacement.id)).scalar() or 0
        self.db.query(LessonInstructorTag).delete()
        self.db.query(LessonInstructorAvailabilityWindow).delete()
        self.db.query(LessonInstructor).delete()
        self.db.flush()

        tags: dict[str, LessonTag] = {}
        tags_created = 0
        for values, windows, names in prepared:
            instructor = LessonInstructor(**values)
            self.db.add(instructor)
            self.db.flush()
            self._replace_instructor_availability(instructor.id, windows)
            linked: set[str] = set()
            for name in names:
                key = lesson_tag_identity_key(name)
                if not key or key in linked:
                    continue
                tag = tags.get(key) or self.find_tag(name)
                if tag is None:
                    tag = self._insert_tag(normalize_lesson_tag_name(name))
                    tags_created += 1
                tags[key] = tag
                linked.add(key)
                self.db.add(LessonInstructorTag(instructor_id=instructor.id, tag_id=tag.id))
            self.db.flush()
        self.record_source_change()
        return LessonReplaceResult(
            instructors_removed=instructors_removed,
            placements_removed=placements_removed,
            placements_preserved=placements_preserved,
            instructors_created=len(prepared),
            tags_created=tags_created,
        )

    def instructor_availability_counts(self) -> dict[str, int]:
        return dict(
            self.db.query(
                LessonInstructorAvailabilityWindow.instructor_id, func.count(LessonInstructorAvailabilityWindow.id)
            )
            .group_by(LessonInstructorAvailabilityWindow.instructor_id)
            .all()
        )

    # --- availability -------------------------------------------------

    @staticmethod
    def _validated_windows(windows: Sequence[AvailabilityWindow]) -> list[AvailabilityWindow]:
        """Merge same-status overlaps; reject invalid or conflicting windows so nothing partial is stored."""
        for window in windows:
            if normalize_day(window.day) != window.day or not 0 <= window.start_minute < window.end_minute <= 24 * 60:
                raise LessonValidationError("INVALID_AVAILABILITY_WINDOW", "An Availability Window is invalid.")
        normalized, issues = normalize_windows(list(windows))
        if any(issue.severity == "error" for issue in issues):
            raise LessonValidationError(
                "INVALID_AVAILABILITY_WINDOW",
                "Availability Windows must be valid and must not overlap with a different status.",
            )
        return normalized

    def _replace_instructor_availability(self, instructor_id: str, windows: Sequence[AvailabilityWindow]) -> None:
        validated = self._validated_windows(windows)
        self.db.query(LessonInstructorAvailabilityWindow).filter(
            LessonInstructorAvailabilityWindow.instructor_id == instructor_id
        ).delete()
        for window in validated:
            self.db.add(LessonInstructorAvailabilityWindow(
                instructor_id=instructor_id,
                day=window.day,
                start_minute=window.start_minute,
                end_minute=window.end_minute,
                status=window.status,
                source=window.source,
                source_row=window.source_row,
            ))
        self.db.flush()

    def set_instructor_availability(
        self, instructor_id: str, windows: Sequence[AvailabilityWindow]
    ) -> list[AvailabilityWindow]:
        """Replace one instructor's recurring weekly windows as a single source-data change."""
        self.get_instructor(instructor_id)
        self._replace_instructor_availability(instructor_id, windows)
        self.record_source_change()
        return self.instructor_availability(instructor_id)

    def instructor_availability(self, instructor_id: str) -> list[AvailabilityWindow]:
        rows = (
            self.db.query(LessonInstructorAvailabilityWindow)
            .filter(LessonInstructorAvailabilityWindow.instructor_id == instructor_id)
            .order_by(LessonInstructorAvailabilityWindow.day, LessonInstructorAvailabilityWindow.start_minute)
            .all()
        )
        return [_window(row) for row in rows]

    def _replace_student_availability(self, lesson_student_id: str, windows: Sequence[AvailabilityWindow]) -> None:
        validated = self._validated_windows(windows)
        self.db.query(LessonStudentAvailabilityWindow).filter(
            LessonStudentAvailabilityWindow.lesson_student_id == lesson_student_id
        ).delete()
        for window in validated:
            self.db.add(LessonStudentAvailabilityWindow(
                lesson_student_id=lesson_student_id,
                day=window.day,
                start_minute=window.start_minute,
                end_minute=window.end_minute,
                status=window.status,
                source=window.source,
                source_row=window.source_row,
            ))
        self.db.flush()

    def set_student_availability(
        self, lesson_student_id: str, windows: Sequence[AvailabilityWindow]
    ) -> list[AvailabilityWindow]:
        """Replace one student's recurring weekly windows as a single source-data change."""
        self.get_student(lesson_student_id)
        self._replace_student_availability(lesson_student_id, windows)
        self.record_source_change()
        return self.student_availability(lesson_student_id)

    def student_availability(self, lesson_student_id: str) -> list[AvailabilityWindow]:
        rows = (
            self.db.query(LessonStudentAvailabilityWindow)
            .filter(LessonStudentAvailabilityWindow.lesson_student_id == lesson_student_id)
            .order_by(LessonStudentAvailabilityWindow.day, LessonStudentAvailabilityWindow.start_minute)
            .all()
        )
        return [_window(row) for row in rows]

    # --- students -----------------------------------------------------

    def _clean_student_values(
        self,
        student_id: str,
        name: str,
        email: str,
        exclude_id: str | None = None,
        check_unique: bool = True,
    ) -> tuple[str, str, str]:
        cleaned_id = (student_id or "").strip()
        if not cleaned_id:
            raise LessonValidationError("INVALID_STUDENT_ID", "Student ID is required.")
        if check_unique:
            query = self.db.query(LessonStudent.id).filter(func.lower(LessonStudent.student_id) == cleaned_id.lower())
            if exclude_id is not None:
                query = query.filter(LessonStudent.id != exclude_id)
            if query.first() is not None:
                raise LessonValidationError("DUPLICATE_STUDENT_ID", "A student with that Student ID already exists.")
        cleaned_name = " ".join((name or "").split())
        if not cleaned_name:
            raise LessonValidationError("INVALID_STUDENT_NAME", "Name is required.")
        return cleaned_id, cleaned_name, (email or "").strip()

    def create_student(
        self,
        student_id: str,
        name: str = "",
        email: str = "",
        availability: Sequence[AvailabilityWindow] = (),
        tag_states: Mapping[str, LessonTagState | str] | None = None,
    ) -> LessonStudent:
        cleaned_id, cleaned_name, cleaned_email = self._clean_student_values(
            student_id, name, email
        )
        validated_states = self._validated_tag_states(tag_states or {})
        student = LessonStudent(student_id=cleaned_id, name=cleaned_name, email=cleaned_email)
        self.db.add(student)
        self.db.flush()
        self._replace_student_availability(student.id, availability)
        self._replace_student_tag_states(student.id, validated_states)
        self.record_source_change()
        return student

    def get_student(self, lesson_student_id: str) -> LessonStudent:
        return self._require(LessonStudent, lesson_student_id, "STUDENT_NOT_FOUND", "The student no longer exists.")

    def list_students(self) -> list[LessonStudent]:
        return self.db.query(LessonStudent).order_by(func.lower(LessonStudent.name), LessonStudent.student_id).all()

    def update_student(
        self,
        lesson_student_id: str,
        student_id: str,
        name: str,
        email: str = "",
        tag_states: Mapping[str, LessonTagState | str] | None = None,
    ) -> LessonStudent:
        """Update fields (and, when given, the full tag-state set) as one source-data change."""
        student = self.get_student(lesson_student_id)
        cleaned = self._clean_student_values(
            student_id, name, email, exclude_id=student.id
        )
        validated_states = self._validated_tag_states(tag_states) if tag_states is not None else None
        student.student_id, student.name, student.email = cleaned
        student.updated_at = _now()
        self.db.flush()
        if validated_states is not None:
            self._replace_student_tag_states(student.id, validated_states)
        self.record_source_change()
        return student

    def delete_student(self, lesson_student_id: str) -> None:
        student = self.get_student(lesson_student_id)
        self.db.query(LessonStudentTagState).filter(LessonStudentTagState.student_id == student.id).delete()
        self.db.query(LessonStudentAvailabilityWindow).filter(
            LessonStudentAvailabilityWindow.lesson_student_id == student.id
        ).delete()
        self.db.delete(student)
        self.db.flush()
        self.record_source_change()

    # --- student tag states -------------------------------------------

    def student_tag_states(self) -> dict[str, dict[str, LessonTagState]]:
        """Stored non-neutral tag states keyed by student then tag (Neutral is never stored)."""
        states: dict[str, dict[str, LessonTagState]] = {}
        for row in self.db.query(LessonStudentTagState).all():
            states.setdefault(row.student_id, {})[row.tag_id] = LessonTagState(row.state)
        return states

    def get_scheduling_input(self) -> LessonSchedulingInput:
        """Return a typed Lesson scheduling snapshot at the current source revision."""
        source_data_revision = self.get_state().source_data_revision
        instructors = self.list_instructors()
        students = self.list_students()

        instructor_ids = [item.id for item in instructors]
        student_ids = [item.id for item in students]
        instructor_windows: dict[str, list[AvailabilityWindow]] = {key: [] for key in instructor_ids}
        student_windows: dict[str, list[AvailabilityWindow]] = {key: [] for key in student_ids}
        instructor_tags: dict[str, set[str]] = {key: set() for key in instructor_ids}
        student_states: dict[str, list[tuple[str, LessonTagState]]] = {key: [] for key in student_ids}

        for row in (
            self.db.query(LessonInstructorAvailabilityWindow)
            .order_by(LessonInstructorAvailabilityWindow.day, LessonInstructorAvailabilityWindow.start_minute)
            .all()
        ):
            if row.instructor_id in instructor_windows:
                instructor_windows[row.instructor_id].append(_window(row))
        for row in (
            self.db.query(LessonStudentAvailabilityWindow)
            .order_by(LessonStudentAvailabilityWindow.day, LessonStudentAvailabilityWindow.start_minute)
            .all()
        ):
            if row.lesson_student_id in student_windows:
                student_windows[row.lesson_student_id].append(_window(row))
        for row in self.db.query(LessonInstructorTag).all():
            if row.instructor_id in instructor_tags:
                instructor_tags[row.instructor_id].add(row.tag_id)
        for row in self.db.query(LessonStudentTagState).all():
            if row.student_id in student_states:
                student_states[row.student_id].append((row.tag_id, LessonTagState(row.state)))

        return LessonSchedulingInput(
            source_data_revision=source_data_revision,
            students=tuple(
                LessonStudentInput(
                    id=student.id,
                    name=student.name,
                    availability=tuple(student_windows[student.id]),
                    tag_states=tuple(sorted(student_states[student.id])),
                )
                for student in students
            ),
            instructors=tuple(
                LessonInstructorInput(
                    id=instructor.id,
                    name=instructor.name,
                    lesson_length_minutes=instructor.lesson_length_minutes,
                    capacity=instructor.capacity,
                    tag_ids=frozenset(instructor_tags[instructor.id]),
                    availability=tuple(instructor_windows[instructor.id]),
                )
                for instructor in instructors
            ),
            tags=tuple(
                LessonTagInput(id=tag.id, name=tag.name)
                for tag in self.list_tags()
            ),
        )

    def generate_candidate_set(self) -> LessonCandidateSet:
        """Generate Lesson candidates from the current typed snapshot."""
        return generate_lesson_candidates(self.get_scheduling_input())

    def get_readiness_report(self) -> LessonReadinessReport:
        """Orchestrate a revision-consistent snapshot, candidate set, and readiness analysis."""
        scheduling_input = self.get_scheduling_input()
        candidate_set = generate_lesson_candidates(scheduling_input)
        return analyze_lesson_readiness(scheduling_input, candidate_set)

    def _stored_assignments(
        self, scheduling_input: LessonSchedulingInput
    ) -> tuple[LessonStoredAssignment, ...]:
        students = {item.id: item for item in scheduling_input.students}
        instructors = {item.id: item for item in scheduling_input.instructors}
        previous = {
            item.student_id: item
            for item in self.db.query(LessonSchedulePlacement).all()
        }
        assignments: list[LessonStoredAssignment] = []
        for row in self.list_placements():
            if row.placement_type == LessonPlacementType.MANUALLY_UNPLACED.value:
                continue
            if row.lesson_instructor_id is None or row.start_minute is None:
                continue
            prior = previous.get(row.lesson_student_id)
            instructor = instructors.get(row.lesson_instructor_id)
            student = students.get(row.lesson_student_id)
            end_minute = row.end_minute or (
                prior.end_minute if prior is not None
                else row.start_minute + (
                    instructor.lesson_length_minutes if instructor is not None
                    else DEFAULT_LESSON_LENGTH_MINUTES
                )
            )
            preferred, tentative = self._placement_metrics(
                student, instructor, row.day, row.start_minute, end_minute
            ) if student is not None and instructor is not None else (
                prior.preferred_tag_match_count if prior is not None else 0,
                prior.tentative_minutes if prior is not None else 0,
            )
            assignments.append(LessonStoredAssignment(
                student_id=row.lesson_student_id,
                placement_type=row.placement_type,
                instructor_id=row.lesson_instructor_id,
                instructor_name=instructor.name if instructor is not None else (
                    prior.instructor_name if prior is not None else row.lesson_instructor_id
                ),
                student_name=student.name if student is not None else (
                    prior.student_name if prior is not None else row.lesson_student_id
                ),
                day=row.day,
                start_minute=row.start_minute,
                end_minute=end_minute,
                preferred_tag_match_count=preferred,
                tentative_minutes=tentative,
            ))
        return tuple(assignments)

    @staticmethod
    def _placement_metrics(student, instructor, day: str, start: int, end: int) -> tuple[int, int]:
        states = dict(student.tag_states)
        preferred = sum(
            state is LessonTagState.PREFERRED and tag_id in instructor.tag_ids
            for tag_id, state in states.items()
        )
        student_tentative = sum(
            max(0, min(end, window.end_minute) - max(start, window.start_minute))
            for window in student.availability
            if window.day == day and window.status == "Tentative"
        )
        instructor_tentative = sum(
            max(0, min(end, window.end_minute) - max(start, window.start_minute))
            for window in instructor.availability
            if window.day == day and window.status == "Tentative"
        )
        return preferred, student_tentative + instructor_tentative

    def _stored_unplaced(self) -> tuple[LessonStoredUnplaced, ...]:
        students = {item.id: item.name for item in self.list_students()}
        decisions = {
            item.lesson_student_id: item for item in self.list_placements()
            if item.placement_type == LessonPlacementType.MANUALLY_UNPLACED.value
        }
        stored = {
            item.student_id: item for item in self.db.query(LessonUnplacedStudent).all()
        }
        for student_id, row in decisions.items():
            snapshot = stored.get(student_id)
            stored[student_id] = LessonUnplacedStudent(
                student_id=student_id,
                student_name=students.get(student_id, snapshot.student_name if snapshot else student_id),
                reason=LessonUnplacedReason.MANUALLY_UNPLACED.value,
                placement_type=LessonPlacementType.MANUALLY_UNPLACED.value,
            )
        return tuple(
            LessonStoredUnplaced(
                student_id=row.student_id,
                student_name=students.get(row.student_id, row.student_name or row.student_id),
                placement_type=row.placement_type,
                reason=row.reason,
            )
            for row in stored.values()
        )

    def _schedule_review(
        self, scheduling_input: LessonSchedulingInput | None = None
    ) -> tuple[
        LessonSchedulingInput,
        tuple[LessonStoredAssignment, ...],
        tuple[LessonStoredUnplaced, ...],
        tuple[LessonScheduleEntry, ...],
        tuple[LessonPlacementVerification, ...],
    ]:
        scheduling_input = scheduling_input or self.get_scheduling_input()
        candidates = generate_lesson_candidates(scheduling_input)
        assignments = self._stored_assignments(scheduling_input)
        unplaced = self._stored_unplaced()
        entries, verification = review_lesson_schedule(
            scheduling_input, candidates, assignments, unplaced
        )
        return scheduling_input, assignments, unplaced, entries, verification

    def _schedule_statistics(
        self,
        scheduling_input: LessonSchedulingInput,
        assignments: Sequence[LessonStoredAssignment],
        entries: Sequence[LessonScheduleEntry],
    ) -> LessonScheduleStatistics:
        instructor_by_id = {item.id: item for item in scheduling_input.instructors}
        capacities = {
            item.id: _schedulable_lesson_capacity(item)
            for item in scheduling_input.instructors
        }
        candidate_rows = []
        for item in assignments:
            instructor = instructor_by_id.get(item.instructor_id)
            if instructor is None or item.day not in WEEKDAYS:
                continue
            candidate_rows.append(LessonCandidate(
                student_id=item.student_id,
                instructor_id=item.instructor_id,
                day=item.day,
                lesson_start_minute=item.start_minute,
                lesson_end_minute=item.end_minute,
                preferred_tag_match_count=item.preferred_tag_match_count,
                tentative_minutes=item.tentative_minutes,
                issues=(),
                is_feasible=True,
                instructor_capacity=instructor.capacity,
            ))
        coverage = {
            item.id: _usable_coverage(item.availability)
            for item in scheduling_input.instructors
        }
        placed_count = len(assignments)
        return LessonScheduleStatistics(
            students_scheduled=placed_count,
            students_unscheduled=max(0, len(entries) - placed_count),
            instructor_utilization=placed_count / sum(capacities.values()) if sum(capacities.values()) else 0.0,
            preferred_tag_matches=sum(item.preferred_tag_match_count for item in assignments),
            tentative_minutes=sum(item.tentative_minutes for item in assignments),
            gap_cost_minutes=_gap_cost(candidate_rows, coverage),
        )

    def _refresh_schedule_snapshots(
        self,
        scheduling_input: LessonSchedulingInput,
        *,
        automatic_unplaced: Sequence[LessonUnplacedStudentResult] | None = None,
    ) -> tuple[LessonStoredAssignment, ...]:
        old_assignments = {
            item.student_id: item for item in self.db.query(LessonSchedulePlacement).all()
        }
        old_unplaced = {
            item.student_id: item for item in self.db.query(LessonUnplacedStudent).all()
        }
        students = {item.id: item for item in scheduling_input.students}
        instructors = {item.id: item for item in scheduling_input.instructors}
        rows: list[LessonStoredAssignment] = []
        self.db.query(LessonSchedulePlacement).delete()
        for decision in self.list_placements():
            if decision.placement_type == LessonPlacementType.MANUALLY_UNPLACED.value:
                continue
            if decision.lesson_instructor_id is None or decision.start_minute is None:
                continue
            old = old_assignments.get(decision.lesson_student_id)
            student = students.get(decision.lesson_student_id)
            instructor = instructors.get(decision.lesson_instructor_id)
            end_minute = decision.end_minute or (
                old.end_minute if old else decision.start_minute + (
                    instructor.lesson_length_minutes if instructor else DEFAULT_LESSON_LENGTH_MINUTES
                )
            )
            preferred, tentative = self._placement_metrics(
                student, instructor, decision.day, decision.start_minute, end_minute
            ) if student is not None and instructor is not None else (
                old.preferred_tag_match_count if old else 0,
                old.tentative_minutes if old else 0,
            )
            item = LessonStoredAssignment(
                student_id=decision.lesson_student_id,
                placement_type=decision.placement_type,
                instructor_id=decision.lesson_instructor_id,
                instructor_name=instructor.name if instructor else (
                    old.instructor_name if old else decision.lesson_instructor_id
                ),
                student_name=student.name if student else (
                    old.student_name if old else decision.lesson_student_id
                ),
                day=decision.day,
                start_minute=decision.start_minute,
                end_minute=end_minute,
                preferred_tag_match_count=preferred,
                tentative_minutes=tentative,
            )
            rows.append(item)
            self.db.add(LessonSchedulePlacement(
                student_id=item.student_id,
                instructor_id=item.instructor_id,
                placement_type=item.placement_type,
                student_name=item.student_name,
                instructor_name=item.instructor_name,
                day=item.day,
                start_minute=item.start_minute,
                end_minute=item.end_minute,
                preferred_tag_match_count=item.preferred_tag_match_count,
                tentative_minutes=item.tentative_minutes,
            ))

        automatic_unplaced_ids: set[str] = set()
        if automatic_unplaced is not None:
            self.db.query(LessonUnplacedStudent).filter(
                LessonUnplacedStudent.placement_type == LessonPlacementType.AUTOMATIC.value
            ).delete()
            for item in automatic_unplaced:
                student = students.get(item.student_id)
                automatic_unplaced_ids.add(item.student_id)
                self.db.add(LessonUnplacedStudent(
                    student_id=item.student_id,
                    student_name=student.name if student else (
                        old_unplaced[item.student_id].student_name
                        if item.student_id in old_unplaced else item.student_id
                    ),
                    reason=item.reason.value,
                    placement_type=LessonPlacementType.AUTOMATIC.value,
                ))

        assigned_ids = {item.student_id for item in rows}
        manual_unplaced_ids = {
            item.lesson_student_id for item in self.list_placements()
            if item.placement_type == LessonPlacementType.MANUALLY_UNPLACED.value
        }
        self.db.query(LessonUnplacedStudent).filter(
            LessonUnplacedStudent.placement_type == LessonPlacementType.MANUALLY_UNPLACED.value
        ).delete()
        if manual_unplaced_ids:
            # An earlier automatic unplaced row for the same student must be replaced.
            self.db.query(LessonUnplacedStudent).filter(
                LessonUnplacedStudent.student_id.in_(manual_unplaced_ids)
            ).delete(synchronize_session=False)
        for student_id in manual_unplaced_ids:
            previous = old_unplaced.get(student_id)
            student = students.get(student_id)
            self.db.add(LessonUnplacedStudent(
                student_id=student_id,
                student_name=student.name if student else (previous.student_name if previous else student_id),
                reason=LessonUnplacedReason.MANUALLY_UNPLACED.value,
                placement_type=LessonPlacementType.MANUALLY_UNPLACED.value,
            ))
        if automatic_unplaced is not None:
            decisions = {item.lesson_student_id for item in self.list_placements()}
            for student in scheduling_input.students:
                if (
                    student.id not in assigned_ids
                    and student.id not in manual_unplaced_ids
                    and student.id not in decisions
                    and student.id not in automatic_unplaced_ids
                ):
                    self.db.add(LessonUnplacedStudent(
                        student_id=student.id,
                        student_name=student.name,
                        reason=LessonUnplacedReason.NO_FEASIBLE_CANDIDATE.value,
                        placement_type=LessonPlacementType.AUTOMATIC.value,
                    ))
        self.db.flush()
        return tuple(rows)

    def generate_schedule(self) -> LessonSchedule:
        """Replace automatic assignments while preserving manual decisions."""
        scheduling_input = self.get_scheduling_input()
        candidate_set = generate_lesson_candidates(scheduling_input)
        readiness = analyze_lesson_readiness(scheduling_input, candidate_set)
        if not readiness.is_ready:
            raise LessonValidationError(
                "NOT_READY", "Lesson Scheduling inputs are not ready for schedule generation."
            )

        decisions = self.list_placements()
        manual_ids = {
            item.lesson_student_id for item in decisions
            if item.placement_type in {
                LessonPlacementType.MANUAL.value,
                LessonPlacementType.MANUALLY_UNPLACED.value,
            }
        }
        fixed = tuple(
            item for item in self._stored_assignments(scheduling_input)
            if item.placement_type == LessonPlacementType.MANUAL.value
        )
        manually_unplaced_ids = frozenset(
            item.lesson_student_id for item in decisions
            if item.placement_type == LessonPlacementType.MANUALLY_UNPLACED.value
        )
        optimized = optimize_lesson_schedule(
            scheduling_input,
            candidate_set,
            fixed_placements=tuple(
                LessonScheduledPlacement(
                    student_id=item.student_id,
                    instructor_id=item.instructor_id,
                    day=item.day,
                    start_minute=item.start_minute,
                    end_minute=item.end_minute,
                    preferred_tag_match_count=item.preferred_tag_match_count,
                    tentative_minutes=item.tentative_minutes,
                )
                for item in fixed
            ),
            excluded_student_ids=manually_unplaced_ids,
        )
        session_uuid = self._session_uuid()
        state = self.get_state()
        self.db.query(LessonPlacement).filter(
            LessonPlacement.placement_type == LessonPlacementType.AUTOMATIC.value
        ).delete()
        for placement in optimized.placements:
            if placement.student_id in manual_ids:
                continue
            self.db.add(LessonPlacement(
                lesson_student_id=placement.student_id,
                lesson_instructor_id=placement.instructor_id,
                day=placement.day,
                start_minute=placement.start_minute,
                end_minute=placement.end_minute,
                placement_type=LessonPlacementType.AUTOMATIC.value,
                created_at=_now(),
                updated_at=_now(),
            ))
        assignment_rows = self._refresh_schedule_snapshots(
            scheduling_input, automatic_unplaced=optimized.unplaced_students
        )
        _, _, _, entries, _ = self._schedule_review(scheduling_input)
        statistics = self._schedule_statistics(scheduling_input, assignment_rows, entries)
        optimized = LessonOptimizationResult(
            source_data_revision=optimized.source_data_revision,
            placements=optimized.placements,
            unplaced_students=optimized.unplaced_students,
            statistics=statistics,
        )
        self._save_schedule_result(optimized, state.schedule_revision + 1, session_uuid)
        state.schedule_revision += 1
        state.is_schedule_stale = False
        state.modified_at = _now()
        self.db.flush()
        return self.get_schedule()

    def _save_schedule_result(
        self, result: LessonOptimizationResult, schedule_revision: int, session_uuid: str
    ) -> None:
        stats = result.statistics
        row = self.db.get(LessonScheduleResult, session_uuid)
        if row is None:
            row = LessonScheduleResult(session_uuid=session_uuid)
            self.db.add(row)
        row.schedule_revision = schedule_revision
        row.source_data_revision = result.source_data_revision
        row.students_scheduled = stats.students_scheduled
        row.students_unscheduled = stats.students_unscheduled
        row.instructor_utilization = stats.instructor_utilization
        row.preferred_tag_matches = stats.preferred_tag_matches
        row.tentative_minutes = stats.tentative_minutes
        row.gap_cost_minutes = stats.gap_cost_minutes
        row.modified_at = _now()

    def get_schedule(self) -> LessonSchedule:
        """Retrieve assignments, status, verification, and summary from current snapshots."""
        state = self.get_state()
        result = self.db.get(LessonScheduleResult, self._session_uuid())
        scheduling_input, assignments, _, entries, verification = self._schedule_review()
        placements = tuple(
            LessonScheduledPlacement(
                student_id=item.student_id,
                instructor_id=item.instructor_id,
                day=item.day,
                start_minute=item.start_minute,
                end_minute=item.end_minute,
                preferred_tag_match_count=item.preferred_tag_match_count,
                tentative_minutes=item.tentative_minutes,
            )
            for item in assignments
        )
        unplaced_students = tuple(
            LessonUnplacedStudentResult(
                student_id=item.student_id,
                reason=LessonUnplacedReason(item.reason),
            )
            for item in self._stored_unplaced()
        )
        has_schedule = (
            result is not None
            or bool(self.list_placements())
            or bool(self.db.query(LessonUnplacedStudent.student_id).first())
        )
        statistics = (
            self._schedule_statistics(scheduling_input, assignments, entries)
            if has_schedule else LessonScheduleStatistics(0, 0, 0.0, 0, 0, 0)
        )
        source_data_revision = (
            result.source_data_revision if result is not None else state.source_data_revision
        )
        verification_count = sum(item.verified is False for item in entries)
        return LessonSchedule(
            schedule_revision=result.schedule_revision if result is not None else state.schedule_revision,
            source_data_revision=source_data_revision,
            is_schedule_stale=(
                state.is_schedule_stale
                or source_data_revision != state.source_data_revision
            ),
            has_schedule=has_schedule,
            placements=placements,
            unplaced_students=unplaced_students,
            statistics=statistics,
            entries=entries,
            verification=verification,
            automatic_count=sum(item.placement_type == LessonPlacementType.AUTOMATIC.value for item in entries),
            manual_count=sum(item.placement_type == LessonPlacementType.MANUAL.value for item in entries),
            manually_unplaced_count=sum(
                item.placement_type == LessonPlacementType.MANUALLY_UNPLACED.value for item in entries
            ),
            verified_count=sum(item.verified is True for item in entries),
            unverified_count=verification_count,
        )

    def clear_schedule(self) -> LessonSchedule:
        """Clear automatic assignments while retaining manual decisions and their state."""
        session_uuid = self._session_uuid()
        self.db.query(LessonPlacement).filter(
            LessonPlacement.placement_type == LessonPlacementType.AUTOMATIC.value
        ).delete()
        self.db.query(LessonUnplacedStudent).filter(
            LessonUnplacedStudent.placement_type == LessonPlacementType.AUTOMATIC.value
        ).delete()
        state = self.get_state()
        state.schedule_revision += 1
        state.modified_at = _now()
        assignments = self._refresh_schedule_snapshots(self.get_scheduling_input())
        if assignments or self.db.query(LessonUnplacedStudent.student_id).first() is not None:
            scheduling_input, _, _, entries, _ = self._schedule_review()
            stats = self._schedule_statistics(scheduling_input, assignments, entries)
            result = LessonOptimizationResult(
                source_data_revision=(
                    self.db.get(LessonScheduleResult, session_uuid).source_data_revision
                    if self.db.get(LessonScheduleResult, session_uuid) is not None
                    else state.source_data_revision
                ),
                placements=(),
                unplaced_students=(),
                statistics=stats,
            )
            self._save_schedule_result(result, state.schedule_revision, session_uuid)
        else:
            self.db.query(LessonScheduleResult).filter(
                LessonScheduleResult.session_uuid == session_uuid
            ).delete()
            state.is_schedule_stale = False
        self.db.flush()
        return self.get_schedule()

    def set_tag_state(self, lesson_student_id: str, tag_id: str, state: LessonTagState) -> None:
        """Persist a student/tag state; Neutral is represented by the absence of a row."""
        self.get_student(lesson_student_id)
        self._require(LessonTag, tag_id, "TAG_NOT_FOUND", "The tag no longer exists.")
        try:
            state = LessonTagState(state)
        except ValueError as error:
            raise LessonValidationError("INVALID_TAG_STATE", "Choose a valid tag state.") from error
        row = self.db.get(LessonStudentTagState, (lesson_student_id, tag_id))
        if state is LessonTagState.NEUTRAL:
            if row is not None:
                self.db.delete(row)
        elif row is None:
            self.db.add(LessonStudentTagState(student_id=lesson_student_id, tag_id=tag_id, state=state.value))
        else:
            row.state = state.value
        self.db.flush()
        self.record_source_change()

    def _validated_tag_states(self, states: Mapping[str, LessonTagState | str]) -> dict[str, LessonTagState]:
        validated: dict[str, LessonTagState] = {}
        for tag_id, raw in states.items():
            self._require(LessonTag, tag_id, "TAG_NOT_FOUND", "The tag no longer exists.")
            try:
                validated[tag_id] = LessonTagState(raw)
            except ValueError as error:
                raise LessonValidationError("INVALID_TAG_STATE", "Choose a valid tag state.") from error
        return validated

    def _replace_student_tag_states(self, lesson_student_id: str, states: Mapping[str, LessonTagState]) -> None:
        """Store only non-neutral states; Neutral is the absence of a row."""
        self.db.query(LessonStudentTagState).filter(LessonStudentTagState.student_id == lesson_student_id).delete()
        for tag_id, state in states.items():
            if state is not LessonTagState.NEUTRAL:
                self.db.add(LessonStudentTagState(student_id=lesson_student_id, tag_id=tag_id, state=state.value))
        self.db.flush()

    def set_student_tag_states(
        self, lesson_student_id: str, states: Mapping[str, LessonTagState | str]
    ) -> LessonStudent:
        """Replace a student's whole tag-state set as a single source-data change."""
        student = self.get_student(lesson_student_id)
        self._replace_student_tag_states(student.id, self._validated_tag_states(states))
        self.record_source_change()
        return student

    def student_availability_counts(self) -> dict[str, int]:
        return dict(
            self.db.query(
                LessonStudentAvailabilityWindow.lesson_student_id, func.count(LessonStudentAvailabilityWindow.id)
            )
            .group_by(LessonStudentAvailabilityWindow.lesson_student_id)
            .all()
        )

    def replace_students(
        self, students: Iterable[tuple[StudentInput, Sequence[AvailabilityWindow]]]
    ) -> LessonStudentReplaceResult:
        """Full replacement of students as one source-data change (caller owns the transaction).

        Existing placement decisions are retained and the resulting schedule becomes stale.
        Student availability and tag state are replaced; instructors and the tag registry remain.
        """
        prepared: list[tuple[tuple[str, str, str], list[AvailabilityWindow]]] = []
        seen: set[str] = set()
        for values, windows in students:
            # Existing students are about to be replaced, so only uniqueness within the batch matters.
            cleaned = self._clean_student_values(
                values.student_id, values.name, values.email, check_unique=False
            )
            if cleaned[0].casefold() in seen:
                raise LessonValidationError("DUPLICATE_STUDENT_ID", "A student with that Student ID already exists.")
            seen.add(cleaned[0].casefold())
            prepared.append((cleaned, self._validated_windows(windows)))
        students_removed = self.db.query(func.count(LessonStudent.id)).scalar() or 0
        placements_removed = 0
        placements_preserved = self.db.query(func.count(LessonPlacement.id)).scalar() or 0
        self.db.query(LessonStudentTagState).delete()
        self.db.query(LessonStudentAvailabilityWindow).delete()
        self.db.query(LessonStudent).delete()
        self.db.flush()
        for (student_id, name, email), windows in prepared:
            student = LessonStudent(student_id=student_id, name=name, email=email)
            self.db.add(student)
            self.db.flush()
            self._replace_student_availability(student.id, windows)
        self.record_source_change()
        return LessonStudentReplaceResult(
            students_removed=students_removed,
            placements_removed=placements_removed,
            placements_preserved=placements_preserved,
            students_created=len(prepared),
        )

    def get_tag_state(self, lesson_student_id: str, tag_id: str) -> LessonTagState:
        row = self.db.get(LessonStudentTagState, (lesson_student_id, tag_id))
        return LessonTagState(row.state) if row is not None else LessonTagState.NEUTRAL

    # --- placements ---------------------------------------------------

    def set_placement(
        self,
        lesson_student_id: str,
        placement_type: LessonPlacementType,
        lesson_instructor_id: str | None = None,
        start_minute: int | None = None,
        day: str = "Monday",
    ) -> LessonPlacement:
        """Create or replace the student's single placement decision."""
        self.get_student(lesson_student_id)
        try:
            placement_type = LessonPlacementType(placement_type)
        except ValueError as error:
            raise LessonValidationError("INVALID_PLACEMENT", "Choose a valid placement type.") from error
        if placement_type is LessonPlacementType.MANUALLY_UNPLACED:
            if lesson_instructor_id is not None or start_minute is not None:
                raise LessonValidationError(
                    "INVALID_PLACEMENT", "A manually unplaced decision has no instructor or start time."
                )
            end_minute = None
            day = "Monday"
        else:
            if lesson_instructor_id is None or start_minute is None:
                raise LessonValidationError("INVALID_PLACEMENT", "A placement requires an instructor and a start time.")
            if not _is_int(start_minute) or not 0 <= start_minute < 24 * 60:
                raise LessonValidationError("INVALID_PLACEMENT", "The start time is not a valid time of day.")
            normalized_day = normalize_day(day)
            if normalized_day is None:
                raise LessonValidationError("INVALID_PLACEMENT", "Choose a valid day of the week.")
            day = normalized_day
            instructor = self.get_instructor(lesson_instructor_id)
            end_minute = start_minute + instructor.lesson_length_minutes
            if end_minute > 24 * 60:
                raise LessonValidationError(
                    "INVALID_PLACEMENT", "The lesson must end before the end of the day."
                )
        placement = self.get_placement(lesson_student_id)
        if placement is None:
            placement = LessonPlacement(lesson_student_id=lesson_student_id)
            self.db.add(placement)
        placement.placement_type = placement_type.value
        placement.lesson_instructor_id = lesson_instructor_id
        placement.day = day
        placement.start_minute = start_minute
        placement.end_minute = end_minute
        placement.updated_at = _now()
        self.db.flush()
        self._record_schedule_edit()
        self._persist_schedule_after_edit()
        return placement

    def set_manual_placement(
        self, lesson_student_id: str, instructor_id: str, day: str, start_minute: int
    ) -> LessonPlacement:
        return self.set_placement(
            lesson_student_id,
            LessonPlacementType.MANUAL,
            instructor_id,
            start_minute,
            day,
        )

    def set_manually_unplaced(self, lesson_student_id: str) -> LessonPlacement:
        return self.set_placement(lesson_student_id, LessonPlacementType.MANUALLY_UNPLACED)

    def get_placement(self, lesson_student_id: str) -> LessonPlacement | None:
        return (
            self.db.query(LessonPlacement)
            .filter(LessonPlacement.lesson_student_id == lesson_student_id)
            .one_or_none()
        )

    def list_placements(self) -> list[LessonPlacement]:
        return self.db.query(LessonPlacement).order_by(LessonPlacement.created_at, LessonPlacement.id).all()

    def remove_placement(self, lesson_student_id: str) -> bool:
        placement = self.get_placement(lesson_student_id)
        if placement is None:
            raise LessonValidationError("PLACEMENT_NOT_FOUND", "The student has no placement decision to remove.")
        self.db.delete(placement)
        self.db.query(LessonUnplacedStudent).filter(
            LessonUnplacedStudent.student_id == lesson_student_id,
            LessonUnplacedStudent.placement_type == placement.placement_type,
        ).delete()
        self.db.flush()
        self._record_schedule_edit()
        self._persist_schedule_after_edit()
        return True

    def remove_manual_override(self, lesson_student_id: str) -> bool:
        placement = self.get_placement(lesson_student_id)
        if placement is None or placement.placement_type == LessonPlacementType.AUTOMATIC.value:
            raise LessonValidationError(
                "NO_MANUAL_OVERRIDE", "The student does not have a manual placement override to remove."
            )
        return self.remove_placement(lesson_student_id)

    def _persist_schedule_after_edit(self) -> None:
        scheduling_input = self.get_scheduling_input()
        existing_result = self.db.get(LessonScheduleResult, self._session_uuid())
        assignments = self._refresh_schedule_snapshots(scheduling_input)
        _, _, _, entries, _ = self._schedule_review(scheduling_input)
        if not self.list_placements() and not self.db.query(LessonUnplacedStudent.student_id).first():
            if existing_result is not None:
                self.db.delete(existing_result)
            return
        statistics = self._schedule_statistics(scheduling_input, assignments, entries)
        result = LessonOptimizationResult(
            source_data_revision=(
                existing_result.source_data_revision
                if existing_result is not None else scheduling_input.source_data_revision
            ),
            placements=tuple(
                LessonScheduledPlacement(
                    student_id=item.student_id,
                    instructor_id=item.instructor_id,
                    day=item.day,
                    start_minute=item.start_minute,
                    end_minute=item.end_minute,
                    preferred_tag_match_count=item.preferred_tag_match_count,
                    tentative_minutes=item.tentative_minutes,
                )
                for item in assignments
            ),
            unplaced_students=(),
            statistics=statistics,
        )
        state = self.get_state()
        self._save_schedule_result(result, state.schedule_revision, self._session_uuid())
        state.is_schedule_stale = (
            state.is_schedule_stale
            or result.source_data_revision != state.source_data_revision
        )

    def clear_all_placements(self) -> int:
        """Delete every placement decision and reset result/staleness state."""
        removed = self.db.query(LessonPlacement).delete()
        self.db.query(LessonSchedulePlacement).delete()
        self.db.query(LessonUnplacedStudent).delete()
        self.db.query(LessonScheduleResult).delete()
        state = self.get_state()
        state.schedule_revision = 0
        state.is_schedule_stale = False
        state.modified_at = _now()
        self.db.flush()
        return removed

    # --- helpers ------------------------------------------------------

    def _require(self, model, primary_key: str, code: str, message: str):
        row = self.db.get(model, primary_key)
        if row is None:
            raise LessonValidationError(code, message)
        return row


def _window(row) -> AvailabilityWindow:
    return AvailabilityWindow(
        day=row.day,
        start_minute=row.start_minute,
        end_minute=row.end_minute,
        status=row.status,
        source=row.source,
        source_row=row.source_row,
    )
