"""Lesson Scheduling persistence and module-state repository.

Pure persistence: CRUD for instructors, students, tags, tag states, availability
windows, and placement records, plus per-session revision/staleness state. It
contains no candidate generation, readiness analysis, or optimizer logic.
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
    LessonStudent,
    LessonStudentAvailabilityWindow,
    LessonStudentTagState,
    LessonTag,
    LessonTagState,
    lesson_tag_identity_key,
    normalize_lesson_tag_name,
)
from .availability import AvailabilityWindow, normalize_day, normalize_windows


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
    students_created: int


@dataclass(frozen=True)
class LessonReplaceResult:
    instructors_removed: int
    placements_removed: int
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

    def record_source_change(self) -> LessonModuleState:
        """Bump the source revision; flag an existing schedule stale without altering it."""
        state = self.get_state()
        state.source_data_revision += 1
        if self.db.query(LessonPlacement.id).first() is not None:
            state.is_schedule_stale = True
        state.modified_at = _now()
        self.db.flush()
        return state

    def _record_schedule_edit(self) -> None:
        state = self.get_state()
        state.schedule_revision += 1
        state.modified_at = _now()
        self.db.flush()

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
        """Delete an instructor and only their placements; returns the removed placement count."""
        instructor = self.get_instructor(instructor_id)
        removed = self.db.query(LessonPlacement).filter(
            LessonPlacement.lesson_instructor_id == instructor.id
        ).delete()
        self.db.query(LessonInstructorTag).filter(LessonInstructorTag.instructor_id == instructor.id).delete()
        self.db.query(LessonInstructorAvailabilityWindow).filter(
            LessonInstructorAvailabilityWindow.instructor_id == instructor.id
        ).delete()
        self.db.delete(instructor)
        self.db.flush()
        self.record_source_change()
        return removed

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
        """Full replacement of instructors as one source-data change (caller owns the transaction).

        Clears every placement, reuses tags by normalized name, creates missing tags, never
        deletes tags, and leaves students and their availability and tag states untouched.
        """
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
        placements_removed = self.clear_all_placements()
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
        require_name: bool = False,
        check_unique: bool = True,
    ) -> tuple[str, str, str]:
        cleaned_id = (student_id or "").strip()
        if not cleaned_id:
            raise LessonValidationError("INVALID_STUDENT_ID", "Student ID is required.")
        if require_name and not (name or "").strip():
            raise LessonValidationError("INVALID_STUDENT_NAME", "Name is required.")
        if check_unique:
            query = self.db.query(LessonStudent.id).filter(func.lower(LessonStudent.student_id) == cleaned_id.lower())
            if exclude_id is not None:
                query = query.filter(LessonStudent.id != exclude_id)
            if query.first() is not None:
                raise LessonValidationError("DUPLICATE_STUDENT_ID", "A student with that Student ID already exists.")
        return cleaned_id, (name or "").strip(), (email or "").strip()

    def create_student(
        self,
        student_id: str,
        name: str = "",
        email: str = "",
        availability: Sequence[AvailabilityWindow] = (),
        require_name: bool = False,
        tag_states: Mapping[str, LessonTagState | str] | None = None,
    ) -> LessonStudent:
        cleaned_id, cleaned_name, cleaned_email = self._clean_student_values(
            student_id, name, email, require_name=require_name
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
        require_name: bool = False,
        tag_states: Mapping[str, LessonTagState | str] | None = None,
    ) -> LessonStudent:
        """Update fields (and, when given, the full tag-state set) as one source-data change."""
        student = self.get_student(lesson_student_id)
        cleaned = self._clean_student_values(
            student_id, name, email, exclude_id=student.id, require_name=require_name
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
        self.db.query(LessonPlacement).filter(LessonPlacement.lesson_student_id == student.id).delete()
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

        Clears every placement, student availability, and student tag state. Instructors,
        instructor availability, and the tag registry are untouched.
        """
        prepared: list[tuple[tuple[str, str, str], list[AvailabilityWindow]]] = []
        seen: set[str] = set()
        for values, windows in students:
            # Existing students are about to be replaced, so only uniqueness within the batch matters.
            cleaned = self._clean_student_values(
                values.student_id, values.name, values.email, require_name=True, check_unique=False
            )
            if cleaned[0].casefold() in seen:
                raise LessonValidationError("DUPLICATE_STUDENT_ID", "A student with that Student ID already exists.")
            seen.add(cleaned[0].casefold())
            prepared.append((cleaned, self._validated_windows(windows)))
        students_removed = self.db.query(func.count(LessonStudent.id)).scalar() or 0
        placements_removed = self.clear_all_placements()
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
    ) -> LessonPlacement:
        """Create or replace the student's single placement decision (persistence only)."""
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
        else:
            if lesson_instructor_id is None or start_minute is None:
                raise LessonValidationError("INVALID_PLACEMENT", "A placement requires an instructor and a start time.")
            if not _is_int(start_minute) or not 0 <= start_minute < 24 * 60:
                raise LessonValidationError("INVALID_PLACEMENT", "The start time is not a valid time of day.")
            self.get_instructor(lesson_instructor_id)
        placement = self.get_placement(lesson_student_id)
        if placement is None:
            placement = LessonPlacement(lesson_student_id=lesson_student_id)
            self.db.add(placement)
        placement.placement_type = placement_type.value
        placement.lesson_instructor_id = lesson_instructor_id
        placement.start_minute = start_minute
        placement.updated_at = _now()
        self.db.flush()
        self._record_schedule_edit()
        return placement

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
        self.db.flush()
        self._record_schedule_edit()
        return True

    def clear_all_placements(self) -> int:
        """Delete every placement decision and reset result/staleness state."""
        removed = self.db.query(LessonPlacement).delete()
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
