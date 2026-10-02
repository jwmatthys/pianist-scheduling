"""Shared identity mapping and module revision lifecycle operations."""

from datetime import datetime
from uuid import uuid4

from sqlalchemy.orm import Session

from .. import models, module_models, schemas

ACCOMPANIST_MODULE_ID = "accompanists"
ACCOMPANIST_RESULT_CONTRACT = "accompanist.assignment-result"
ACCOMPANIST_RESULT_CONTRACT_VERSION = 2


def active_session_uuid(db: Session) -> str:
    sessions = db.query(models.SchedulingSession).all()
    if len(sessions) != 1:
        raise ValueError("The active database must contain exactly one Scheduling Session.")
    return sessions[0].session_uuid


def module_revision(db: Session, module_id: str = ACCOMPANIST_MODULE_ID) -> module_models.ModuleRevision:
    session_uuid = active_session_uuid(db)
    state = db.get(module_models.ModuleRevision, (session_uuid, module_id))
    if state is None:
        raise ValueError(f"The {module_id} module revision is unavailable.")
    return state


def bump_accompanist_revision(db: Session) -> int:
    state = module_revision(db)
    state.source_revision += 1
    state.modified_at = datetime.utcnow()
    if state.current_result_uuid:
        current = db.get(module_models.ModuleResult, state.current_result_uuid)
        if current is not None and current.state == "finalized":
            current.state = "superseded"
        state.current_result_uuid = None
    db.flush()
    return state.source_revision


def bump_jury_revision(db: Session) -> int:
    session_uuid = active_session_uuid(db)
    configuration = db.get(module_models.JuryConfiguration, session_uuid)
    if configuration is None:
        raise ValueError("The Jury configuration is unavailable.")
    configuration.input_revision += 1
    configuration.modified_at = datetime.utcnow()
    db.flush()
    return configuration.input_revision


def _refresh_student_profile(db: Session, profile: module_models.AccompanistStudentProfile) -> None:
    lesson_rows = (
        db.query(models.Lesson.student)
        .join(
            module_models.AccompanistLessonIdentity,
            module_models.AccompanistLessonIdentity.lesson_id == models.Lesson.id,
        )
        .filter(module_models.AccompanistLessonIdentity.student_person_uuid == profile.person_uuid)
        .all()
    )
    names = {
        module_models.normalize_student_name(row[0])
        for row in lesson_rows
        if module_models.normalize_student_name(row[0])
    }
    profile.seen_names_json = module_models.normalized_names_json(names)
    profile.identity_conflict = len(names) > 1


def refresh_student_identity_state(db: Session, person_uuid: str) -> None:
    profile = db.get(module_models.AccompanistStudentProfile, person_uuid)
    if profile is not None:
        _refresh_student_profile(db, profile)


def synchronize_lesson_identity(
    db: Session,
    lesson: models.Lesson,
    *,
    identity_fields_changed: bool = False,
) -> str:
    """Resolve a lesson to its stable UUID without matching by display name."""
    mapping = db.get(module_models.AccompanistLessonIdentity, lesson.id)
    if mapping is not None and not identity_fields_changed:
        profile = db.get(module_models.AccompanistStudentProfile, mapping.student_person_uuid)
        if profile is not None:
            _refresh_student_profile(db, profile)
            return profile.person_uuid

    old_profile = (
        db.get(module_models.AccompanistStudentProfile, mapping.student_person_uuid)
        if mapping is not None else None
    )
    session_uuid = active_session_uuid(db)
    normalized_id = module_models.normalize_student_id(lesson.student_id)
    profile = None
    if normalized_id:
        profile = (
            db.query(module_models.AccompanistStudentProfile)
            .filter(
                module_models.AccompanistStudentProfile.session_uuid == session_uuid,
                module_models.AccompanistStudentProfile.normalized_student_id == normalized_id,
            )
            .one_or_none()
        )
    elif old_profile is not None and old_profile.normalized_student_id is None:
        profile = old_profile

    if profile is None:
        person_uuid = str(uuid4())
        identity = module_models.PersonIdentity(
            person_uuid=person_uuid,
            display_name=lesson.student or "",
        )
        profile = module_models.AccompanistStudentProfile(
            person_uuid=person_uuid,
            session_uuid=session_uuid,
            student_id=normalized_id or None,
            normalized_student_id=normalized_id or None,
            seen_names_json="[]",
            identity_conflict=False,
        )
        db.add_all([identity, profile])
        db.flush()

    if mapping is None:
        mapping = module_models.AccompanistLessonIdentity(
            lesson_id=lesson.id,
            lesson_uuid=str(uuid4()),
            student_person_uuid=profile.person_uuid,
        )
        db.add(mapping)
    else:
        mapping.student_person_uuid = profile.person_uuid

    identity = db.get(module_models.PersonIdentity, profile.person_uuid)
    if identity is not None and not identity.display_name:
        identity.display_name = lesson.student or ""
    if normalized_id:
        profile.student_id = normalized_id
    _refresh_student_profile(db, profile)
    if old_profile is not None and old_profile.person_uuid != profile.person_uuid:
        _refresh_student_profile(db, old_profile)
    db.flush()
    return profile.person_uuid


def ensure_pianist_identity(db: Session, pianist: models.Pianist) -> str:
    mapping = db.get(module_models.AccompanistPianistIdentity, pianist.id)
    if mapping is None:
        person_uuid = str(uuid4())
        db.add(module_models.PersonIdentity(
            person_uuid=person_uuid,
            display_name=pianist.name or "",
        ))
        mapping = module_models.AccompanistPianistIdentity(
            pianist_id=pianist.id,
            person_uuid=person_uuid,
        )
        db.add(mapping)
        db.flush()
        return person_uuid

    identity = db.get(module_models.PersonIdentity, mapping.person_uuid)
    if identity is None:
        raise ValueError("The pianist identity mapping is invalid.")
    identity.display_name = pianist.name or ""
    return identity.person_uuid


def remove_pianist_identity_mapping(db: Session, pianist_id: int) -> None:
    mapping = db.get(module_models.AccompanistPianistIdentity, pianist_id)
    if mapping is not None:
        db.delete(mapping)


def get_lesson_jury_required(db: Session, lesson_id: int) -> bool:
    row = db.get(module_models.AccompanistLessonJuryRequirement, lesson_id)
    return bool(row and row.jury_required)


def set_lesson_jury_required(db: Session, lesson_id: int, value: bool) -> bool:
    row = db.get(module_models.AccompanistLessonJuryRequirement, lesson_id)
    if row is None:
        row = module_models.AccompanistLessonJuryRequirement(
            lesson_id=lesson_id,
            jury_required=value,
        )
        db.add(row)
        db.flush()
        return True
    changed = row.jury_required != value
    row.jury_required = value
    return changed


def lesson_response(db: Session, lesson: models.Lesson) -> schemas.LessonOut:
    response = schemas.LessonOut.model_validate(lesson)
    return response.model_copy(update={"jury_required": get_lesson_jury_required(db, lesson.id)})
