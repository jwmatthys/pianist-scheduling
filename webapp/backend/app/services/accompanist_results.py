"""Publish and resolve the typed finalized Accompanist result contract."""

import hashlib
import json
from dataclasses import dataclass
from datetime import datetime
from uuid import UUID, uuid4

from sqlalchemy import func
from sqlalchemy.orm import Session

from .. import models, module_models
from ..jury_schemas import (
    AccompanistAssignmentPayload,
    AccompanistAssignmentPayloadV1,
    FinalizedLessonEntry,
    FinalizedPianist,
    HistoricalResultEnvelopeV1,
    ResultEnvelope,
)
from .module_lifecycle import (
    ACCOMPANIST_MODULE_ID,
    ACCOMPANIST_RESULT_CONTRACT,
    ACCOMPANIST_RESULT_CONTRACT_VERSION,
    active_session_uuid,
    get_lesson_jury_required,
    module_revision,
)

PAYLOAD_SCHEMA_VERSION = 2


@dataclass(frozen=True)
class AccompanistSourceIdentitySnapshot:
    session_uuid: UUID
    source_revision: int
    current_result_uuid: UUID | None
    lesson_identities: tuple[tuple[UUID, UUID], ...]
    pianist_person_uuids: tuple[UUID, ...]


class ResultPublicationError(ValueError):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


def _build_payload(db: Session) -> AccompanistAssignmentPayload:
    lessons = db.query(models.Lesson).order_by(models.Lesson.id).all()
    entries: list[FinalizedLessonEntry] = []
    seen_lesson_uuids: set[str] = set()

    for lesson in lessons:
        mapping = db.get(module_models.AccompanistLessonIdentity, lesson.id)
        if mapping is None or not mapping.lesson_uuid or not mapping.student_person_uuid:
            raise ResultPublicationError(
                "UNRESOLVED_LESSON_IDENTITY",
                f"Lesson {lesson.id} has no valid shared lesson/student identity mapping.",
            )
        if mapping.lesson_uuid in seen_lesson_uuids:
            raise ResultPublicationError(
                "DUPLICATE_LESSON_IDENTITY",
                "Multiple Accompanist lessons resolve to the same Lesson UUID.",
            )
        seen_lesson_uuids.add(mapping.lesson_uuid)

        profile = db.get(module_models.AccompanistStudentProfile, mapping.student_person_uuid)
        student = db.get(module_models.PersonIdentity, mapping.student_person_uuid)
        if profile is None or student is None or profile.session_uuid != active_session_uuid(db):
            raise ResultPublicationError(
                "INVALID_STUDENT_REFERENCE",
                f"Lesson {lesson.id} references an invalid student identity.",
            )
        if profile.identity_conflict:
            raise ResultPublicationError(
                "STUDENT_IDENTITY_CONFLICT",
                f"Lesson {lesson.id} belongs to a Student ID group with conflicting identity evidence.",
            )

        assigned_pianist = None
        if lesson.assigned_pianist_id is not None:
            pianist = db.get(models.Pianist, lesson.assigned_pianist_id)
            pianist_mapping = db.get(
                module_models.AccompanistPianistIdentity,
                lesson.assigned_pianist_id,
            )
            if pianist is None or pianist_mapping is None:
                raise ResultPublicationError(
                    "INVALID_PIANIST_REFERENCE",
                    f"Lesson {lesson.id} references a pianist without a valid shared identity.",
                )
            pianist_identity = db.get(module_models.PersonIdentity, pianist_mapping.person_uuid)
            if pianist_identity is None:
                raise ResultPublicationError(
                    "INVALID_PIANIST_IDENTITY",
                    f"Lesson {lesson.id} references an unavailable pianist identity.",
                )
            assigned_pianist = FinalizedPianist(
                person_uuid=pianist_mapping.person_uuid,
                display_name=pianist_identity.display_name or pianist.name,
            )

        entries.append(FinalizedLessonEntry(
            source_lesson_uuid=mapping.lesson_uuid,
            student_person_uuid=mapping.student_person_uuid,
            student_display_name=lesson.student,
            instrument=lesson.instrument,
            teacher=lesson.teacher,
            pianist_required=lesson.need_pianist,
            assigned_pianist=assigned_pianist,
            jury_required=get_lesson_jury_required(db, lesson.id),
        ))

    return AccompanistAssignmentPayload(entries=entries)


def _envelope(
    db: Session,
    result: module_models.ModuleResult,
) -> ResultEnvelope | HistoricalResultEnvelopeV1:
    common = {
        "result_uuid": result.result_uuid,
        "session_uuid": result.session_uuid,
        "module_id": result.module_id,
        "contract_id": result.contract_id,
        "contract_version": result.contract_version,
        "source_revision": result.source_revision,
        "result_version": result.result_version,
        "state": result.state,
        "created_at": result.created_at,
        "finalized_at": result.finalized_at,
        "payload_schema_version": result.payload_schema_version,
    }
    if result.contract_version == 1 and result.payload_schema_version == 1:
        return HistoricalResultEnvelopeV1(
            **common,
            payload=AccompanistAssignmentPayloadV1.model_validate_json(result.payload_json),
        )
    return ResultEnvelope(
        **common,
        payload=AccompanistAssignmentPayload.model_validate_json(result.payload_json),
    )


def finalize_accompanist_result(
    db: Session,
    expected_source_revision: int,
) -> ResultEnvelope:
    state = module_revision(db)
    if state.source_revision != expected_source_revision:
        raise ResultPublicationError(
            "SOURCE_REVISION_CHANGED",
            "Accompanist data changed after it was reviewed. Refresh and finalize the current revision.",
        )

    payload = _build_payload(db)
    session_uuid = active_session_uuid(db)
    if state.current_result_uuid:
        previous = db.get(module_models.ModuleResult, state.current_result_uuid)
        if previous is not None and previous.state == "finalized":
            previous.state = "superseded"

    latest_version = db.query(func.max(module_models.ModuleResult.result_version)).filter(
        module_models.ModuleResult.session_uuid == session_uuid,
        module_models.ModuleResult.module_id == ACCOMPANIST_MODULE_ID,
    ).scalar()
    result_uuid = str(uuid4())
    now = datetime.utcnow()
    payload_json = json.dumps(
        payload.model_dump(mode="json"),
        ensure_ascii=True,
        separators=(",", ":"),
        sort_keys=True,
    )
    result = module_models.ModuleResult(
        result_uuid=result_uuid,
        session_uuid=session_uuid,
        module_id=ACCOMPANIST_MODULE_ID,
        contract_id=ACCOMPANIST_RESULT_CONTRACT,
        contract_version=ACCOMPANIST_RESULT_CONTRACT_VERSION,
        payload_schema_version=PAYLOAD_SCHEMA_VERSION,
        result_version=(latest_version or 0) + 1,
        source_revision=state.source_revision,
        state="finalized",
        payload_json=payload_json,
        payload_sha256=hashlib.sha256(payload_json.encode("utf-8")).hexdigest(),
        created_at=now,
        finalized_at=now,
    )
    db.add(result)
    state.current_result_uuid = result_uuid
    db.flush()
    return _envelope(db, result)


def get_current_accompanist_result(db: Session) -> ResultEnvelope | None:
    state = module_revision(db)
    if not state.current_result_uuid:
        return None
    result = get_latest_accompanist_result(db)
    if (
        result is None
        or result.state != "finalized"
        or result.source_revision != state.source_revision
    ):
        return None
    return result


def get_latest_accompanist_result(db: Session) -> ResultEnvelope | None:
    session_uuid = active_session_uuid(db)
    result = db.query(module_models.ModuleResult).filter(
        module_models.ModuleResult.session_uuid == session_uuid,
        module_models.ModuleResult.module_id == ACCOMPANIST_MODULE_ID,
        module_models.ModuleResult.contract_id == ACCOMPANIST_RESULT_CONTRACT,
        module_models.ModuleResult.contract_version == ACCOMPANIST_RESULT_CONTRACT_VERSION,
        module_models.ModuleResult.payload_schema_version == PAYLOAD_SCHEMA_VERSION,
        module_models.ModuleResult.state.in_(["finalized", "superseded"]),
    ).order_by(module_models.ModuleResult.result_version.desc()).first()
    if (
        result is None
    ):
        return None
    return _envelope(db, result)


def get_result_by_uuid(
    db: Session,
    result_uuid: str,
) -> ResultEnvelope | HistoricalResultEnvelopeV1 | None:
    result = db.get(module_models.ModuleResult, result_uuid)
    if result is None or result.module_id != ACCOMPANIST_MODULE_ID:
        return None
    if result.contract_id != ACCOMPANIST_RESULT_CONTRACT:
        return None
    return _envelope(db, result)


def get_current_source_identity_snapshot(db: Session) -> AccompanistSourceIdentitySnapshot:
    """Project current Accompanist identity references for dependent-module cleanup."""
    session_uuid = active_session_uuid(db)
    state = module_revision(db)
    lessons = db.query(
        module_models.AccompanistLessonIdentity.lesson_uuid,
        module_models.AccompanistLessonIdentity.student_person_uuid,
    ).join(
        models.Lesson,
        models.Lesson.id == module_models.AccompanistLessonIdentity.lesson_id,
    ).join(
        module_models.AccompanistStudentProfile,
        module_models.AccompanistStudentProfile.person_uuid == module_models.AccompanistLessonIdentity.student_person_uuid,
    ).join(
        module_models.PersonIdentity,
        module_models.PersonIdentity.person_uuid == module_models.AccompanistLessonIdentity.student_person_uuid,
    ).filter(
        module_models.AccompanistStudentProfile.session_uuid == session_uuid,
    ).order_by(module_models.AccompanistLessonIdentity.lesson_uuid).all()
    pianists = db.query(
        module_models.AccompanistPianistIdentity.person_uuid,
    ).join(
        models.Pianist,
        models.Pianist.id == module_models.AccompanistPianistIdentity.pianist_id,
    ).join(
        module_models.PersonIdentity,
        module_models.PersonIdentity.person_uuid == module_models.AccompanistPianistIdentity.person_uuid,
    ).order_by(module_models.AccompanistPianistIdentity.person_uuid).all()
    return AccompanistSourceIdentitySnapshot(
        session_uuid=UUID(session_uuid),
        source_revision=state.source_revision,
        current_result_uuid=UUID(state.current_result_uuid) if state.current_result_uuid else None,
        lesson_identities=tuple((UUID(lesson_uuid), UUID(person_uuid)) for lesson_uuid, person_uuid in lessons),
        pianist_person_uuids=tuple(UUID(person_uuid) for (person_uuid,) in pianists),
    )