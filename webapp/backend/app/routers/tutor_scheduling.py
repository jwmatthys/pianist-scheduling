"""Tutor Scheduling module API: Tutor and Student Rosters, Tutor Tags, and the roster imports.

Handlers stay thin; domain rules and revision/staleness tracking live in TutorRepository.
Each request commits as one transaction.
"""

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from pydantic import ValidationError
from pydantic import BaseModel
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from .. import tutor_schemas as schemas
from ..database import get_db
from ..tutor_models import TutorTutor, TutorPlacement, TutorPlacementState, TutorStudent
from ..services import tutor_tutor_importer as importer
from ..services import tutor_student_importer as student_importer
from ..services.availability import AvailabilityWindow, normalize_day
from ..services.tutor_repository import (
    TutorInput,
    TutorRepository,
    TutorValidationError,
    StudentInput,
)
from ..services.tutor_scheduling_models import TutorSchedule

router = APIRouter(prefix="/api/tutor", tags=["tutor scheduling"])

_NOT_FOUND_CODES = {"TAG_NOT_FOUND", "TUTOR_NOT_FOUND", "STUDENT_NOT_FOUND"}
_CONFLICT_CODES = {
    "DUPLICATE_TAG",
    "STUDENT_ID_NAME_CONFLICT",
    "NOT_READY",
    "MANUAL_PLACEMENTS_PRESENT",
}


def _http_error(error: TutorValidationError) -> HTTPException:
    if error.code in _NOT_FOUND_CODES:
        status = 404
    elif error.code in _CONFLICT_CODES:
        status = 409
    else:
        status = 422
    return HTTPException(status_code=status, detail={"code": error.code, "message": str(error)})


def _commit(db: Session) -> None:
    try:
        db.commit()
    except IntegrityError as error:
        db.rollback()
        raise HTTPException(
            status_code=409,
            detail={"code": "TUTOR_CONFLICT", "message": "The change conflicts with existing Tutor Scheduling data."},
        ) from error


def _tutor_out(repo: TutorRepository, tutor: TutorTutor) -> schemas.TutorTutorOut:
    minutes = repo.tutor_availability_minutes().get(tutor.id, (0, 0))
    return schemas.TutorTutorOut(
        id=tutor.id,
        name=tutor.name,
        email=tutor.email,
        room=tutor.room,
        lesson_length_minutes=tutor.lesson_length_minutes,
        capacity=tutor.capacity,
        tags=[schemas.TutorTagRef(id=tag.id, name=tag.name) for tag in tutor.tags],
        availability_window_count=repo.tutor_availability_counts().get(tutor.id, 0),
        available_minutes=minutes[0],
        tentative_minutes=minutes[1],
        created_at=tutor.created_at,
        updated_at=tutor.updated_at,
    )


def _fields(payload: schemas.TutorTutorFields) -> dict:
    return {
        "name": payload.name,
        "email": payload.email,
        "room": payload.room,
        "lesson_length_minutes": payload.lesson_length_minutes,
        "capacity": payload.capacity,
    }


def _tag_out(summary) -> schemas.TutorTagOut:
    return schemas.TutorTagOut(
        id=summary.tag.id,
        name=summary.tag.name,
        tutor_count=summary.tutor_count,
        student_count=summary.student_count,
        is_unused=summary.is_unused,
    )


def _tag_summary(repo: TutorRepository, tag_id: str) -> schemas.TutorTagOut:
    for summary in repo.tag_summaries():
        if summary.tag.id == tag_id:
            return _tag_out(summary)
    raise HTTPException(404, detail={"code": "TAG_NOT_FOUND", "message": "The tag no longer exists."})


@router.get("/state", response_model=schemas.TutorModuleStateOut)
def read_state(db: Session = Depends(get_db)):
    state = TutorRepository(db).get_state()
    db.commit()
    return schemas.TutorModuleStateOut(
        schedule_revision=state.schedule_revision,
        source_data_revision=state.source_data_revision,
        is_schedule_stale=state.is_schedule_stale,
    )


@router.get("/readiness", response_model=schemas.TutorReadinessOut)
def get_readiness(db: Session = Depends(get_db)):
    report = TutorRepository(db).get_readiness_report()
    db.commit()
    return schemas.TutorReadinessOut(
        source_data_revision=report.source_data_revision,
        is_ready=report.is_ready,
        student_count=report.student_count,
        tutor_count=report.tutor_count,
        feasible_candidate_count=report.feasible_candidate_count,
        total_candidate_count=report.total_candidate_count,
        students_with_candidates=report.students_with_candidates,
        tutors_with_schedulable_tutors=report.tutors_with_schedulable_tutors,
        warnings=[
            schemas.TutorReadinessWarningOut(
                code=warning.code.value,
                severity=warning.severity.value,
                category=warning.category.value,
                subjects=[
                    schemas.TutorReadinessSubjectOut(
                        kind=subject.kind.value,
                        id=subject.id,
                        label=subject.label,
                    )
                    for subject in warning.subjects
                ],
            )
            for warning in report.warnings
        ],
    )


def _placement_state(entry) -> TutorPlacementState:
    """Derive the state from actual placement data: an automatically unplaced row has no tutor time."""
    if entry.placement_type == "manuallyUnplaced":
        return TutorPlacementState.MANUALLY_UNPLACED
    if entry.day is None or entry.start_minute is None:
        return TutorPlacementState.UNASSIGNED
    return TutorPlacementState.ASSIGNED


def _schedule_out(schedule: TutorSchedule) -> schemas.TutorScheduleOut:
    stats = schedule.statistics
    return schemas.TutorScheduleOut(
        schedule_revision=schedule.schedule_revision,
        source_data_revision=schedule.source_data_revision,
        is_schedule_stale=schedule.is_schedule_stale,
        has_schedule=schedule.has_schedule,
        students_scheduled=stats.students_scheduled,
        students_unscheduled=stats.students_unscheduled,
        tutor_utilization=stats.tutor_utilization,
        preferred_tag_matches=stats.preferred_tag_matches,
        tentative_minutes=stats.tentative_minutes,
        gap_cost_minutes=stats.gap_cost_minutes,
        automatic_count=schedule.automatic_count,
        manual_count=schedule.manual_count,
        manually_unplaced_count=schedule.manually_unplaced_count,
        verified_count=schedule.verified_count,
        unverified_count=schedule.unverified_count,
        entries=[
            schemas.TutorScheduleEntryOut(
                student_id=item.student_id,
                student_name=item.student_name,
                placement_type=item.placement_type,
                placement_state=_placement_state(item).value,
                tutor_id=item.tutor_id,
                tutor_name=item.tutor_name,
                day=item.day,
                start_minute=item.start_minute,
                end_minute=item.end_minute,
                preferred_tag_match_count=item.preferred_tag_match_count,
                tentative_minutes=item.tentative_minutes,
                verified=item.verified,
                warnings=[
                    schemas.TutorPlacementWarningOut(code=warning.code.value, message=warning.message)
                    for warning in item.warnings
                ],
                unplaced_reason=item.unplaced_reason,
            )
            for item in schedule.entries
        ],
        verification=[
            schemas.TutorPlacementVerificationOut(
                student_id=item.student_id,
                student_name=item.student_name,
                tutor_id=item.tutor_id,
                tutor_name=item.tutor_name,
                day=item.day,
                start_minute=item.start_minute,
                warnings=[
                    schemas.TutorPlacementWarningOut(code=warning.code.value, message=warning.message)
                    for warning in item.warnings
                ],
            )
            for item in schedule.verification
        ],
    )


@router.get("/schedule", response_model=schemas.TutorScheduleOut)
def get_schedule(db: Session = Depends(get_db)):
    result = TutorRepository(db).get_schedule()
    db.commit()
    return _schedule_out(result)


@router.post("/schedule/generate", response_model=schemas.TutorScheduleOut)
def generate_schedule(db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    try:
        result = repo.generate_schedule()
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _schedule_out(result)


@router.post("/schedule/clear", response_model=schemas.TutorScheduleOut)
def clear_schedule(db: Session = Depends(get_db)):
    result = TutorRepository(db).clear_schedule()
    _commit(db)
    return _schedule_out(result)


@router.get("/schedule/verification", response_model=list[schemas.TutorPlacementVerificationOut])
def get_schedule_verification(db: Session = Depends(get_db)):
    return _schedule_out(TutorRepository(db).get_schedule()).verification


def _save_manual_placement(
    payload: schemas.TutorManualPlacementIn, db: Session
) -> schemas.TutorScheduleOut:
    try:
        TutorRepository(db).set_manual_placement(
            payload.student_id, payload.tutor_id, payload.day, payload.start_minute
        )
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _schedule_out(TutorRepository(db).get_schedule())


@router.post("/schedule/placements", response_model=schemas.TutorScheduleOut)
def create_manual_placement(payload: schemas.TutorManualPlacementIn, db: Session = Depends(get_db)):
    return _save_manual_placement(payload, db)


@router.put("/schedule/placements/{student_id}", response_model=schemas.TutorScheduleOut)
def move_manual_placement(
    student_id: str, payload: schemas.TutorManualPlacementIn, db: Session = Depends(get_db)
):
    if payload.student_id != student_id:
        raise HTTPException(
            status_code=422,
            detail={"code": "INVALID_PLACEMENT", "message": "Student identifier does not match the placement URL."},
        )
    return _save_manual_placement(payload, db)


@router.post("/schedule/students/{student_id}/manually-unplace", response_model=schemas.TutorScheduleOut)
def manually_unplace_student(student_id: str, db: Session = Depends(get_db)):
    try:
        TutorRepository(db).set_manually_unplaced(student_id)
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _schedule_out(TutorRepository(db).get_schedule())


@router.delete("/schedule/placements/{student_id}", response_model=schemas.TutorScheduleOut)
def remove_placement_decision(student_id: str, db: Session = Depends(get_db)):
    """Return a student to Unassigned: remove any automatic, manual, or Manually Unplaced decision."""
    try:
        TutorRepository(db).remove_placement(student_id)
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _schedule_out(TutorRepository(db).get_schedule())


@router.delete("/schedule/placements/{student_id}/override", response_model=schemas.TutorScheduleOut)
def remove_manual_placement_override(student_id: str, db: Session = Depends(get_db)):
    try:
        TutorRepository(db).remove_manual_override(student_id)
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _schedule_out(TutorRepository(db).get_schedule())


# --- Tutors ------------------------------------------------------------


@router.get("/tutors", response_model=list[schemas.TutorTutorOut])
def list_tutors(db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    return [_tutor_out(repo, item) for item in repo.list_tutors()]


@router.post("/tutors", response_model=schemas.TutorTutorOut, status_code=201)
def create_tutor(payload: schemas.TutorTutorIn, db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    try:
        tutor = repo.create_tutor(**_fields(payload), tag_ids=payload.tag_ids or [])
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _tutor_out(repo, tutor)


@router.put("/tutors/{tutor_id}", response_model=schemas.TutorTutorOut)
def update_tutor(tutor_id: str, payload: schemas.TutorTutorIn, db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    try:
        tutor = repo.update_tutor(tutor_id, **_fields(payload), tag_ids=payload.tag_ids)
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _tutor_out(repo, tutor)


@router.put("/tutors/{tutor_id}/tags", response_model=schemas.TutorTutorOut)
def set_tutor_tags(tutor_id: str, payload: schemas.TutorTutorTagsIn, db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    try:
        tutor = repo.set_tutor_tags(tutor_id, payload.tag_ids)
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _tutor_out(repo, tutor)


@router.delete("/tutors/{tutor_id}", response_model=schemas.TutorTutorDeleteOut)
def delete_tutor(tutor_id: str, db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    try:
        preserved = sum(
            placement.tutor_tutor_id == tutor_id
            for placement in repo.list_placements()
        )
        removed = repo.delete_tutor(tutor_id)
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return schemas.TutorTutorDeleteOut(
        removed_placement_count=removed,
        placements_preserved=preserved,
    )


def _availability_response(repo: TutorRepository, tutor_id: str) -> schemas.TutorTutorAvailabilityIO:
    return schemas.TutorTutorAvailabilityIO(
        windows=[
            schemas.TutorAvailabilityWindowIO(
                day=window.day, start_minute=window.start_minute, end_minute=window.end_minute, status=window.status
            )
            for window in repo.tutor_availability(tutor_id)
        ]
    )


@router.get("/tutors/{tutor_id}/availability", response_model=schemas.TutorTutorAvailabilityIO)
def get_tutor_availability(tutor_id: str, db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    try:
        repo.get_tutor(tutor_id)
    except TutorValidationError as error:
        raise _http_error(error) from error
    return _availability_response(repo, tutor_id)


@router.put("/tutors/{tutor_id}/availability", response_model=schemas.TutorTutorAvailabilityIO)
def set_tutor_availability(
    tutor_id: str, payload: schemas.TutorTutorAvailabilityIO, db: Session = Depends(get_db)
):
    repo = TutorRepository(db)
    windows = [
        AvailabilityWindow(
            day=normalize_day(item.day) or item.day,
            start_minute=item.start_minute,
            end_minute=item.end_minute,
            status=item.status,  # type: ignore[arg-type]
        )
        for item in payload.windows
    ]
    try:
        repo.set_tutor_availability(tutor_id, windows)
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _availability_response(repo, tutor_id)


# --- Students ---------------------------------------------------------------


def _student_out(student: TutorStudent, states: dict, availability_count: int) -> schemas.TutorStudentOut:
    return schemas.TutorStudentOut(
        id=student.id,
        student_id=student.student_id,
        name=student.name,
        email=student.email,
        availability_window_count=availability_count,
        tag_states={tag_id: state.value for tag_id, state in states.items()},
        created_at=student.created_at,
        updated_at=student.updated_at,
    )


def _student_response(repo: TutorRepository, student: TutorStudent) -> schemas.TutorStudentOut:
    return _student_out(
        student,
        repo.student_tag_states().get(student.id, {}),
        repo.student_availability_counts().get(student.id, 0),
    )


@router.get("/students", response_model=list[schemas.TutorStudentOut])
def list_students(db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    states = repo.student_tag_states()
    counts = repo.student_availability_counts()
    return [_student_out(item, states.get(item.id, {}), counts.get(item.id, 0)) for item in repo.list_students()]


@router.post("/students", response_model=schemas.TutorStudentOut, status_code=201)
def create_student(payload: schemas.TutorStudentIn, db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    try:
        student = repo.create_student(
            payload.student_id,
            payload.name,
            payload.email,
            tag_states=payload.tag_states,
        )
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _student_response(repo, student)


@router.put("/students/{student_pk}", response_model=schemas.TutorStudentOut)
def update_student(student_pk: str, payload: schemas.TutorStudentIn, db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    try:
        student = repo.update_student(
            student_pk,
            payload.student_id,
            payload.name,
            payload.email,
            tag_states=payload.tag_states,
        )
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _student_response(repo, student)


@router.put("/students/{student_pk}/tag-states", response_model=schemas.TutorStudentOut)
def set_student_tag_states(
    student_pk: str, payload: schemas.TutorStudentTagStatesIn, db: Session = Depends(get_db)
):
    repo = TutorRepository(db)
    try:
        student = repo.set_student_tag_states(student_pk, payload.tag_states)
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _student_response(repo, student)


@router.delete("/students/{student_pk}", status_code=204)
def delete_student(student_pk: str, db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    try:
        repo.delete_student(student_pk)
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error


def _student_availability_response(repo: TutorRepository, student_pk: str) -> schemas.TutorStudentAvailabilityIO:
    return schemas.TutorStudentAvailabilityIO(
        windows=[
            schemas.TutorAvailabilityWindowIO(
                day=window.day, start_minute=window.start_minute, end_minute=window.end_minute, status=window.status
            )
            for window in repo.student_availability(student_pk)
        ]
    )


@router.get("/students/{student_pk}/availability", response_model=schemas.TutorStudentAvailabilityIO)
def get_student_availability(student_pk: str, db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    try:
        repo.get_student(student_pk)
    except TutorValidationError as error:
        raise _http_error(error) from error
    return _student_availability_response(repo, student_pk)


@router.put("/students/{student_pk}/availability", response_model=schemas.TutorStudentAvailabilityIO)
def set_student_availability(
    student_pk: str, payload: schemas.TutorStudentAvailabilityIO, db: Session = Depends(get_db)
):
    repo = TutorRepository(db)
    windows = [
        AvailabilityWindow(
            day=normalize_day(item.day) or item.day,
            start_minute=item.start_minute,
            end_minute=item.end_minute,
            status=item.status,  # type: ignore[arg-type]
        )
        for item in payload.windows
    ]
    try:
        repo.set_student_availability(student_pk, windows)
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _student_availability_response(repo, student_pk)


# --- Tutor Tags ------------------------------------------------------------


@router.get("/tags", response_model=list[schemas.TutorTagOut])
def list_tags(db: Session = Depends(get_db)):
    return [_tag_out(summary) for summary in TutorRepository(db).tag_summaries()]


@router.post("/tags", response_model=schemas.TutorTagOut, status_code=201)
def create_tag(payload: schemas.TutorTagIn, db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    try:
        tag = repo.create_tag(payload.name)
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _tag_summary(repo, tag.id)


@router.patch("/tags/{tag_id}", response_model=schemas.TutorTagOut)
def rename_tag(tag_id: str, payload: schemas.TutorTagIn, db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    try:
        repo.rename_tag(tag_id, payload.name)
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _tag_summary(repo, tag_id)


@router.delete("/tags/{tag_id}", response_model=schemas.TutorTagDeleteOut)
def delete_tag(tag_id: str, db: Session = Depends(get_db)):
    repo = TutorRepository(db)
    try:
        summary = _tag_summary(repo, tag_id)
        repo.delete_tag(tag_id)
        _commit(db)
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return schemas.TutorTagDeleteOut(
        affected_tutor_count=summary.tutor_count, affected_student_count=summary.student_count
    )


# --- Tutor import --------------------------------------------------------


class TutorTutorImportResult(BaseModel):
    tutors_removed: int
    placements_removed: int
    tutors_created: int
    tags_created: int


def _import_http_error(error: importer.TutorTutorImportError) -> HTTPException:
    status = 413 if error.code == "FILE_TOO_LARGE" else 400
    return HTTPException(status, detail={"code": error.code, "message": str(error), "issues": error.issues})


async def _read_upload(file: UploadFile) -> bytes:
    content = await file.read(importer.MAX_UPLOAD_BYTES + 1)
    if len(content) > importer.MAX_UPLOAD_BYTES:
        raise importer.TutorTutorImportError("FILE_TOO_LARGE", "Tutor roster files must be 25 MB or smaller.")
    return content


def _parse_mapping(raw: str) -> importer.TutorTutorImportMapping:
    try:
        return importer.TutorTutorImportMapping.model_validate_json(raw)
    except ValidationError as error:
        raise importer.TutorTutorImportError("INVALID_MAPPING", "The column mapping is not valid.") from error


@router.post("/tutors/import/inspect", response_model=importer.TutorTutorImportInspection)
async def inspect_tutor_import(file: UploadFile = File(...), sheet_name: str | None = Form(default=None)):
    try:
        return importer.inspect_file(file.filename or "", await _read_upload(file), sheet_name or None)
    except importer.TutorTutorImportError as error:
        raise _import_http_error(error) from error


@router.post("/tutors/import/preview", response_model=importer.TutorTutorImportPreview)
async def preview_tutor_import(
    file: UploadFile = File(...),
    mapping: str = Form(...),
    sheet_name: str | None = Form(default=None),
    db: Session = Depends(get_db),
):
    """Phase A: parse and validate only. Nothing is written."""
    try:
        parsed = importer.parse_tutors(
            file.filename or "", await _read_upload(file), sheet_name or None, _parse_mapping(mapping)
        )
    except importer.TutorTutorImportError as error:
        raise _import_http_error(error) from error
    repo = TutorRepository(db)
    return importer.build_preview(
        parsed,
        {tag.normalized_name for tag in repo.list_tags()},
        len(repo.list_tutors()),
        db.query(TutorPlacement).count(),
    )


@router.post("/tutors/import/apply", response_model=TutorTutorImportResult)
async def apply_tutor_import(
    file: UploadFile = File(...),
    mapping: str = Form(...),
    sheet_name: str | None = Form(default=None),
    confirmed: bool = Form(default=False),
    db: Session = Depends(get_db),
):
    """Phase B: re-parse the file, then atomically replace all tutors."""
    try:
        if not confirmed:
            raise importer.TutorTutorImportError(
                "CONFIRMATION_REQUIRED", "Confirm the replacement before importing tutors."
            )
        parsed = importer.parse_tutors(
            file.filename or "", await _read_upload(file), sheet_name or None, _parse_mapping(mapping)
        )
        invalid = [row for row in parsed.rows if row.errors]
        if invalid:
            raise importer.TutorTutorImportError(
                "VALIDATION_FAILED",
                "Fix the listed rows in the source file and try again.",
                [f"Row {row.row_number}: {message}" for row in invalid[:25] for message in row.errors],
            )
        result = TutorRepository(db).replace_tutors(
            (
                TutorInput(
                    name=row.tutor_name,
                    email=row.tutor_email,
                    room=row.room,
                    lesson_length_minutes=row.lesson_length_minutes or 0,
                    capacity=row.capacity or 0,
                ),
                row.windows,
                row.tags,
            )
            for row in parsed.rows
        )
        db.commit()
        return TutorTutorImportResult(**result.__dict__)
    except importer.TutorTutorImportError as error:
        db.rollback()
        raise _import_http_error(error) from error
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    except Exception:
        db.rollback()
        raise


# --- Student import -----------------------------------------------------------


class TutorStudentImportResult(BaseModel):
    students_removed: int
    placements_removed: int
    students_created: int
    tags_created: int


async def _read_student_upload(file: UploadFile) -> bytes:
    content = await file.read(student_importer.MAX_UPLOAD_BYTES + 1)
    if len(content) > student_importer.MAX_UPLOAD_BYTES:
        raise student_importer.TutorStudentImportError(
            "FILE_TOO_LARGE", "Student roster files must be 25 MB or smaller."
        )
    return content


def _parse_student_mapping(raw: str) -> student_importer.TutorStudentImportMapping:
    try:
        return student_importer.TutorStudentImportMapping.model_validate_json(raw)
    except ValidationError as error:
        raise student_importer.TutorStudentImportError(
            "INVALID_MAPPING", "The column mapping is not valid."
        ) from error


@router.post("/students/import/inspect", response_model=student_importer.TutorStudentImportInspection)
async def inspect_student_import(file: UploadFile = File(...), sheet_name: str | None = Form(default=None)):
    try:
        return student_importer.inspect_file(file.filename or "", await _read_student_upload(file), sheet_name or None)
    except student_importer.TutorStudentImportError as error:
        raise _import_http_error(error) from error


@router.post("/students/import/preview", response_model=student_importer.TutorStudentImportPreview)
async def preview_student_import(
    file: UploadFile = File(...),
    mapping: str = Form(...),
    sheet_name: str | None = Form(default=None),
    db: Session = Depends(get_db),
):
    """Phase A: parse and validate only. Nothing is written."""
    try:
        parsed = student_importer.parse_students(
            file.filename or "", await _read_student_upload(file), sheet_name or None, _parse_student_mapping(mapping)
        )
    except student_importer.TutorStudentImportError as error:
        raise _import_http_error(error) from error
    repo = TutorRepository(db)
    return student_importer.build_preview(
        parsed,
        {tag.normalized_name for tag in repo.list_tags()},
        db.query(TutorStudent).count(),
        db.query(TutorPlacement).count(),
    )


@router.post("/students/import/apply", response_model=TutorStudentImportResult)
async def apply_student_import(
    file: UploadFile = File(...),
    mapping: str = Form(...),
    sheet_name: str | None = Form(default=None),
    confirmed: bool = Form(default=False),
    db: Session = Depends(get_db),
):
    """Phase B: re-parse the file, then atomically replace all students."""
    try:
        if not confirmed:
            raise student_importer.TutorStudentImportError(
                "CONFIRMATION_REQUIRED", "Confirm the replacement before importing students."
            )
        parsed = student_importer.parse_students(
            file.filename or "", await _read_student_upload(file), sheet_name or None, _parse_student_mapping(mapping)
        )
        invalid = [row for row in parsed.rows if row.errors]
        if invalid:
            raise student_importer.TutorStudentImportError(
                "VALIDATION_FAILED",
                "Fix the listed rows in the source file and try again.",
                [f"Row {row.row_number}: {message}" for row in invalid[:25] for message in row.errors],
            )
        result = TutorRepository(db).replace_students(
            (StudentInput(row.student_id, row.name, row.email), row.windows, row.tags) for row in parsed.rows
        )
        db.commit()
        return TutorStudentImportResult(**result.__dict__)
    except student_importer.TutorStudentImportError as error:
        db.rollback()
        raise _import_http_error(error) from error
    except TutorValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    except Exception:
        db.rollback()
        raise
