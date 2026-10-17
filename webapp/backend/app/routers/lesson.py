"""Lesson Scheduling module API: Instructor and Student Rosters, Lesson Tags, and the roster imports.

Handlers stay thin; domain rules and revision/staleness tracking live in LessonRepository.
Each request commits as one transaction.
"""

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from pydantic import ValidationError
from pydantic import BaseModel
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from .. import lesson_schemas as schemas
from ..database import get_db
from ..lesson_models import LessonInstructor, LessonPlacement, LessonStudent
from ..services import lesson_instructor_importer as importer
from ..services import lesson_student_importer as student_importer
from ..services.availability import AvailabilityWindow, normalize_day
from ..services.lesson_repository import (
    InstructorInput,
    LessonRepository,
    LessonValidationError,
    StudentInput,
)

router = APIRouter(prefix="/api/lesson", tags=["lesson scheduling"])

_NOT_FOUND_CODES = {"TAG_NOT_FOUND", "INSTRUCTOR_NOT_FOUND", "STUDENT_NOT_FOUND"}
_CONFLICT_CODES = {"DUPLICATE_TAG", "DUPLICATE_STUDENT_ID"}


def _http_error(error: LessonValidationError) -> HTTPException:
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
            detail={"code": "LESSON_CONFLICT", "message": "The change conflicts with existing Lesson Scheduling data."},
        ) from error


def _instructor_out(repo: LessonRepository, instructor: LessonInstructor) -> schemas.LessonInstructorOut:
    return schemas.LessonInstructorOut(
        id=instructor.id,
        name=instructor.name,
        email=instructor.email,
        room=instructor.room,
        lesson_length_minutes=instructor.lesson_length_minutes,
        capacity=instructor.capacity,
        tags=[schemas.LessonTagRef(id=tag.id, name=tag.name) for tag in instructor.tags],
        availability_window_count=repo.instructor_availability_counts().get(instructor.id, 0),
        created_at=instructor.created_at,
        updated_at=instructor.updated_at,
    )


def _fields(payload: schemas.LessonInstructorFields) -> dict:
    return {
        "name": payload.name,
        "email": payload.email,
        "room": payload.room,
        "lesson_length_minutes": payload.lesson_length_minutes,
        "capacity": payload.capacity,
    }


def _tag_out(summary) -> schemas.LessonTagOut:
    return schemas.LessonTagOut(
        id=summary.tag.id,
        name=summary.tag.name,
        instructor_count=summary.instructor_count,
        student_count=summary.student_count,
        is_unused=summary.is_unused,
    )


def _tag_summary(repo: LessonRepository, tag_id: str) -> schemas.LessonTagOut:
    for summary in repo.tag_summaries():
        if summary.tag.id == tag_id:
            return _tag_out(summary)
    raise HTTPException(404, detail={"code": "TAG_NOT_FOUND", "message": "The tag no longer exists."})


@router.get("/state", response_model=schemas.LessonModuleStateOut)
def read_state(db: Session = Depends(get_db)):
    state = LessonRepository(db).get_state()
    db.commit()
    return schemas.LessonModuleStateOut(
        schedule_revision=state.schedule_revision,
        source_data_revision=state.source_data_revision,
        is_schedule_stale=state.is_schedule_stale,
    )


# --- Instructors ------------------------------------------------------------


@router.get("/instructors", response_model=list[schemas.LessonInstructorOut])
def list_instructors(db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    return [_instructor_out(repo, item) for item in repo.list_instructors()]


@router.post("/instructors", response_model=schemas.LessonInstructorOut, status_code=201)
def create_instructor(payload: schemas.LessonInstructorIn, db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    try:
        instructor = repo.create_instructor(**_fields(payload), tag_ids=payload.tag_ids or [])
        _commit(db)
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _instructor_out(repo, instructor)


@router.put("/instructors/{instructor_id}", response_model=schemas.LessonInstructorOut)
def update_instructor(instructor_id: str, payload: schemas.LessonInstructorIn, db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    try:
        instructor = repo.update_instructor(instructor_id, **_fields(payload), tag_ids=payload.tag_ids)
        _commit(db)
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _instructor_out(repo, instructor)


@router.put("/instructors/{instructor_id}/tags", response_model=schemas.LessonInstructorOut)
def set_instructor_tags(instructor_id: str, payload: schemas.LessonInstructorTagsIn, db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    try:
        instructor = repo.set_instructor_tags(instructor_id, payload.tag_ids)
        _commit(db)
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _instructor_out(repo, instructor)


@router.delete("/instructors/{instructor_id}", response_model=schemas.LessonInstructorDeleteOut)
def delete_instructor(instructor_id: str, db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    try:
        removed = repo.delete_instructor(instructor_id)
        _commit(db)
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return schemas.LessonInstructorDeleteOut(removed_placement_count=removed)


def _availability_response(repo: LessonRepository, instructor_id: str) -> schemas.LessonInstructorAvailabilityIO:
    return schemas.LessonInstructorAvailabilityIO(
        windows=[
            schemas.LessonAvailabilityWindowIO(
                day=window.day, start_minute=window.start_minute, end_minute=window.end_minute, status=window.status
            )
            for window in repo.instructor_availability(instructor_id)
        ]
    )


@router.get("/instructors/{instructor_id}/availability", response_model=schemas.LessonInstructorAvailabilityIO)
def get_instructor_availability(instructor_id: str, db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    try:
        repo.get_instructor(instructor_id)
    except LessonValidationError as error:
        raise _http_error(error) from error
    return _availability_response(repo, instructor_id)


@router.put("/instructors/{instructor_id}/availability", response_model=schemas.LessonInstructorAvailabilityIO)
def set_instructor_availability(
    instructor_id: str, payload: schemas.LessonInstructorAvailabilityIO, db: Session = Depends(get_db)
):
    repo = LessonRepository(db)
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
        repo.set_instructor_availability(instructor_id, windows)
        _commit(db)
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _availability_response(repo, instructor_id)


# --- Students ---------------------------------------------------------------


def _student_out(student: LessonStudent, states: dict, availability_count: int) -> schemas.LessonStudentOut:
    return schemas.LessonStudentOut(
        id=student.id,
        student_id=student.student_id,
        name=student.name,
        email=student.email,
        availability_window_count=availability_count,
        tag_states={tag_id: state.value for tag_id, state in states.items()},
        created_at=student.created_at,
        updated_at=student.updated_at,
    )


def _student_response(repo: LessonRepository, student: LessonStudent) -> schemas.LessonStudentOut:
    return _student_out(
        student,
        repo.student_tag_states().get(student.id, {}),
        repo.student_availability_counts().get(student.id, 0),
    )


@router.get("/students", response_model=list[schemas.LessonStudentOut])
def list_students(db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    states = repo.student_tag_states()
    counts = repo.student_availability_counts()
    return [_student_out(item, states.get(item.id, {}), counts.get(item.id, 0)) for item in repo.list_students()]


@router.post("/students", response_model=schemas.LessonStudentOut, status_code=201)
def create_student(payload: schemas.LessonStudentIn, db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    try:
        student = repo.create_student(
            payload.student_id,
            payload.name,
            payload.email,
            require_name=True,
            tag_states=payload.tag_states,
        )
        _commit(db)
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _student_response(repo, student)


@router.put("/students/{student_pk}", response_model=schemas.LessonStudentOut)
def update_student(student_pk: str, payload: schemas.LessonStudentIn, db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    try:
        student = repo.update_student(
            student_pk,
            payload.student_id,
            payload.name,
            payload.email,
            require_name=True,
            tag_states=payload.tag_states,
        )
        _commit(db)
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _student_response(repo, student)


@router.put("/students/{student_pk}/tag-states", response_model=schemas.LessonStudentOut)
def set_student_tag_states(
    student_pk: str, payload: schemas.LessonStudentTagStatesIn, db: Session = Depends(get_db)
):
    repo = LessonRepository(db)
    try:
        student = repo.set_student_tag_states(student_pk, payload.tag_states)
        _commit(db)
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _student_response(repo, student)


@router.delete("/students/{student_pk}", status_code=204)
def delete_student(student_pk: str, db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    try:
        repo.delete_student(student_pk)
        _commit(db)
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error


def _student_availability_response(repo: LessonRepository, student_pk: str) -> schemas.LessonStudentAvailabilityIO:
    return schemas.LessonStudentAvailabilityIO(
        windows=[
            schemas.LessonAvailabilityWindowIO(
                day=window.day, start_minute=window.start_minute, end_minute=window.end_minute, status=window.status
            )
            for window in repo.student_availability(student_pk)
        ]
    )


@router.get("/students/{student_pk}/availability", response_model=schemas.LessonStudentAvailabilityIO)
def get_student_availability(student_pk: str, db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    try:
        repo.get_student(student_pk)
    except LessonValidationError as error:
        raise _http_error(error) from error
    return _student_availability_response(repo, student_pk)


@router.put("/students/{student_pk}/availability", response_model=schemas.LessonStudentAvailabilityIO)
def set_student_availability(
    student_pk: str, payload: schemas.LessonStudentAvailabilityIO, db: Session = Depends(get_db)
):
    repo = LessonRepository(db)
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
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _student_availability_response(repo, student_pk)


# --- Lesson Tags ------------------------------------------------------------


@router.get("/tags", response_model=list[schemas.LessonTagOut])
def list_tags(db: Session = Depends(get_db)):
    return [_tag_out(summary) for summary in LessonRepository(db).tag_summaries()]


@router.post("/tags", response_model=schemas.LessonTagOut, status_code=201)
def create_tag(payload: schemas.LessonTagIn, db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    try:
        tag = repo.create_tag(payload.name)
        _commit(db)
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _tag_summary(repo, tag.id)


@router.patch("/tags/{tag_id}", response_model=schemas.LessonTagOut)
def rename_tag(tag_id: str, payload: schemas.LessonTagIn, db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    try:
        repo.rename_tag(tag_id, payload.name)
        _commit(db)
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _tag_summary(repo, tag_id)


@router.delete("/tags/{tag_id}", response_model=schemas.LessonTagDeleteOut)
def delete_tag(tag_id: str, db: Session = Depends(get_db)):
    repo = LessonRepository(db)
    try:
        summary = _tag_summary(repo, tag_id)
        repo.delete_tag(tag_id)
        _commit(db)
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return schemas.LessonTagDeleteOut(
        affected_instructor_count=summary.instructor_count, affected_student_count=summary.student_count
    )


# --- Instructor import --------------------------------------------------------


class LessonInstructorImportResult(BaseModel):
    instructors_removed: int
    placements_removed: int
    instructors_created: int
    tags_created: int


def _import_http_error(error: importer.LessonInstructorImportError) -> HTTPException:
    status = 413 if error.code == "FILE_TOO_LARGE" else 400
    return HTTPException(status, detail={"code": error.code, "message": str(error), "issues": error.issues})


async def _read_upload(file: UploadFile) -> bytes:
    content = await file.read(importer.MAX_UPLOAD_BYTES + 1)
    if len(content) > importer.MAX_UPLOAD_BYTES:
        raise importer.LessonInstructorImportError("FILE_TOO_LARGE", "Instructor roster files must be 25 MB or smaller.")
    return content


def _parse_mapping(raw: str) -> importer.LessonInstructorImportMapping:
    try:
        return importer.LessonInstructorImportMapping.model_validate_json(raw)
    except ValidationError as error:
        raise importer.LessonInstructorImportError("INVALID_MAPPING", "The column mapping is not valid.") from error


@router.post("/instructors/import/inspect", response_model=importer.LessonInstructorImportInspection)
async def inspect_instructor_import(file: UploadFile = File(...), sheet_name: str | None = Form(default=None)):
    try:
        return importer.inspect_file(file.filename or "", await _read_upload(file), sheet_name or None)
    except importer.LessonInstructorImportError as error:
        raise _import_http_error(error) from error


@router.post("/instructors/import/preview", response_model=importer.LessonInstructorImportPreview)
async def preview_instructor_import(
    file: UploadFile = File(...),
    mapping: str = Form(...),
    sheet_name: str | None = Form(default=None),
    db: Session = Depends(get_db),
):
    """Phase A: parse and validate only. Nothing is written."""
    try:
        parsed = importer.parse_instructors(
            file.filename or "", await _read_upload(file), sheet_name or None, _parse_mapping(mapping)
        )
    except importer.LessonInstructorImportError as error:
        raise _import_http_error(error) from error
    repo = LessonRepository(db)
    return importer.build_preview(
        parsed,
        {tag.normalized_name for tag in repo.list_tags()},
        len(repo.list_instructors()),
        db.query(LessonPlacement).count(),
    )


@router.post("/instructors/import/apply", response_model=LessonInstructorImportResult)
async def apply_instructor_import(
    file: UploadFile = File(...),
    mapping: str = Form(...),
    sheet_name: str | None = Form(default=None),
    confirmed: bool = Form(default=False),
    db: Session = Depends(get_db),
):
    """Phase B: re-parse the file, then atomically replace all instructors."""
    try:
        if not confirmed:
            raise importer.LessonInstructorImportError(
                "CONFIRMATION_REQUIRED", "Confirm the replacement before importing instructors."
            )
        parsed = importer.parse_instructors(
            file.filename or "", await _read_upload(file), sheet_name or None, _parse_mapping(mapping)
        )
        invalid = [row for row in parsed.rows if row.errors]
        if invalid:
            raise importer.LessonInstructorImportError(
                "VALIDATION_FAILED",
                "Fix the listed rows in the source file and try again.",
                [f"Row {row.row_number}: {message}" for row in invalid[:25] for message in row.errors],
            )
        result = LessonRepository(db).replace_instructors(
            (
                InstructorInput(
                    name=row.instructor_name,
                    email=row.instructor_email,
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
        return LessonInstructorImportResult(**result.__dict__)
    except importer.LessonInstructorImportError as error:
        db.rollback()
        raise _import_http_error(error) from error
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    except Exception:
        db.rollback()
        raise


# --- Student import -----------------------------------------------------------


class LessonStudentImportResult(BaseModel):
    students_removed: int
    placements_removed: int
    students_created: int


async def _read_student_upload(file: UploadFile) -> bytes:
    content = await file.read(student_importer.MAX_UPLOAD_BYTES + 1)
    if len(content) > student_importer.MAX_UPLOAD_BYTES:
        raise student_importer.LessonStudentImportError(
            "FILE_TOO_LARGE", "Student roster files must be 25 MB or smaller."
        )
    return content


def _parse_student_mapping(raw: str) -> student_importer.LessonStudentImportMapping:
    try:
        return student_importer.LessonStudentImportMapping.model_validate_json(raw)
    except ValidationError as error:
        raise student_importer.LessonStudentImportError(
            "INVALID_MAPPING", "The column mapping is not valid."
        ) from error


@router.post("/students/import/inspect", response_model=student_importer.LessonStudentImportInspection)
async def inspect_student_import(file: UploadFile = File(...), sheet_name: str | None = Form(default=None)):
    try:
        return student_importer.inspect_file(file.filename or "", await _read_student_upload(file), sheet_name or None)
    except student_importer.LessonStudentImportError as error:
        raise _import_http_error(error) from error


@router.post("/students/import/preview", response_model=student_importer.LessonStudentImportPreview)
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
    except student_importer.LessonStudentImportError as error:
        raise _import_http_error(error) from error
    return student_importer.build_preview(
        parsed, db.query(LessonStudent).count(), db.query(LessonPlacement).count()
    )


@router.post("/students/import/apply", response_model=LessonStudentImportResult)
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
            raise student_importer.LessonStudentImportError(
                "CONFIRMATION_REQUIRED", "Confirm the replacement before importing students."
            )
        parsed = student_importer.parse_students(
            file.filename or "", await _read_student_upload(file), sheet_name or None, _parse_student_mapping(mapping)
        )
        invalid = [row for row in parsed.rows if row.errors]
        if invalid:
            raise student_importer.LessonStudentImportError(
                "VALIDATION_FAILED",
                "Fix the listed rows in the source file and try again.",
                [f"Row {row.row_number}: {message}" for row in invalid[:25] for message in row.errors],
            )
        result = LessonRepository(db).replace_students(
            (StudentInput(row.student_id, row.name, row.email), row.windows) for row in parsed.rows
        )
        db.commit()
        return LessonStudentImportResult(**result.__dict__)
    except student_importer.LessonStudentImportError as error:
        db.rollback()
        raise _import_http_error(error) from error
    except LessonValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    except Exception:
        db.rollback()
        raise
