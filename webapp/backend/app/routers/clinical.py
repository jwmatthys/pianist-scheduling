"""Clinical Placement module API: Placement Opportunities and Clinical Tags.

Handlers stay thin; domain rules and revision/staleness tracking live in
ClinicalRepository. Each request commits as one transaction.
"""

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from pydantic import BaseModel, ValidationError
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from .. import clinical_schemas as schemas
from ..clinical_models import ClinicalPlacement, ClinicalStudent, PlacementOpportunity, TagState
from ..database import get_db
from ..services import clinical_sites_importer as importer
from ..services.availability import AvailabilityWindow, normalize_day
from ..services import clinical_students_importer as students_importer
from ..services.clinical_candidates import generate_candidates
from ..services.clinical_import_common import ClinicalImportError
from ..services.clinical_readiness import analyze_readiness
from ..services.clinical_repository import (
    ClinicalRepository,
    ClinicalValidationError,
    PlacementOpportunityInput,
)

router = APIRouter(prefix="/api/clinical", tags=["clinical placements"])

_NOT_FOUND_CODES = {"TAG_NOT_FOUND", "OPPORTUNITY_NOT_FOUND", "STUDENT_NOT_FOUND"}
_CONFLICT_CODES = {"DUPLICATE_TAG", "DUPLICATE_STUDENT_ID", "NOT_READY"}


def _http_error(error: ClinicalValidationError) -> HTTPException:
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
            detail={"code": "CLINICAL_CONFLICT", "message": "The change conflicts with existing Clinical data."},
        ) from error


def _opportunity_out(repo: ClinicalRepository, opportunity: PlacementOpportunity) -> schemas.PlacementOpportunityOut:
    return schemas.PlacementOpportunityOut(
        id=opportunity.id,
        site_name=opportunity.site_name,
        site_address=opportunity.site_address,
        contact_name=opportunity.contact_name,
        contact_email=opportunity.contact_email,
        contact_phone=opportunity.contact_phone,
        day_of_week=opportunity.day_of_week,
        start_minute=opportunity.start_minute,
        end_minute=opportunity.end_minute,
        session_length_minutes=opportunity.session_length_minutes,
        travel_minutes=opportunity.travel_minutes,
        capacity=opportunity.capacity,
        tags=[schemas.ClinicalTagRef(id=tag.id, name=tag.name) for tag in opportunity.tags],
        placement_count=repo.count_placements_for_opportunity(opportunity.id),
        created_at=opportunity.created_at,
        updated_at=opportunity.updated_at,
    )


def _input(payload: schemas.PlacementOpportunityFields) -> PlacementOpportunityInput:
    return PlacementOpportunityInput(
        site_name=payload.site_name,
        site_address=payload.site_address,
        contact_name=payload.contact_name,
        contact_email=payload.contact_email,
        contact_phone=payload.contact_phone,
        day_of_week=payload.day_of_week,
        start_minute=payload.start_minute,
        end_minute=payload.end_minute,
        session_length_minutes=payload.session_length_minutes,
        travel_minutes=payload.travel_minutes,
        capacity=payload.capacity,
    )


def _tag_out(summary) -> schemas.ClinicalTagOut:
    return schemas.ClinicalTagOut(
        id=summary.tag.id,
        name=summary.tag.name,
        opportunity_count=summary.opportunity_count,
        student_count=summary.student_count,
        is_unused=summary.is_unused,
    )


def _tag_summary(repo: ClinicalRepository, tag_id: str) -> schemas.ClinicalTagOut:
    for summary in repo.tag_summaries():
        if summary.tag.id == tag_id:
            return _tag_out(summary)
    raise HTTPException(404, detail={"code": "TAG_NOT_FOUND", "message": "The tag no longer exists."})


@router.get("/state", response_model=schemas.ClinicalModuleStateOut)
def read_state(db: Session = Depends(get_db)):
    state = ClinicalRepository(db).get_state()
    db.commit()
    return schemas.ClinicalModuleStateOut(
        schedule_revision=state.schedule_revision,
        source_data_revision=state.source_data_revision,
        is_schedule_stale=state.is_schedule_stale,
    )


@router.get("/opportunities", response_model=list[schemas.PlacementOpportunityOut])
def list_opportunities(db: Session = Depends(get_db)):
    repo = ClinicalRepository(db)
    return [_opportunity_out(repo, item) for item in repo.list_opportunities()]


@router.post("/opportunities", response_model=schemas.PlacementOpportunityOut, status_code=201)
def create_opportunity(payload: schemas.PlacementOpportunityIn, db: Session = Depends(get_db)):
    repo = ClinicalRepository(db)
    try:
        opportunity = repo.create_opportunity(_input(payload), tag_ids=payload.tag_ids)
        _commit(db)
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _opportunity_out(repo, opportunity)


@router.put("/opportunities/{opportunity_id}", response_model=schemas.PlacementOpportunityOut)
def update_opportunity(
    opportunity_id: str, payload: schemas.PlacementOpportunityUpdate, db: Session = Depends(get_db)
):
    repo = ClinicalRepository(db)
    try:
        opportunity = repo.update_opportunity(opportunity_id, _input(payload), tag_ids=payload.tag_ids)
        _commit(db)
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _opportunity_out(repo, opportunity)


@router.put("/opportunities/{opportunity_id}/tags", response_model=schemas.PlacementOpportunityOut)
def set_opportunity_tags(
    opportunity_id: str, payload: schemas.PlacementOpportunityTagsIn, db: Session = Depends(get_db)
):
    repo = ClinicalRepository(db)
    try:
        opportunity = repo.set_opportunity_tags(opportunity_id, payload.tag_ids)
        _commit(db)
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _opportunity_out(repo, opportunity)


@router.delete("/opportunities/{opportunity_id}", response_model=schemas.PlacementOpportunityDeleteOut)
def delete_opportunity(opportunity_id: str, db: Session = Depends(get_db)):
    repo = ClinicalRepository(db)
    try:
        removed = repo.delete_opportunity(opportunity_id)
        _commit(db)
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return schemas.PlacementOpportunityDeleteOut(removed_placement_count=removed)


@router.get("/tags", response_model=list[schemas.ClinicalTagOut])
def list_tags(db: Session = Depends(get_db)):
    return [_tag_out(summary) for summary in ClinicalRepository(db).tag_summaries()]


@router.post("/tags", response_model=schemas.ClinicalTagOut, status_code=201)
def create_tag(payload: schemas.ClinicalTagIn, db: Session = Depends(get_db)):
    repo = ClinicalRepository(db)
    try:
        tag = repo.create_tag(payload.name)
        _commit(db)
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _tag_summary(repo, tag.id)


@router.patch("/tags/{tag_id}", response_model=schemas.ClinicalTagOut)
def rename_tag(tag_id: str, payload: schemas.ClinicalTagIn, db: Session = Depends(get_db)):
    repo = ClinicalRepository(db)
    try:
        repo.rename_tag(tag_id, payload.name)
        _commit(db)
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _tag_summary(repo, tag_id)


@router.delete("/tags/{tag_id}", response_model=schemas.ClinicalTagDeleteOut)
def delete_tag(tag_id: str, db: Session = Depends(get_db)):
    repo = ClinicalRepository(db)
    try:
        affected = repo.delete_tag(tag_id)
        _commit(db)
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return schemas.ClinicalTagDeleteOut(affected_student_count=affected)


# --- Clinical Sites import -------------------------------------------------


class ClinicalSitesImportResult(BaseModel):
    opportunities_removed: int
    placements_removed: int
    opportunities_created: int
    tags_created: int


def _import_http_error(error: ClinicalImportError) -> HTTPException:
    status = 413 if error.code == "FILE_TOO_LARGE" else 400
    return HTTPException(status, detail={"code": error.code, "message": str(error), "issues": error.issues})


async def _read_upload(file: UploadFile) -> bytes:
    content = await file.read(importer.MAX_UPLOAD_BYTES + 1)
    if len(content) > importer.MAX_UPLOAD_BYTES:
        raise importer.ClinicalSitesImportError("FILE_TOO_LARGE", "Clinical Sites files must be 25 MB or smaller.")
    return content


def _parse_mapping(raw: str) -> importer.ClinicalSitesImportMapping:
    try:
        return importer.ClinicalSitesImportMapping.model_validate_json(raw)
    except ValidationError as error:
        raise importer.ClinicalSitesImportError("INVALID_MAPPING", "The column mapping is not valid.") from error


@router.post("/opportunities/import/inspect", response_model=importer.ClinicalSitesImportInspection)
async def inspect_sites_import(file: UploadFile = File(...), sheet_name: str | None = Form(default=None)):
    try:
        return importer.inspect_file(file.filename or "", await _read_upload(file), sheet_name or None)
    except importer.ClinicalSitesImportError as error:
        raise _import_http_error(error) from error


@router.post("/opportunities/import/preview", response_model=importer.ClinicalSitesImportPreview)
async def preview_sites_import(
    file: UploadFile = File(...),
    mapping: str = Form(...),
    sheet_name: str | None = Form(default=None),
    db: Session = Depends(get_db),
):
    """Phase A: parse and validate only. Nothing is written."""
    try:
        parsed = importer.parse_sites(
            file.filename or "", await _read_upload(file), sheet_name or None, _parse_mapping(mapping)
        )
    except importer.ClinicalSitesImportError as error:
        raise _import_http_error(error) from error
    repo = ClinicalRepository(db)
    return importer.build_preview(
        parsed,
        {tag.normalized_name for tag in repo.list_tags()},
        len(repo.list_opportunities()),
        db.query(ClinicalPlacement).count(),
    )


@router.post("/opportunities/import/apply", response_model=ClinicalSitesImportResult)
async def apply_sites_import(
    file: UploadFile = File(...),
    mapping: str = Form(...),
    sheet_name: str | None = Form(default=None),
    confirmed: bool = Form(default=False),
    db: Session = Depends(get_db),
):
    """Phase B: re-parse the file, then atomically replace all Placement Opportunities."""
    try:
        if not confirmed:
            raise importer.ClinicalSitesImportError(
                "CONFIRMATION_REQUIRED", "Confirm the replacement before importing Clinical Sites."
            )
        parsed = importer.parse_sites(
            file.filename or "", await _read_upload(file), sheet_name or None, _parse_mapping(mapping)
        )
        invalid = [site for site in parsed.rows if site.errors]
        if invalid:
            raise importer.ClinicalSitesImportError(
                "VALIDATION_FAILED",
                "Fix the listed rows in the source file and try again.",
                [f"Row {site.row_number}: {message}" for site in invalid[:25] for message in site.errors],
            )
        result = ClinicalRepository(db).replace_opportunities(
            (
                PlacementOpportunityInput(
                    site_name=site.site_name,
                    site_address=site.site_address,
                    contact_name=site.contact_name,
                    contact_email=site.contact_email,
                    contact_phone=site.contact_phone,
                    day_of_week=site.day_of_week or "",
                    start_minute=site.start_minute or 0,
                    end_minute=site.end_minute or 0,
                    session_length_minutes=site.session_length_minutes or 0,
                    travel_minutes=site.travel_minutes or 0,
                    capacity=site.capacity or 1,
                ),
                site.tags,
            )
            for site in parsed.rows
        )
        db.commit()
        return ClinicalSitesImportResult(**result.__dict__)
    except importer.ClinicalSitesImportError as error:
        db.rollback()
        raise _import_http_error(error) from error
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    except Exception:
        db.rollback()
        raise


# --- Student Roster -------------------------------------------------------


def _student_out(
    student: ClinicalStudent, states: dict[str, TagState], availability_count: int
) -> schemas.ClinicalStudentOut:
    return schemas.ClinicalStudentOut(
        id=student.id,
        student_id=student.student_id,
        name=student.name,
        email=student.email,
        availability_window_count=availability_count,
        tag_states={tag_id: state.value for tag_id, state in states.items()},
        created_at=student.created_at,
        updated_at=student.updated_at,
    )


def _student_response(repo: ClinicalRepository, student: ClinicalStudent) -> schemas.ClinicalStudentOut:
    return _student_out(
        student,
        repo.student_tag_states().get(student.id, {}),
        repo.student_availability_counts().get(student.id, 0),
    )


@router.get("/students", response_model=list[schemas.ClinicalStudentOut])
def list_students(db: Session = Depends(get_db)):
    repo = ClinicalRepository(db)
    states = repo.student_tag_states()
    counts = repo.student_availability_counts()
    return [_student_out(item, states.get(item.id, {}), counts.get(item.id, 0)) for item in repo.list_students()]


@router.post("/students", response_model=schemas.ClinicalStudentOut, status_code=201)
def create_student(payload: schemas.ClinicalStudentFields, db: Session = Depends(get_db)):
    repo = ClinicalRepository(db)
    try:
        student = repo.create_student(payload.student_id, payload.name, payload.email, require_name=True)
        _commit(db)
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _student_response(repo, student)


@router.put("/students/{student_pk}", response_model=schemas.ClinicalStudentOut)
def update_student(student_pk: str, payload: schemas.ClinicalStudentFields, db: Session = Depends(get_db)):
    repo = ClinicalRepository(db)
    try:
        student = repo.update_student(student_pk, payload.student_id, payload.name, payload.email)
        _commit(db)
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _student_response(repo, student)


@router.put("/students/{student_pk}/tag-states", response_model=schemas.ClinicalStudentOut)
def set_student_tag_states(
    student_pk: str, payload: schemas.ClinicalStudentTagStatesIn, db: Session = Depends(get_db)
):
    repo = ClinicalRepository(db)
    try:
        student = repo.get_student(student_pk)
        repo.set_student_tag_states(student_pk, payload.tag_states)
        _commit(db)
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _student_response(repo, student)


def _availability_response(repo: ClinicalRepository, student_pk: str) -> schemas.ClinicalStudentAvailabilityIO:
    return schemas.ClinicalStudentAvailabilityIO(windows=[
        schemas.ClinicalAvailabilityWindowIO(
            day=item.day, start_minute=item.start_minute, end_minute=item.end_minute, status=item.status
        )
        for item in repo.student_availability(student_pk)
    ])


@router.get("/students/{student_pk}/availability", response_model=schemas.ClinicalStudentAvailabilityIO)
def get_student_availability(student_pk: str, db: Session = Depends(get_db)):
    repo = ClinicalRepository(db)
    try:
        repo.get_student(student_pk)
    except ClinicalValidationError as error:
        raise _http_error(error) from error
    return _availability_response(repo, student_pk)


@router.put("/students/{student_pk}/availability", response_model=schemas.ClinicalStudentAvailabilityIO)
def set_student_availability(
    student_pk: str, payload: schemas.ClinicalStudentAvailabilityIO, db: Session = Depends(get_db)
):
    repo = ClinicalRepository(db)
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
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    return _availability_response(repo, student_pk)


@router.delete("/students/{student_pk}", status_code=204)
def delete_student(student_pk: str, db: Session = Depends(get_db)):
    try:
        ClinicalRepository(db).delete_student(student_pk)
        _commit(db)
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error


# --- Student Roster import -------------------------------------------------


class ClinicalStudentsImportResult(BaseModel):
    students_removed: int
    placements_removed: int
    students_created: int


def _parse_students_mapping(raw: str) -> students_importer.ClinicalStudentsImportMapping:
    try:
        return students_importer.ClinicalStudentsImportMapping.model_validate_json(raw)
    except ValidationError as error:
        raise ClinicalImportError("INVALID_MAPPING", "The column mapping is not valid.") from error


@router.post("/students/import/inspect", response_model=students_importer.ClinicalStudentsImportInspection)
async def inspect_students_import(file: UploadFile = File(...), sheet_name: str | None = Form(default=None)):
    try:
        return students_importer.inspect_file(file.filename or "", await _read_upload(file), sheet_name or None)
    except ClinicalImportError as error:
        raise _import_http_error(error) from error


@router.post("/students/import/preview", response_model=students_importer.ClinicalStudentsImportPreview)
async def preview_students_import(
    file: UploadFile = File(...),
    mapping: str = Form(...),
    sheet_name: str | None = Form(default=None),
    db: Session = Depends(get_db),
):
    """Phase A: parse and validate only. Nothing is written."""
    try:
        parsed = students_importer.parse_students(
            file.filename or "", await _read_upload(file), sheet_name or None, _parse_students_mapping(mapping)
        )
    except ClinicalImportError as error:
        raise _import_http_error(error) from error
    repo = ClinicalRepository(db)
    return students_importer.build_preview(
        parsed, len(repo.list_students()), db.query(ClinicalPlacement).count()
    )


@router.post("/students/import/apply", response_model=ClinicalStudentsImportResult)
async def apply_students_import(
    file: UploadFile = File(...),
    mapping: str = Form(...),
    sheet_name: str | None = Form(default=None),
    confirmed: bool = Form(default=False),
    db: Session = Depends(get_db),
):
    """Phase B: re-parse the file, then atomically replace the whole Student Roster."""
    try:
        if not confirmed:
            raise ClinicalImportError(
                "CONFIRMATION_REQUIRED", "Confirm the replacement before importing the Student Roster."
            )
        parsed = students_importer.parse_students(
            file.filename or "", await _read_upload(file), sheet_name or None, _parse_students_mapping(mapping)
        )
        invalid = [row for row in parsed.rows if row.errors]
        if invalid:
            raise ClinicalImportError(
                "VALIDATION_FAILED",
                "Fix the listed rows in the source file and try again.",
                [f"Row {row.row_number}: {message}" for row in invalid[:25] for message in row.errors],
            )
        if not parsed.rows:
            raise ClinicalImportError("EMPTY_IMPORT", "The selected file contains no students.")
        result = ClinicalRepository(db).replace_students(
            (row.student_id, row.name, row.email) for row in parsed.rows
        )
        db.commit()
        return ClinicalStudentsImportResult(**result.__dict__)
    except ClinicalImportError as error:
        db.rollback()
        raise _import_http_error(error) from error
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error) from error
    except Exception:
        db.rollback()
        raise


# --- Readiness and candidates (read-only; never assigns students) ---------


@router.get("/readiness", response_model=schemas.ClinicalReadinessOut)
def get_readiness(db: Session = Depends(get_db)):
    scheduling_input = ClinicalRepository(db).scheduling_input()
    candidates = generate_candidates(scheduling_input)
    report = analyze_readiness(scheduling_input, candidates)
    return schemas.ClinicalReadinessOut(
        source_data_revision=report.source_data_revision,
        is_ready=report.is_ready,
        student_count=report.student_count,
        opportunity_count=report.opportunity_count,
        feasible_candidate_count=report.feasible_candidate_count,
        total_candidate_count=report.total_candidate_count,
        students_with_candidates=len({item.student_id for item in candidates.feasible}),
        warnings=[
            schemas.ClinicalReadinessWarningOut(
                code=item.code.value,
                severity=item.severity.value,
                category=item.category.value,
                subject=None if item.subject is None else schemas.ClinicalReadinessSubjectOut(
                    kind=item.subject.kind.value, id=item.subject.id, label=item.subject.label
                ),
            )
            for item in report.warnings
        ],
    )


# --- Schedule generation ---------------------------------------------------


def _schedule_out(db: Session) -> schemas.ClinicalScheduleOut:
    repo = ClinicalRepository(db)
    state = repo.get_state()
    summary = repo.schedule_summary()
    records = repo.list_placement_records()
    names = {s.id: s.label for s in repo.scheduling_input().students}
    return schemas.ClinicalScheduleOut(
        schedule_revision=state.schedule_revision,
        source_data_revision=state.source_data_revision,
        is_schedule_stale=state.is_schedule_stale,
        has_placements=bool(records),
        student_count=summary.student_count,
        placed_count=summary.placed_count,
        unplaced_count=summary.unplaced_count,
        used_capacity=summary.used_capacity,
        total_capacity=summary.total_capacity,
        utilization=summary.utilization,
        preferred_tag_matches=summary.preferred_tag_matches,
        tentative_minutes=summary.tentative_minutes,
        unverified_placement_count=summary.unverified_placement_count,
        unplaced=[
            schemas.ClinicalUnplacedOut(student_id=u.student_id, student_name=names.get(u.student_id, ""), reason=u.reason.value)
            for u in summary.unplaced
        ],
        placements=[
            schemas.ClinicalPlacementOut(
                student_id=r.student_id, kind=r.kind.value, opportunity_id=r.opportunity_id, start_minute=r.start_minute
            )
            for r in records
        ],
    )


@router.get("/schedule", response_model=schemas.ClinicalScheduleOut)
def get_schedule(db: Session = Depends(get_db)):
    return _schedule_out(db)


@router.post("/schedule/generate", response_model=schemas.ClinicalScheduleOut)
def generate_schedule(db: Session = Depends(get_db)):
    try:
        ClinicalRepository(db).generate_schedule()
        _commit(db)
    except ClinicalValidationError as error:
        db.rollback()
        raise _http_error(error)
    return _schedule_out(db)


@router.post("/schedule/clear", response_model=schemas.ClinicalScheduleOut)
def clear_schedule(db: Session = Depends(get_db)):
    ClinicalRepository(db).clear_all_placements()
    _commit(db)
    return _schedule_out(db)
