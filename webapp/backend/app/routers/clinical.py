"""Clinical Placement module API: Placement Opportunities and Clinical Tags.

Handlers stay thin; domain rules and revision/staleness tracking live in
ClinicalRepository. Each request commits as one transaction.
"""

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from pydantic import BaseModel, ValidationError
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from .. import clinical_schemas as schemas
from ..clinical_models import ClinicalPlacement, PlacementOpportunity
from ..database import get_db
from ..services import clinical_sites_importer as importer
from ..services.clinical_repository import (
    ClinicalRepository,
    ClinicalValidationError,
    PlacementOpportunityInput,
)

router = APIRouter(prefix="/api/clinical", tags=["clinical placements"])

_NOT_FOUND_CODES = {"TAG_NOT_FOUND", "OPPORTUNITY_NOT_FOUND"}
_CONFLICT_CODES = {"DUPLICATE_TAG"}


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


def _import_http_error(error: importer.ClinicalSitesImportError) -> HTTPException:
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
