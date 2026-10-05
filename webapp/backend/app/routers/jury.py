from datetime import date
from uuid import UUID

from fastapi import APIRouter, Depends, HTTPException, Response
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from .. import jury_schemas
from ..database import get_db
from ..services import jury, jury_results
from ..services.jury import JuryDataError
from ..services.jury_results import JuryGenerationError

router = APIRouter(prefix="/api/jury", tags=["jury inputs"])


def _jury_error(error: JuryDataError) -> HTTPException:
    status = 409 if error.code in {
        "NO_CURRENT_FINALIZED_RESULT",
        "JURY_DATE_MISMATCH",
        "OVERLAPPING_AVAILABILITY",
    } else 404 if error.code in {"PANEL_NOT_FOUND", "PIANIST_NOT_FOUND", "SOURCE_LESSON_NOT_FOUND"} else 422
    return HTTPException(status_code=status, detail={"code": error.code, "message": str(error)})


def _commit_or_conflict(db: Session):
    try:
        db.commit()
    except IntegrityError as error:
        db.rollback()
        raise HTTPException(
            status_code=409,
            detail={"code": "DUPLICATE_OR_INVALID_JURY_INPUT", "message": "Jury input conflicts with an existing record."},
        ) from error


@router.get("/configuration", response_model=jury_schemas.JuryConfigurationOut)
def get_configuration(db: Session = Depends(get_db)):
    try:
        return jury.get_configuration(db)
    except JuryDataError as error:
        raise _jury_error(error) from error


@router.get("/panels", response_model=list[jury_schemas.JuryPanelOut])
def list_panels(db: Session = Depends(get_db)):
    return jury.list_panels(db)


@router.post("/panels", response_model=jury_schemas.JuryPanelOut, status_code=201)
def create_panel(payload: jury_schemas.JuryPanelFields, db: Session = Depends(get_db)):
    try:
        result = jury.create_panel(db, payload)
        _commit_or_conflict(db)
        return result
    except IntegrityError as error:
        db.rollback()
        raise HTTPException(status_code=409, detail={"code": "DUPLICATE_PANEL_NAME", "message": "A Panel with this name already exists."}) from error


@router.put("/panels/{panel_uuid}", response_model=jury_schemas.JuryPanelOut)
def update_panel(panel_uuid: UUID, payload: jury_schemas.JuryPanelFields, db: Session = Depends(get_db)):
    try:
        result = jury.update_panel(db, str(panel_uuid), payload)
        _commit_or_conflict(db)
        return result
    except JuryDataError as error:
        db.rollback()
        raise _jury_error(error) from error
    except IntegrityError as error:
        db.rollback()
        raise HTTPException(status_code=409, detail={"code": "DUPLICATE_PANEL_NAME", "message": "A Panel with this name already exists."}) from error


@router.delete("/panels/{panel_uuid}", status_code=204)
def delete_panel(panel_uuid: UUID, db: Session = Depends(get_db)):
    try:
        jury.delete_panel(db, str(panel_uuid))
        db.commit()
        return Response(status_code=204)
    except JuryDataError as error:
        db.rollback()
        raise _jury_error(error) from error


@router.post("/roster/synchronize", response_model=list[jury_schemas.JuryLessonEntryOut])
def synchronize_roster(db: Session = Depends(get_db)):
    try:
        entries = jury.sync_roster_from_current_result(db)
        db.commit()
        return entries
    except JuryDataError as error:
        db.rollback()
        raise _jury_error(error) from error


@router.get("/entries", response_model=list[jury_schemas.JuryLessonEntryOut])
def list_entries(db: Session = Depends(get_db)):
    return jury.list_entries(db)


@router.post("/entries/assign-panels-by-instrument", response_model=jury_schemas.JuryPanelAutoAssignmentOut)
def assign_panels_by_instrument(db: Session = Depends(get_db)):
    assigned_count, entries = jury.assign_panels_by_instrument(db)
    db.commit()
    return jury_schemas.JuryPanelAutoAssignmentOut(assigned_count=assigned_count, entries=entries)


@router.patch("/entries/{source_lesson_uuid}", response_model=jury_schemas.JuryLessonEntryOut)
def update_entry(
    source_lesson_uuid: UUID,
    payload: jury_schemas.JuryLessonEntryUpdate,
    db: Session = Depends(get_db),
):
    try:
        result = jury.update_lesson_entry(
            db,
            str(source_lesson_uuid),
            panel_uuid=str(payload.panel_uuid) if payload.panel_uuid else None,
        )
        db.commit()
        return result
    except JuryDataError as error:
        db.rollback()
        raise _jury_error(error) from error


@router.patch("/entries/{source_lesson_uuid}/jury-required", response_model=jury_schemas.JuryLessonEntryOut)
def update_jury_required(
    source_lesson_uuid: UUID,
    payload: jury_schemas.JuryRequiredUpdate,
    db: Session = Depends(get_db),
):
    try:
        result = jury.update_lesson_jury_required(db, str(source_lesson_uuid), payload.jury_required)
        db.commit()
        return result
    except JuryDataError as error:
        db.rollback()
        raise _jury_error(error) from error


@router.put(
    "/pianists/{pianist_person_uuid}/availability/{jury_date}",
    response_model=jury_schemas.JuryAvailabilityOut,
)
def save_availability(
    pianist_person_uuid: UUID,
    jury_date: date,
    payload: jury_schemas.JuryAvailabilityIn,
    db: Session = Depends(get_db),
):
    try:
        result = jury.save_availability(db, str(pianist_person_uuid), jury_date, payload)
        db.commit()
        return result
    except JuryDataError as error:
        db.rollback()
        raise _jury_error(error) from error


@router.get(
    "/pianists/{pianist_person_uuid}/availability/{jury_date}",
    response_model=jury_schemas.JuryAvailabilityOut,
)
def get_availability(
    pianist_person_uuid: UUID,
    jury_date: date,
    db: Session = Depends(get_db),
):
    result = jury.get_availability(db, str(pianist_person_uuid), jury_date)
    if result is None:
        raise HTTPException(status_code=404, detail={"code": "AVAILABILITY_NOT_FOUND", "message": "No Availability Windows are saved for this pianist and Scheduling Date."})
    return result


@router.get("/readiness", response_model=jury_schemas.JuryReadinessOut)
def readiness(db: Session = Depends(get_db)):
    try:
        return jury.readiness(db)
    except JuryDataError as error:
        raise _jury_error(error) from error


def _generation_error(error: JuryGenerationError) -> HTTPException:
    status = 409 if error.code in {
        "JURY_INPUT_REVISION_CHANGED",
        "JURY_NOT_READY",
        "INPUT_REVISION_CHANGED",
    } else 404 if error.code == "JURY_CONFIGURATION_MISSING" else 422
    detail = {"code": error.code, "message": str(error)}
    if error.readiness is not None:
        detail["readiness"] = error.readiness.model_dump(mode="json")
    return HTTPException(status_code=status, detail=detail)


@router.post("/generate", response_model=jury_schemas.JuryScheduleViewOut)
def generate_schedule(payload: jury_schemas.JuryGenerateRequest, db: Session = Depends(get_db)):
    try:
        result = jury_results.generate_schedule(
            db,
            expected_jury_input_revision=payload.expected_jury_input_revision,
        )
        db.commit()
        return result
    except JuryGenerationError as error:
        db.rollback()
        raise _generation_error(error) from error


@router.get("/results/current", response_model=jury_schemas.JuryScheduleViewOut)
def get_current_schedule(db: Session = Depends(get_db)):
    result = jury_results.get_current_schedule(db)
    if result is None:
        raise HTTPException(
            status_code=404,
            detail={"code": "NO_CURRENT_JURY_SCHEDULE", "message": "No Jury schedule has been generated for the active session."},
        )
    return result


@router.get("/results/history", response_model=list[jury_schemas.JuryScheduleHistoryItemOut])
def get_schedule_history(db: Session = Depends(get_db)):
    return jury_results.list_schedule_history(db)


@router.get("/results/{result_uuid}", response_model=jury_schemas.JuryScheduleViewOut)
def get_schedule_result(result_uuid: UUID, db: Session = Depends(get_db)):
    result = jury_results.get_schedule_by_uuid(db, str(result_uuid))
    if result is None:
        raise HTTPException(
            status_code=404,
            detail={"code": "JURY_SCHEDULE_NOT_FOUND", "message": "The requested Jury schedule is not available."},
        )
    return result