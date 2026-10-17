from datetime import datetime, timezone

from fastapi import APIRouter, Depends, HTTPException, Path
from fastapi.responses import PlainTextResponse
from sqlalchemy.orm import Session

from .. import models, module_models, schemas
from ..database import get_db
from ..services import reports
from ..services.module_lifecycle import active_session_uuid

router = APIRouter(prefix="/api/reports", tags=["reports"])

MODULE_ID = Path(pattern=r"^[a-z][a-z0-9_-]{0,79}$")
MAX_SAVED_REPORT_BYTES = 10 * 1024 * 1024


@router.get("/markdown", response_class=PlainTextResponse)
def markdown_report(db: Session = Depends(get_db)):
    lessons = db.query(models.Lesson).all()
    pianists = db.query(models.Pianist).all()
    pianists_by_id = {p.id: p for p in pianists}
    return reports.build_markdown(lessons, pianists_by_id)


@router.get("/saved/{module_id}", response_model=schemas.SavedReportOut | None)
def get_saved_report(module_id: str = MODULE_ID, db: Session = Depends(get_db)):
    row = db.get(module_models.SavedReport, (active_session_uuid(db), module_id))
    if row is None:
        return None
    return schemas.SavedReportOut(
        **schemas.SavedReportContent.model_validate_json(row.payload_json).model_dump(),
        source_stamp=row.source_stamp,
        generated_at=row.generated_at,
    )


@router.put("/saved/{module_id}", response_model=schemas.SavedReportOut)
def put_saved_report(
    payload: schemas.SavedReportIn,
    module_id: str = MODULE_ID,
    db: Session = Depends(get_db),
):
    payload_json = payload.model_dump_json(exclude={"source_stamp"})
    if len(payload_json.encode("utf-8")) > MAX_SAVED_REPORT_BYTES:
        raise HTTPException(413, "The report is too large to save.")
    session_uuid = active_session_uuid(db)
    generated_at = datetime.now(timezone.utc).replace(tzinfo=None)
    row = db.get(module_models.SavedReport, (session_uuid, module_id))
    if row is None:
        row = module_models.SavedReport(session_uuid=session_uuid, module_id=module_id)
        db.add(row)
    row.payload_json = payload_json
    row.source_stamp = payload.source_stamp
    row.generated_at = generated_at
    db.commit()
    return schemas.SavedReportOut(**payload.model_dump(), generated_at=generated_at)
