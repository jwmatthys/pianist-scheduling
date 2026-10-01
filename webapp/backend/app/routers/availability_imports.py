from fastapi import APIRouter, Depends, File, HTTPException, Query, UploadFile
from sqlalchemy.orm import Session

from .. import schemas
from ..database import get_db
from ..services import availability_importer
from ..services.availability_importer import AvailabilityImportError

router = APIRouter(prefix="/api/availability-import", tags=["availability import"])


def _http_error(error: AvailabilityImportError) -> HTTPException:
    status_code = 413 if error.code == "FILE_TOO_LARGE" else 400
    if error.code in {"CONFIRMATION_REQUIRED", "PIANISTS_CHANGED", "PREVIEW_EXPIRED"}:
        status_code = 409
    return HTTPException(status_code, detail={"code": error.code, "message": str(error)})


@router.post("/inspect", response_model=schemas.AvailabilityImportInspection)
async def inspect_file(file: UploadFile = File(...)):
    try:
        content = await file.read(availability_importer.MAX_UPLOAD_BYTES + 1)
        if len(content) > availability_importer.MAX_UPLOAD_BYTES:
            raise AvailabilityImportError("FILE_TOO_LARGE", "Availability files must be 25 MB or smaller.")
        return availability_importer.inspect_upload(file.filename or "", content)
    except AvailabilityImportError as error:
        raise _http_error(error) from error


@router.get("/inspect/{upload_token}", response_model=schemas.AvailabilityImportInspection)
def inspect_sheet(upload_token: str, sheet_name: str | None = Query(default=None)):
    try:
        return availability_importer.inspect_sheet(upload_token, sheet_name)
    except AvailabilityImportError as error:
        raise _http_error(error) from error


@router.post("/preview", response_model=schemas.AvailabilityImportPreviewOut)
def preview_import(payload: schemas.AvailabilityImportPreviewRequest, db: Session = Depends(get_db)):
    try:
        return availability_importer.preview_import(payload, db)
    except AvailabilityImportError as error:
        raise _http_error(error) from error


@router.post("/apply", response_model=schemas.AvailabilityImportApplyResult)
def apply_import(payload: schemas.AvailabilityImportApplyRequest, db: Session = Depends(get_db)):
    try:
        return availability_importer.apply_import(payload, db)
    except AvailabilityImportError as error:
        raise _http_error(error) from error