from fastapi import APIRouter, HTTPException, Query, Request, Response

from .. import database, schemas
from ..database import engine
from ..services.session_files import (
    MAX_ARCHIVE_BYTES,
    SessionFileError,
    create_new_session,
    export_session_archive,
    get_active_session,
    restore_session,
    update_session_metadata,
)

router = APIRouter(prefix="/api/session", tags=["session"])


def _http_error(error: SessionFileError) -> HTTPException:
    status = 413 if error.code == "SESSION_TOO_LARGE" else 400
    if error.code in {"DATABASE_BUSY", "SESSION_REPLACEMENT_FAILED"}:
        status = 409
    return HTTPException(status_code=status, detail={"code": error.code, "message": str(error)})


@router.get("", response_model=schemas.SessionMetadataOut)
def read_active_session():
    try:
        return get_active_session(engine)
    except SessionFileError as error:
        raise _http_error(error) from error


@router.post("/new", response_model=schemas.SessionMetadataOut)
def new_session(payload: schemas.SessionCreateRequest):
    if not payload.confirmed:
        raise HTTPException(409, detail={"code": "CONFIRMATION_REQUIRED", "message": "Confirm replacing the active session."})
    try:
        metadata = schemas.SessionMetadataIn.model_validate(payload.model_dump(exclude={"confirmed"}))
        return create_new_session(engine, metadata)
    except SessionFileError as error:
        raise _http_error(error) from error


@router.get("/export")
def export_active_session():
    try:
        archive_bytes = export_session_archive(engine)
    except SessionFileError as error:
        raise _http_error(error) from error
    return Response(
        content=archive_bytes,
        media_type="application/vnd.music-program-scheduler.session+zip",
        headers={"Content-Disposition": 'attachment; filename="music-program-session.mpsession"'},
    )


@router.put("", response_model=schemas.SessionMetadataOut)
def update_active_session(payload: schemas.SessionMetadataIn):
    try:
        return update_session_metadata(engine, payload)
    except SessionFileError as error:
        raise _http_error(error) from error


@router.post("/restore", response_model=schemas.SessionMetadataOut)
async def restore_active_session(request: Request, confirmed: bool = Query(default=False)):
    if not confirmed:
        raise HTTPException(409, detail={"code": "CONFIRMATION_REQUIRED", "message": "Confirm replacing the active session."})
    archive = bytearray()
    async for chunk in request.stream():
        archive.extend(chunk)
        if len(archive) > MAX_ARCHIVE_BYTES:
            raise HTTPException(413, detail={"code": "SESSION_TOO_LARGE", "message": "The selected session archive exceeds the supported size."})
    try:
        return restore_session(engine, bytes(archive))
    except SessionFileError as error:
        raise _http_error(error) from error