"""HTTP endpoints for offline license activation."""

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile

from ..license_schemas import LicenseKeyRequest, LicenseStatus
from ..services.license_service import LicenseService, LicenseServiceError, service

router = APIRouter(prefix="/api/license", tags=["license"])


def get_license_service() -> LicenseService:
    return service


def _http_error(error: LicenseServiceError) -> HTTPException:
    return HTTPException(
        status_code=400,
        detail={"code": error.code, "message": str(error)},
    )


@router.get("/status", response_model=LicenseStatus)
def get_status(license_service: LicenseService = Depends(get_license_service)):
    return license_service.get_license_status()


@router.post("/trial-heartbeat", response_model=LicenseStatus)
def trial_heartbeat(license_service: LicenseService = Depends(get_license_service)):
    return license_service.trial_heartbeat()


@router.post("/activate-file", response_model=LicenseStatus)
async def activate_file(
    file: UploadFile = File(...),
    license_service: LicenseService = Depends(get_license_service),
):
    contents = await file.read(license_service.max_license_file_bytes + 1)
    if len(contents) > license_service.max_license_file_bytes:
        raise HTTPException(
            status_code=413,
            detail={"code": "LICENSE_FILE_TOO_LARGE", "message": "The selected license file is too large."},
        )
    try:
        return license_service.activate_license_file(contents)
    except LicenseServiceError as error:
        raise _http_error(error) from error


@router.post("/activate-key", response_model=LicenseStatus)
def activate_key(
    request: LicenseKeyRequest,
    license_service: LicenseService = Depends(get_license_service),
):
    try:
        return license_service.activate_license_key(request.key)
    except LicenseServiceError as error:
        raise _http_error(error) from error
