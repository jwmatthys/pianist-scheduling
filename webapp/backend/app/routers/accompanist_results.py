from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from .. import jury_schemas
from ..database import get_db
from ..services import accompanist_results
from ..services.accompanist_results import ResultPublicationError
from ..services.module_lifecycle import module_revision
from .. import module_models

router = APIRouter(tags=["module results"])


def _result_error(error: ResultPublicationError) -> HTTPException:
    status = 409 if error.code in {"SOURCE_REVISION_CHANGED", "STUDENT_IDENTITY_CONFLICT"} else 422
    return HTTPException(status_code=status, detail={"code": error.code, "message": str(error)})


@router.post("/api/accompanist/finalize", response_model=jury_schemas.ResultEnvelope)
def finalize_accompanist(
    payload: jury_schemas.FinalizeAccompanistRequest,
    db: Session = Depends(get_db),
):
    try:
        result = accompanist_results.finalize_accompanist_result(
            db,
            payload.expected_source_revision,
        )
        db.commit()
        return result
    except ResultPublicationError as error:
        db.rollback()
        raise _result_error(error) from error
    except Exception:
        db.rollback()
        raise


@router.get("/api/results/accompanist/finalized", response_model=jury_schemas.ResultEnvelope)
def current_finalized_accompanist(db: Session = Depends(get_db)):
    result = accompanist_results.get_current_accompanist_result(db)
    if result is None:
        raise HTTPException(
            status_code=404,
            detail={"code": "NO_CURRENT_FINALIZED_RESULT", "message": "No current finalized Accompanist result exists."},
        )
    return result


@router.get(
    "/api/accompanist/finalization-state",
    response_model=jury_schemas.AccompanistFinalizationStateOut,
)
def accompanist_finalization_state(db: Session = Depends(get_db)):
    state = module_revision(db)
    result = (
        db.get(module_models.ModuleResult, state.current_result_uuid)
        if state.current_result_uuid else None
    )
    return jury_schemas.AccompanistFinalizationStateOut(
        session_uuid=state.session_uuid,
        source_revision=state.source_revision,
        current_result_uuid=state.current_result_uuid,
        current_result_version=result.result_version if result and result.state == "finalized" else None,
    )