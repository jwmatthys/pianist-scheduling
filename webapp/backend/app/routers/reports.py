from fastapi import APIRouter, Depends
from fastapi.responses import PlainTextResponse
from sqlalchemy.orm import Session

from .. import models
from ..database import get_db
from ..services import reports

router = APIRouter(prefix="/api/reports", tags=["reports"])


@router.get("/markdown", response_class=PlainTextResponse)
def markdown_report(db: Session = Depends(get_db)):
    lessons = db.query(models.Lesson).all()
    pianists = db.query(models.Pianist).all()
    pianists_by_id = {p.id: p for p in pianists}
    return reports.build_markdown(lessons, pianists_by_id, source_name="Pianist Scheduling Webapp")
