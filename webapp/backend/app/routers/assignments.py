from fastapi import APIRouter, Depends
from sqlalchemy.orm import Session

from .. import models, schemas
from ..database import get_db
from ..services import scheduling

router = APIRouter(prefix="/api/assignments", tags=["assignments"])


def _build_engine_state(db: Session):
    pianists = db.query(models.Pianist).all()
    lessons = db.query(models.Lesson).order_by(models.Lesson.day, models.Lesson.start_minute).all()

    engine_pianists = []
    for p in pianists:
        avail: dict[str, dict[int, str]] = {}
        for slot in p.availability:
            avail.setdefault(slot.day, {})[slot.slot_start_minute] = slot.status
        engine_pianists.append(
            scheduling.EnginePianist(id=p.id, name=p.name, max_hours=p.max_hours_per_week, avail=avail)
        )

    engine_lessons = [
        scheduling.EngineLesson(
            id=l.id,
            day=l.day,
            start_min=l.start_minute,
            end_min=l.end_minute,
            required_pianist_name=l.required_pianist_name,
            need_pianist=l.need_pianist,
            assigned_pianist_id=l.assigned_pianist_id,
            fit_quality=l.fit_quality,
            notes=l.notes,
            hours=l.hours,
        )
        for l in lessons
    ]
    return lessons, pianists, engine_lessons, engine_pianists


def _persist_and_respond(db: Session, lessons, engine_lessons, hours_by_name, conflicts):
    by_id = {l.id: l for l in engine_lessons}
    for lesson in lessons:
        el = by_id[lesson.id]
        lesson.assigned_pianist_id = el.assigned_pianist_id
        lesson.fit_quality = el.fit_quality
        lesson.notes = el.notes
        lesson.hours = el.hours
    db.commit()
    for lesson in lessons:
        db.refresh(lesson)

    unassigned_count = sum(1 for l in lessons if l.need_pianist and l.assigned_pianist_id is None)
    return schemas.RunAssignmentResult(
        lessons=[schemas.LessonOut.model_validate(l) for l in lessons],
        hours_by_pianist=hours_by_name,
        conflicts=conflicts,
        unassigned_count=unassigned_count,
    )


@router.post("/run", response_model=schemas.RunAssignmentResult)
def run_assignment(db: Session = Depends(get_db)):
    """Runs the best-fit algorithm over ALL lessons, overwriting prior
    (non-manually-edited) assignments. Manually edited lessons are left as-is
    but still count toward conflict/workload calculations.
    """
    lessons, pianists, engine_lessons, engine_pianists = _build_engine_state(db)

    manual_ids = {l.id for l in lessons if l.manually_edited}
    hours_by_name, conflicts = scheduling.assign_lessons(engine_lessons, engine_pianists, locked_ids=manual_ids)

    return _persist_and_respond(db, lessons, engine_lessons, hours_by_name, conflicts)


@router.get("/validate", response_model=schemas.ValidationResult)
def validate(db: Session = Depends(get_db)):
    """Recomputes hours/conflicts without re-running the assignment algorithm
    (used after manual edits in the UI)."""
    lessons, pianists, engine_lessons, engine_pianists = _build_engine_state(db)
    hours_by_name, conflicts = scheduling.recompute_hours_and_conflicts(engine_lessons, engine_pianists)

    by_id = {l.id: l for l in engine_lessons}
    for lesson in lessons:
        el = by_id[lesson.id]
        lesson.hours = el.hours
        lesson.notes = el.notes
    db.commit()
    for lesson in lessons:
        db.refresh(lesson)

    unassigned_count = sum(1 for l in lessons if l.need_pianist and l.assigned_pianist_id is None)
    return schemas.ValidationResult(
        lessons=[schemas.LessonOut.model_validate(l) for l in lessons],
        hours_by_pianist=hours_by_name,
        conflicts=conflicts,
        unassigned_count=unassigned_count,
    )
