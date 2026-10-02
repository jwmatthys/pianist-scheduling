from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from .. import models, schemas
from ..database import get_db
from ..services.module_lifecycle import (
    bump_accompanist_revision,
    lesson_response,
    refresh_student_identity_state,
    set_lesson_jury_required,
    synchronize_lesson_identity,
)
from ..services.jury_sync import sync_jury_with_accompanist
from .. import module_models

router = APIRouter(prefix="/api/lessons", tags=["lessons"])


@router.get("", response_model=list[schemas.LessonOut])
def list_lessons(db: Session = Depends(get_db)):
    lessons = db.query(models.Lesson).order_by(models.Lesson.day, models.Lesson.start_minute).all()
    return [lesson_response(db, lesson) for lesson in lessons]


@router.post("", response_model=schemas.LessonOut)
def create_lesson(payload: schemas.LessonCreate, db: Session = Depends(get_db)):
    values = payload.model_dump()
    jury_required = values.pop("jury_required", False)
    lesson = models.Lesson(**values)
    db.add(lesson)
    db.flush()
    synchronize_lesson_identity(db, lesson, identity_fields_changed=True)
    set_lesson_jury_required(db, lesson.id, jury_required)
    bump_accompanist_revision(db)
    sync_jury_with_accompanist(db)
    db.commit()
    db.refresh(lesson)
    return lesson_response(db, lesson)


@router.patch("/{lesson_id}", response_model=schemas.LessonOut)
def update_lesson(lesson_id: int, payload: schemas.LessonUpdate, db: Session = Depends(get_db)):
    lesson = db.get(models.Lesson, lesson_id)
    if not lesson:
        raise HTTPException(404, "Lesson not found")
    data = payload.model_dump(exclude_unset=True)
    clear = data.pop("clear_assigned_pianist", False)
    jury_required = data.pop("jury_required", None)
    changed = any(getattr(lesson, key) != value for key, value in data.items())
    if clear and lesson.assigned_pianist_id is not None:
        changed = True
    for key, value in data.items():
        setattr(lesson, key, value)
    if clear:
        lesson.assigned_pianist_id = None
    if "assigned_pianist_id" in data or clear:
        lesson.manually_edited = True
        lesson.fit_quality = "Manual"
    identity_changed = any(key in data for key in ("student", "student_id"))
    if identity_changed:
        db.flush()
        synchronize_lesson_identity(db, lesson, identity_fields_changed=True)
    if jury_required is not None:
        changed = set_lesson_jury_required(db, lesson.id, jury_required) or changed
    if changed:
        bump_accompanist_revision(db)
        sync_jury_with_accompanist(db)
    db.commit()
    db.refresh(lesson)
    return lesson_response(db, lesson)


@router.delete("/{lesson_id}")
def delete_lesson(lesson_id: int, db: Session = Depends(get_db)):
    lesson = db.get(models.Lesson, lesson_id)
    if not lesson:
        raise HTTPException(404, "Lesson not found")
    identity = db.get(module_models.AccompanistLessonIdentity, lesson.id)
    student_person_uuid = identity.student_person_uuid if identity is not None else None
    if identity is not None:
        db.delete(identity)
    jury_requirement = db.get(module_models.AccompanistLessonJuryRequirement, lesson.id)
    if jury_requirement is not None:
        db.delete(jury_requirement)
    db.delete(lesson)
    db.flush()
    if student_person_uuid:
        refresh_student_identity_state(db, student_person_uuid)
    bump_accompanist_revision(db)
    sync_jury_with_accompanist(db)
    db.commit()
    return {"ok": True}


@router.delete("")
def delete_all_lessons(db: Session = Depends(get_db)):
    lesson_ids = [lesson.id for lesson in db.query(models.Lesson.id).all()]
    if lesson_ids:
        db.query(module_models.AccompanistLessonIdentity).filter(
            module_models.AccompanistLessonIdentity.lesson_id.in_(lesson_ids)
        ).delete(synchronize_session=False)
        db.query(module_models.AccompanistLessonJuryRequirement).filter(
            module_models.AccompanistLessonJuryRequirement.lesson_id.in_(lesson_ids)
        ).delete(synchronize_session=False)
    db.query(models.Lesson).delete()
    if lesson_ids:
        bump_accompanist_revision(db)
        sync_jury_with_accompanist(db)
    db.commit()
    return {"ok": True}
