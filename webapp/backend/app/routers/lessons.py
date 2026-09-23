from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from .. import models, schemas
from ..database import get_db

router = APIRouter(prefix="/api/lessons", tags=["lessons"])


@router.get("", response_model=list[schemas.LessonOut])
def list_lessons(db: Session = Depends(get_db)):
    return db.query(models.Lesson).order_by(models.Lesson.day, models.Lesson.start_minute).all()


@router.post("", response_model=schemas.LessonOut)
def create_lesson(payload: schemas.LessonCreate, db: Session = Depends(get_db)):
    lesson = models.Lesson(**payload.model_dump())
    db.add(lesson)
    db.commit()
    db.refresh(lesson)
    return lesson


@router.patch("/{lesson_id}", response_model=schemas.LessonOut)
def update_lesson(lesson_id: int, payload: schemas.LessonUpdate, db: Session = Depends(get_db)):
    lesson = db.get(models.Lesson, lesson_id)
    if not lesson:
        raise HTTPException(404, "Lesson not found")
    data = payload.model_dump(exclude_unset=True)
    clear = data.pop("clear_assigned_pianist", False)
    for key, value in data.items():
        setattr(lesson, key, value)
    if clear:
        lesson.assigned_pianist_id = None
    if "assigned_pianist_id" in data or clear:
        lesson.manually_edited = True
        lesson.fit_quality = lesson.fit_quality or "Manual"
    db.commit()
    db.refresh(lesson)
    return lesson


@router.delete("/{lesson_id}")
def delete_lesson(lesson_id: int, db: Session = Depends(get_db)):
    lesson = db.get(models.Lesson, lesson_id)
    if not lesson:
        raise HTTPException(404, "Lesson not found")
    db.delete(lesson)
    db.commit()
    return {"ok": True}


@router.delete("")
def delete_all_lessons(db: Session = Depends(get_db)):
    db.query(models.Lesson).delete()
    db.commit()
    return {"ok": True}
