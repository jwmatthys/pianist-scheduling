from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from .. import models, module_models, schemas
from ..database import get_db
from ..services.module_lifecycle import (
    bump_accompanist_revision,
    ensure_pianist_identity,
    remove_pianist_identity_mapping,
)

router = APIRouter(prefix="/api/pianists", tags=["pianists"])


@router.get("", response_model=list[schemas.PianistOut])
def list_pianists(db: Session = Depends(get_db)):
    return db.query(models.Pianist).order_by(models.Pianist.name).all()


@router.post("", response_model=schemas.PianistOut)
def create_pianist(payload: schemas.PianistCreate, db: Session = Depends(get_db)):
    pianist = models.Pianist(**payload.model_dump())
    db.add(pianist)
    db.flush()
    ensure_pianist_identity(db, pianist)
    db.commit()
    db.refresh(pianist)
    return pianist


@router.patch("/{pianist_id}", response_model=schemas.PianistOut)
def update_pianist(pianist_id: int, payload: schemas.PianistUpdate, db: Session = Depends(get_db)):
    pianist = db.get(models.Pianist, pianist_id)
    if not pianist:
        raise HTTPException(404, "Pianist not found")
    data = payload.model_dump(exclude_unset=True)
    name_changed = "name" in data and pianist.name != data["name"]
    for key, value in data.items():
        setattr(pianist, key, value)
    ensure_pianist_identity(db, pianist)
    assigned = name_changed and db.query(models.Lesson.id).filter(
        models.Lesson.assigned_pianist_id == pianist_id
    ).first() is not None
    if assigned:
        bump_accompanist_revision(db)
    db.commit()
    db.refresh(pianist)
    return pianist


@router.delete("/{pianist_id}")
def delete_pianist(pianist_id: int, db: Session = Depends(get_db)):
    pianist = db.get(models.Pianist, pianist_id)
    if not pianist:
        raise HTTPException(404, "Pianist not found")
    assigned = db.query(models.Lesson.id).filter(
        models.Lesson.assigned_pianist_id == pianist_id
    ).first() is not None
    remove_pianist_identity_mapping(db, pianist_id)
    db.delete(pianist)
    if assigned:
        bump_accompanist_revision(db)
    db.commit()
    return {"ok": True}


@router.get("/{pianist_id}/availability", response_model=list[schemas.AvailabilitySlotOut])
def get_availability(pianist_id: int, db: Session = Depends(get_db)):
    return (
        db.query(models.AvailabilitySlot)
        .filter(models.AvailabilitySlot.pianist_id == pianist_id)
        .all()
    )


@router.put("/{pianist_id}/availability", response_model=list[schemas.AvailabilitySlotOut])
def set_availability(pianist_id: int, payload: schemas.AvailabilityBulkIn, db: Session = Depends(get_db)):
    """Replaces the full availability grid for a pianist in one call."""
    pianist = db.get(models.Pianist, pianist_id)
    if not pianist:
        raise HTTPException(404, "Pianist not found")

    db.query(models.AvailabilitySlot).filter(
        models.AvailabilitySlot.pianist_id == pianist_id
    ).delete()

    availability_state = db.get(models.PianistAvailabilityState, pianist_id)
    if availability_state is None:
        db.add(models.PianistAvailabilityState(pianist_id=pianist_id, is_complete=True))
    else:
        availability_state.is_complete = True

    slots = [
        models.AvailabilitySlot(
            pianist_id=pianist_id,
            day=s.day,
            slot_start_minute=s.slot_start_minute,
            status=s.status,
        )
        for s in payload.slots
        if s.status in ("Available", "Tentative", "Unavailable")
    ]
    db.add_all(slots)
    db.commit()
    return db.query(models.AvailabilitySlot).filter(
        models.AvailabilitySlot.pianist_id == pianist_id
    ).all()
