import json

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from sqlalchemy.orm import Session

from .. import models, schemas
from ..database import get_db
from ..services import importer

router = APIRouter(prefix="/api/import", tags=["import"])


@router.post("/preview", response_model=schemas.ImportPreview)
async def preview(file: UploadFile = File(...)):
    content = await file.read()
    try:
        token, columns, rows = importer.stage_upload(file.filename, content)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(400, f"Could not read file: {exc}") from exc
    return schemas.ImportPreview(columns=columns, rows=rows, upload_token=token)


@router.get("/target-fields")
def target_fields():
    return {"fields": importer.TARGET_FIELDS}


@router.post("/commit", response_model=schemas.ImportCommitResult)
def commit(payload: schemas.ImportCommit, db: Session = Depends(get_db)):
    try:
        lesson_dicts, warnings = importer.commit_upload(payload.upload_token, payload.mapping)
    except KeyError as exc:
        raise HTTPException(400, str(exc)) from exc

    db.query(models.Lesson).delete()
    for d in lesson_dicts:
        db.add(models.Lesson(**d))

    if payload.save_profile_name:
        existing = (
            db.query(models.ImportProfile)
            .filter(models.ImportProfile.name == payload.save_profile_name)
            .first()
        )
        mapping_json = json.dumps(payload.mapping)
        if existing:
            existing.mapping_json = mapping_json
        else:
            db.add(models.ImportProfile(name=payload.save_profile_name, mapping_json=mapping_json))

    db.commit()
    return schemas.ImportCommitResult(created=len(lesson_dicts), skipped=len(warnings), warnings=warnings)


@router.get("/profiles")
def list_profiles(db: Session = Depends(get_db)):
    profiles = db.query(models.ImportProfile).all()
    return [
        {"id": p.id, "name": p.name, "mapping": json.loads(p.mapping_json)}
        for p in profiles
    ]
