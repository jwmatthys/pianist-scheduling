"""Reconcile Jury-owned references against current Accompanist identities."""

import logging
from uuid import UUID

from sqlalchemy.orm import Session

from .. import module_models
from ..jury_schemas import JurySynchronizationSummary
from . import accompanist_results, jury
from .module_lifecycle import active_session_uuid, bump_jury_revision


logger = logging.getLogger(__name__)
JURY_MODULE_ID = "juries"


def _stale_result_count(
    db: Session,
    session_uuid: str,
    source_revision: int,
    current_result_uuid: str | None,
) -> int:
    results = db.query(module_models.ModuleResult).filter(
        module_models.ModuleResult.session_uuid == session_uuid,
        module_models.ModuleResult.module_id == JURY_MODULE_ID,
    ).all()
    stale_count = 0
    for result in results:
        dependency = db.get(
            module_models.ModuleResultDependency,
            (result.result_uuid, accompanist_results.ACCOMPANIST_RESULT_CONTRACT),
        )
        if (
            dependency is None
            or dependency.source_session_uuid != session_uuid
            or dependency.source_revision != source_revision
            or current_result_uuid is None
            or dependency.source_result_uuid != current_result_uuid
        ):
            stale_count += 1
    return stale_count


def sync_jury_with_accompanist(db: Session) -> JurySynchronizationSummary:
    """Prune orphaned references, refresh a current finalized roster, and report stale results."""
    session_uuid = active_session_uuid(db)
    source = accompanist_results.get_current_source_identity_snapshot(db)
    lesson_identities = {str(lesson_uuid): str(student_uuid) for lesson_uuid, student_uuid in source.lesson_identities}
    current_pianist_uuids = {str(person_uuid) for person_uuid in source.pianist_person_uuids}

    lesson_rows = db.query(module_models.JuryLessonEntry).filter_by(session_uuid=session_uuid).all()
    lesson_entries_removed = 0
    lesson_identity_references_updated = 0
    for entry in lesson_rows:
        current_student_uuid = lesson_identities.get(entry.source_lesson_uuid)
        if current_student_uuid is None:
            db.delete(entry)
            lesson_entries_removed += 1
        elif entry.student_person_uuid != current_student_uuid:
            entry.student_person_uuid = current_student_uuid
            lesson_identity_references_updated += 1

    declarations = db.query(module_models.JuryPianistAvailabilityDeclaration).filter_by(
        session_uuid=session_uuid
    ).all()
    declaration_keys = {
        (record.pianist_person_uuid, record.jury_date)
        for record in declarations
        if record.pianist_person_uuid in current_pianist_uuids
    }
    orphan_declarations = [
        record for record in declarations
        if record.pianist_person_uuid not in current_pianist_uuids
    ]
    windows = db.query(module_models.JuryPianistAvailableWindow).filter_by(
        session_uuid=session_uuid
    ).all()
    orphan_windows = [
        window for window in windows
        if window.pianist_person_uuid not in current_pianist_uuids
        or (window.pianist_person_uuid, window.jury_date) not in declaration_keys
    ]
    if orphan_windows:
        orphan_window_uuids = [window.window_uuid for window in orphan_windows]
        db.query(module_models.JuryPianistAvailableWindow).filter(
            module_models.JuryPianistAvailableWindow.window_uuid.in_(orphan_window_uuids)
        ).delete(synchronize_session=False)
    for record in orphan_declarations:
        db.delete(record)

    roster_entries_created = 0
    current_result = accompanist_results.get_current_accompanist_result(db)
    if current_result is not None:
        previous_lesson_uuids = {entry.source_lesson_uuid for entry in lesson_rows}
        current_result_entries = jury.sync_roster_from_current_result(db)
        roster_entries_created = sum(
            str(entry.source_lesson_uuid) not in previous_lesson_uuids
            for entry in current_result_entries
        )

    changed = bool(lesson_entries_removed or lesson_identity_references_updated or orphan_declarations or orphan_windows)
    if changed:
        bump_jury_revision(db)

    stale_results_detected = _stale_result_count(
        db,
        session_uuid,
        source.source_revision,
        str(source.current_result_uuid) if source.current_result_uuid else None,
    )
    summary = JurySynchronizationSummary(
        session_uuid=UUID(session_uuid),
        accompanist_source_revision=source.source_revision,
        current_accompanist_result_uuid=source.current_result_uuid,
        lesson_references_checked=len(source.lesson_identities),
        lesson_entries_removed=lesson_entries_removed,
        lesson_identity_references_updated=lesson_identity_references_updated,
        roster_entries_created=roster_entries_created,
        pianist_references_checked=len(source.pianist_person_uuids),
        availability_records_removed=len(orphan_declarations),
        availability_windows_removed=len(orphan_windows),
        stale_results_detected=stale_results_detected,
    )

    if lesson_entries_removed:
        logger.warning(
            "Orphaned Jury lesson references removed",
            extra={"event": "jury_orphaned_lessons_removed", "count": lesson_entries_removed},
        )
    if orphan_declarations or orphan_windows:
        logger.warning(
            "Orphaned Jury pianist availability removed",
            extra={
                "event": "jury_orphaned_pianist_availability_removed",
                "records_removed": len(orphan_declarations),
                "windows_removed": len(orphan_windows),
            },
        )
    if stale_results_detected:
        logger.warning(
            "Jury results reference an older Accompanist dependency",
            extra={"event": "jury_stale_results_detected", "count": stale_results_detected},
        )
    logger.info(
        "Jury references synchronized with Accompanist",
        extra={
            "event": "jury_accompanist_sync_completed",
            "lesson_entries_removed": lesson_entries_removed,
            "lesson_identity_references_updated": lesson_identity_references_updated,
            "roster_entries_created": roster_entries_created,
            "availability_records_removed": len(orphan_declarations),
            "availability_windows_removed": len(orphan_windows),
            "stale_results_detected": stale_results_detected,
        },
    )
    return summary