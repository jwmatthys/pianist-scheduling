"""Shared local availability import parsing with an Accompanist apply boundary."""

from __future__ import annotations

import re
import uuid
from dataclasses import dataclass, field

import pandas as pd
from sqlalchemy.orm import Session

from .. import models, module_models, schemas
from . import tabular
from .accompanist_availability import AccompanistAvailabilitySlot, windows_to_accompanist_slots
from .availability import (
    AvailabilityIssue,
    AvailabilityWindow,
    normalize_day,
    normalize_status,
    normalize_windows,
    parse_time_minutes,
)
from .module_lifecycle import (
    active_session_uuid,
    bump_accompanist_revision,
    bump_jury_revision,
    ensure_pianist_identity,
    remove_pianist_identity_mapping,
)

MAX_UPLOAD_BYTES = 25 * 1024 * 1024
MAX_STAGED_UPLOADS = 20
MAX_PREVIEW_WINDOWS = 300
_STAGED_UPLOADS: dict[str, StagedUpload] = {}
_STAGED_PREVIEWS: dict[str, PreparedImport] = {}


class AvailabilityImportError(RuntimeError):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


@dataclass(frozen=True)
class StagedUpload:
    filename: str
    content: bytes
    sheets: tuple[str, ...]


@dataclass
class PianistImportPlan:
    key: str
    pianist_name: str
    email: str
    row_numbers: set[int] = field(default_factory=set)
    days: set[str] = field(default_factory=set)


@dataclass
class PreparedImport:
    upload_token: str
    sheet_name: str | None
    pianists: dict[str, PianistImportPlan] = field(default_factory=dict)
    review_pianists: list[schemas.AvailabilityImportPianistOut] = field(default_factory=list)
    windows_by_pianist: dict[str, list[AvailabilityWindow]] = field(default_factory=dict)
    slots_by_pianist: dict[str, list[AccompanistAvailabilitySlot]] = field(default_factory=dict)
    errors: list[AvailabilityIssue] = field(default_factory=list)
    warnings: list[AvailabilityIssue] = field(default_factory=list)
    rows_processed: int = 0
    existing_pianist_ids: set[int] = field(default_factory=set)
    existing_assignment_count: int = 0
    existing_slots_in_scope: int = 0


def _is_blank(value: object) -> bool:
    if value is None or isinstance(value, str) and not value.strip():
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _frame_for(upload: StagedUpload, sheet_name: str | None) -> pd.DataFrame:
    if upload.sheets:
        sheet_name = sheet_name or upload.sheets[0]
        if sheet_name not in upload.sheets:
            raise AvailabilityImportError("UNKNOWN_SHEET", "Choose a worksheet that exists in the selected file.")
    elif sheet_name:
        raise AvailabilityImportError("UNKNOWN_SHEET", "CSV files do not contain worksheets.")
    frame = tabular.read_table(upload.filename, upload.content, sheet_name)
    frame = frame.copy()
    frame.columns = [str(column) for column in frame.columns]
    return frame


def _sample_rows(frame: pd.DataFrame) -> list[dict[str, str]]:
    return frame.head(15).fillna("").astype(str).to_dict(orient="records")


def _suggest_column(columns: list[str], aliases: tuple[str, ...]) -> str | None:
    normalized = {column: re.sub(r"[^a-z0-9]+", " ", column.casefold()).strip() for column in columns}
    for alias in aliases:
        for column, header in normalized.items():
            if header == alias:
                return column
    for alias in aliases:
        for column, header in normalized.items():
            if alias in header.split():
                return column
    return None


def _suggest_wide_windows(columns: list[str]) -> dict[str, list[schemas.AvailabilityWindowColumns]]:
    starts: dict[tuple[str, int], str] = {}
    ends: dict[tuple[str, int], str] = {}
    for column in columns:
        tokens = re.findall(r"[a-z]+|\d+", column.casefold())
        day = next((normalize_day(token) for token in tokens if normalize_day(token)), None)
        if not day:
            continue
        kind = None
        if any(token in {"start", "from", "begin", "begins"} for token in tokens):
            kind = starts
        elif any(token in {"end", "to", "until", "finish"} for token in tokens):
            kind = ends
        if kind is None:
            continue
        numbers = [int(token) for token in tokens if token.isdigit()]
        window_index = numbers[-1] if numbers else 1
        kind[(day, window_index)] = column

    result: dict[str, list[schemas.AvailabilityWindowColumns]] = {}
    for day, window_index in sorted(set(starts) | set(ends), key=lambda item: (models.DAYS_ORDER.index(item[0]), item[1])):
        result.setdefault(day, []).append(schemas.AvailabilityWindowColumns(
            start_column=starts.get((day, window_index)),
            end_column=ends.get((day, window_index)),
        ))
    return result


def _inspection(upload_token: str, upload: StagedUpload, sheet_name: str | None) -> schemas.AvailabilityImportInspection:
    frame = _frame_for(upload, sheet_name)
    columns = [str(column) for column in frame.columns]
    suggested = {
        "person_name_column": _suggest_column(columns, ("pianist name", "person name", "full name", "name", "pianist", "respondent")),
        "email_column": _suggest_column(columns, ("pianist email", "email address", "email")),
        "day_column": _suggest_column(columns, ("day", "weekday", "day of week")),
        "start_column": _suggest_column(columns, ("start time", "start", "available from")),
        "end_column": _suggest_column(columns, ("end time", "end", "available until")),
        "status_column": _suggest_column(columns, ("availability status", "status", "availability")),
    }
    return schemas.AvailabilityImportInspection(
        upload_token=upload_token,
        sheets=list(upload.sheets),
        selected_sheet=sheet_name,
        columns=columns,
        sample_rows=_sample_rows(frame),
        suggested_normalized=suggested,
        suggested_wide_windows=_suggest_wide_windows(columns),
    )


def inspect_upload(filename: str, content: bytes) -> schemas.AvailabilityImportInspection:
    if len(content) > MAX_UPLOAD_BYTES:
        raise AvailabilityImportError("FILE_TOO_LARGE", "Availability files must be 25 MB or smaller.")
    try:
        sheets = tuple(tabular.list_sheets(filename, content))
        if sheets == () and not filename.casefold().endswith(".csv"):
            raise AvailabilityImportError("EMPTY_WORKBOOK", "The workbook does not contain any worksheets.")
    except AvailabilityImportError:
        raise
    except Exception as error:
        raise AvailabilityImportError("FILE_READ_FAILED", f"Could not read the selected file: {error}") from error

    upload_token = uuid.uuid4().hex
    upload = StagedUpload(filename, content, sheets)
    _STAGED_UPLOADS[upload_token] = upload
    while len(_STAGED_UPLOADS) > MAX_STAGED_UPLOADS:
        _STAGED_UPLOADS.pop(next(iter(_STAGED_UPLOADS)))
    return _inspection(upload_token, upload, sheets[0] if sheets else None)


def inspect_sheet(upload_token: str, sheet_name: str | None) -> schemas.AvailabilityImportInspection:
    upload = _STAGED_UPLOADS.get(upload_token)
    if upload is None:
        raise AvailabilityImportError("UPLOAD_EXPIRED", "The staged file expired; select it again to continue.")
    try:
        return _inspection(upload_token, upload, sheet_name)
    except AvailabilityImportError:
        raise
    except Exception as error:
        raise AvailabilityImportError("SHEET_READ_FAILED", f"Could not read the selected worksheet: {error}") from error


def _issue(severity: str, code: str, message: str, row_number: int | None = None) -> AvailabilityIssue:
    return AvailabilityIssue(severity, code, message, row_number)  # type: ignore[arg-type]


def _mapping_columns(
    mapping: schemas.AvailabilityImportMapping,
    columns: list[str],
) -> list[AvailabilityIssue]:
    issues: list[AvailabilityIssue] = []

    def check(column: str | None, label: str, required: bool) -> None:
        if not column:
            if required:
                issues.append(_issue("error", "MISSING_MAPPING", f"Map the {label} column."))
        elif column not in columns:
            issues.append(_issue("error", "INVALID_MAPPING", f"The mapped {label} column is not in the selected worksheet."))

    check(mapping.person_name_column, "pianist name", True)
    check(mapping.email_column, "email", False)
    if mapping.layout == "normalized":
        check(mapping.day_column, "weekday", True)
        check(mapping.start_column, "start time", True)
        check(mapping.end_column, "end time", True)
        check(mapping.status_column, "availability status", True)
    else:
        mapped_days: set[str] = set()
        for raw_day, pairs in mapping.wide_windows.items():
            day = normalize_day(raw_day)
            if not day:
                issues.append(_issue("error", "INVALID_DAY_MAPPING", f"'{raw_day}' is not a weekday."))
                continue
            mapped_days.add(day)
            if not pairs:
                issues.append(_issue("error", "MISSING_MAPPING", f"Map at least one {day} start/end pair."))
            for pair in pairs:
                check(pair.start_column, f"{day} start time", True)
                check(pair.end_column, f"{day} end time", True)
        for day in models.DAYS_ORDER:
            if day not in mapped_days:
                issues.append(_issue(
                    "error", "MISSING_MAPPING",
                    f"Map at least one {day} start/end pair; blank mapped cells mean Unavailable.",
                ))
    return issues


def _parse_row_window(
    *,
    row: pd.Series,
    row_number: int,
    day_value: object,
    start_value: object,
    end_value: object,
    status_value: object,
    errors: list[AvailabilityIssue],
) -> AvailabilityWindow | None:
    day = normalize_day(day_value)
    if not day:
        errors.append(_issue("error", "INVALID_DAY", f"Unrecognized weekday '{day_value}'.", row_number))
        return None
    start = parse_time_minutes(start_value)
    end = parse_time_minutes(end_value, allow_end_of_day=True)
    if start is None:
        errors.append(_issue("error", "INVALID_START_TIME", f"Invalid start time '{start_value}'.", row_number))
    if end is None:
        errors.append(_issue("error", "INVALID_END_TIME", f"Invalid end time '{end_value}'.", row_number))
    if start is None or end is None:
        return None
    if end <= start:
        errors.append(_issue("error", "INVALID_TIME_RANGE", "End time must be later than start time.", row_number))
        return None
    status = normalize_status(status_value)
    if status is None:
        errors.append(_issue("error", "INVALID_STATUS", f"Unrecognized availability status '{status_value}'.", row_number))
        return None
    return AvailabilityWindow(day, start, end, status, source="import", source_row=row_number)


def _person_for_row(
    row: pd.Series,
    row_number: int,
    mapping: schemas.AvailabilityImportMapping,
    prepared: PreparedImport,
) -> PianistImportPlan | None:
    name_value = row.get(mapping.person_name_column) if mapping.person_name_column else None
    if _is_blank(name_value):
        prepared.errors.append(_issue("error", "MISSING_PERSON", "Pianist name is missing.", row_number))
        prepared.review_pianists.append(schemas.AvailabilityImportPianistOut(
            action="invalid", pianist_name="", days=[], row_numbers=[row_number],
        ))
        return None
    name = " ".join(str(name_value).split())
    if len(name) > 200:
        prepared.errors.append(_issue("error", "INVALID_PERSON_NAME", "Pianist name must be 200 characters or fewer.", row_number))
        prepared.review_pianists.append(schemas.AvailabilityImportPianistOut(
            action="invalid", pianist_name=name[:200], days=[], row_numbers=[row_number],
        ))
        return None

    email_value = row.get(mapping.email_column) if mapping.email_column else None
    imported_email = "" if _is_blank(email_value) else str(email_value).strip()
    if len(imported_email) > 200:
        prepared.errors.append(_issue("error", "INVALID_EMAIL", "Email must be 200 characters or fewer.", row_number))
        prepared.review_pianists.append(schemas.AvailabilityImportPianistOut(
            action="invalid", pianist_name=name, days=[], row_numbers=[row_number],
        ))
        return None

    identity_key = name
    plan = prepared.pianists.get(identity_key)
    if plan is None:
        plan = PianistImportPlan(key=identity_key, pianist_name=name, email=imported_email)
        prepared.pianists[identity_key] = plan
    elif imported_email and not plan.email:
        plan.email = imported_email

    plan.row_numbers.add(row_number)
    prepared.windows_by_pianist.setdefault(identity_key, [])
    return plan


def preview_import(
    request: schemas.AvailabilityImportPreviewRequest,
    db: Session,
) -> schemas.AvailabilityImportPreviewOut:
    upload = _STAGED_UPLOADS.get(request.upload_token)
    if upload is None:
        raise AvailabilityImportError("UPLOAD_EXPIRED", "The staged file expired; select it again to continue.")
    try:
        frame = _frame_for(upload, request.sheet_name)
    except Exception as error:
        if isinstance(error, AvailabilityImportError):
            raise
        raise AvailabilityImportError("SHEET_READ_FAILED", f"Could not read the selected worksheet: {error}") from error

    frame = frame.dropna(how="all")
    columns = [str(column) for column in frame.columns]
    prepared = PreparedImport(
        upload_token=request.upload_token,
        sheet_name=request.sheet_name or (upload.sheets[0] if upload.sheets else None),
        rows_processed=len(frame),
    )
    prepared.errors.extend(_mapping_columns(request.mapping, columns))
    mapping = request.mapping
    if not prepared.errors:
        for index, row in frame.iterrows():
            row_number = int(index) + 2 if isinstance(index, (int, float)) else 0
            if all(_is_blank(value) for value in row.values):
                continue
            pianist = _person_for_row(
                row,
                row_number,
                mapping,
                prepared,
            )
            if pianist is None:
                continue

            row_windows: list[AvailabilityWindow] = []
            if mapping.layout == "normalized":
                values = [
                    row.get(mapping.day_column),
                    row.get(mapping.start_column),
                    row.get(mapping.end_column),
                    row.get(mapping.status_column),
                ]
                if all(_is_blank(value) for value in values):
                    continue
                window = _parse_row_window(
                    row=row,
                    row_number=row_number,
                    day_value=values[0],
                    start_value=values[1],
                    end_value=values[2],
                    status_value=values[3],
                    errors=prepared.errors,
                )
                if window:
                    row_windows.append(window)
            else:
                for raw_day, pairs in mapping.wide_windows.items():
                    day = normalize_day(raw_day)
                    if not day:
                        continue
                    for pair in pairs:
                        if not pair.start_column or not pair.end_column:
                            continue
                        start_value = row.get(pair.start_column)
                        end_value = row.get(pair.end_column)
                        if _is_blank(start_value) and _is_blank(end_value):
                            continue
                        if _is_blank(start_value) or _is_blank(end_value):
                            prepared.errors.append(_issue(
                                "error", "MISSING_TIME_ENDPOINT",
                                f"{day} needs both a start and end time, or both cells left blank.", row_number,
                            ))
                            continue
                        window = _parse_row_window(
                            row=row,
                            row_number=row_number,
                            day_value=day,
                            start_value=start_value,
                            end_value=end_value,
                            status_value=mapping.wide_status,
                            errors=prepared.errors,
                        )
                        if window:
                            row_windows.append(window)
            prepared.windows_by_pianist[pianist.key].extend(row_windows)
            pianist.days.update(window.day for window in row_windows)

    for pianist_key, windows in prepared.windows_by_pianist.items():
        normalized, issues = normalize_windows(windows)
        prepared.windows_by_pianist[pianist_key] = normalized
        prepared.errors.extend(issue for issue in issues if issue.severity == "error")
        prepared.warnings.extend(issue for issue in issues if issue.severity == "warning")
        slots, adapter_issues = windows_to_accompanist_slots(normalized)
        prepared.slots_by_pianist[pianist_key] = slots
        prepared.errors.extend(issue for issue in adapter_issues if issue.severity == "error")

    existing_pianists = db.query(models.Pianist).all()
    prepared.existing_pianist_ids = {pianist.id for pianist in existing_pianists}
    if prepared.existing_pianist_ids:
        prepared.existing_slots_in_scope = db.query(models.AvailabilitySlot).filter(
            models.AvailabilitySlot.pianist_id.in_(prepared.existing_pianist_ids)
        ).count()
        prepared.existing_assignment_count = db.query(models.Lesson).filter(
            models.Lesson.assigned_pianist_id.in_(prepared.existing_pianist_ids)
        ).count()
    prepared.warnings.append(_issue(
        "warning", "FULL_ROSTER_REPLACEMENT",
        "Applying this import replaces the entire Pianist roster and Accompanist weekly availability, clears Lesson assignments, and removes all Jury Availability Windows.",
    ))
    preview_token = uuid.uuid4().hex
    _STAGED_PREVIEWS[preview_token] = prepared
    while len(_STAGED_PREVIEWS) > MAX_STAGED_UPLOADS:
        _STAGED_PREVIEWS.pop(next(iter(_STAGED_PREVIEWS)))

    preview_pianists = [
        schemas.AvailabilityImportPianistOut(
            action="new",
            pianist_name=plan.pianist_name,
            email=plan.email,
            max_hours_per_week=40,
            days=sorted(plan.days, key=models.DAYS_ORDER.index),
            row_numbers=sorted(plan.row_numbers),
        )
        for plan in prepared.pianists.values()
    ] + prepared.review_pianists
    preview_windows = [
        schemas.AvailabilityImportWindowOut(
            pianist_name=prepared.pianists[pianist_key].pianist_name,
            day=window.day,
            start_minute=window.start_minute,
            end_minute=window.end_minute,
            status=window.status,
        )
        for pianist_key, windows in prepared.windows_by_pianist.items()
        for window in windows
    ][:MAX_PREVIEW_WINDOWS]
    can_apply = not prepared.errors and bool(prepared.pianists)
    return schemas.AvailabilityImportPreviewOut(
        preview_token=preview_token,
        sheet_name=request.sheet_name,
        rows_processed=prepared.rows_processed,
        existing_pianist_count=len(prepared.existing_pianist_ids),
        existing_assignment_count=prepared.existing_assignment_count,
        incoming_pianist_count=len(prepared.pianists),
        valid_window_count=sum(len(windows) for windows in prepared.windows_by_pianist.values()),
        existing_slots_in_scope=prepared.existing_slots_in_scope,
        pianists=preview_pianists,
        windows=preview_windows,
        errors=[schemas.AvailabilityImportIssueOut.model_validate(issue.__dict__) for issue in prepared.errors],
        warnings=[schemas.AvailabilityImportIssueOut.model_validate(issue.__dict__) for issue in prepared.warnings],
        can_apply=can_apply,
    )


def apply_import(
    request: schemas.AvailabilityImportApplyRequest,
    db: Session,
) -> schemas.AvailabilityImportApplyResult:
    if not request.confirmed:
        raise AvailabilityImportError("CONFIRMATION_REQUIRED", "Confirm the availability replacement before applying.")
    prepared = _STAGED_PREVIEWS.get(request.preview_token)
    if prepared is None:
        raise AvailabilityImportError("PREVIEW_EXPIRED", "The availability preview expired; review the file again.")
    if prepared.errors:
        raise AvailabilityImportError("VALIDATION_ERRORS", "Resolve all validation errors before applying availability.")
    if not prepared.pianists:
        raise AvailabilityImportError("NO_PIANISTS", "The selected file contains no valid Pianist respondents.")
    try:
        current_pianists = db.query(models.Pianist).all()
        current_ids = {pianist.id for pianist in current_pianists}
        if current_ids != prepared.existing_pianist_ids:
            raise AvailabilityImportError(
                "PIANISTS_CHANGED", "The Pianist roster changed after preview; review the replacement import again."
            )

        session_uuid = active_session_uuid(db)
        jury_availability_windows_removed = db.query(
            module_models.JuryPianistAvailableWindow
        ).filter_by(session_uuid=session_uuid).delete(synchronize_session=False)
        jury_availability_records_removed = db.query(
            module_models.JuryPianistAvailabilityDeclaration
        ).filter_by(session_uuid=session_uuid).delete(synchronize_session=False)
        if jury_availability_windows_removed or jury_availability_records_removed:
            bump_jury_revision(db)

        assignments_cleared = 0
        slots_replaced = 0
        if current_ids:
            assignments_cleared = db.query(models.Lesson).filter(
                models.Lesson.assigned_pianist_id.in_(current_ids)
            ).update({models.Lesson.assigned_pianist_id: None}, synchronize_session=False)
            slots_replaced = db.query(models.AvailabilitySlot).filter(
                models.AvailabilitySlot.pianist_id.in_(current_ids)
            ).delete(synchronize_session=False)
            db.query(models.PianistAvailabilityState).filter(
                models.PianistAvailabilityState.pianist_id.in_(current_ids)
            ).delete(synchronize_session=False)
            for pianist in current_pianists:
                remove_pianist_identity_mapping(db, pianist.id)
                db.delete(pianist)
            db.flush()

        pianist_ids_by_key: dict[str, int] = {}
        created_count = 0
        for key, plan in prepared.pianists.items():
            pianist = models.Pianist(
                organization_id=1,
                name=plan.pianist_name,
                email=plan.email,
                max_hours_per_week=40,
            )
            db.add(pianist)
            db.flush()
            ensure_pianist_identity(db, pianist)
            pianist_id = pianist.id
            created_count += 1
            pianist_ids_by_key[key] = pianist_id

        pianist_ids = list(pianist_ids_by_key.values())
        days_replaced = (len(current_pianists) + len(pianist_ids)) * len(models.DAYS_ORDER)

        new_slots = [
            models.AvailabilitySlot(
                pianist_id=pianist_id,
                day=slot.day,
                slot_start_minute=slot.slot_start_minute,
                status=slot.status,
            )
            for pianist_key, slots in prepared.slots_by_pianist.items()
            for pianist_id in [pianist_ids_by_key[pianist_key]]
            for slot in slots
        ]
        db.add_all(new_slots)
        for pianist_id in pianist_ids:
            availability_state = db.get(models.PianistAvailabilityState, pianist_id)
            if availability_state is None:
                db.add(models.PianistAvailabilityState(pianist_id=pianist_id, is_complete=True))
            else:
                availability_state.is_complete = True
        bump_accompanist_revision(db)
        db.commit()
    except Exception as error:
        db.rollback()
        if isinstance(error, AvailabilityImportError):
            raise
        raise AvailabilityImportError("APPLY_FAILED", f"Availability was not applied: {error}") from error

    _STAGED_PREVIEWS.pop(request.preview_token, None)
    return schemas.AvailabilityImportApplyResult(
        pianists_removed=len(current_pianists),
        pianists_created=created_count,
        assignments_cleared=assignments_cleared,
        slots_replaced=slots_replaced,
        slots_created=len(new_slots),
        days_replaced=days_replaced,
        jury_availability_windows_removed=jury_availability_windows_removed,
    )