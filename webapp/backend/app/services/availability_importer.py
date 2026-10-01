"""Shared local availability import parsing with an Accompanist apply boundary."""

from __future__ import annotations

import re
import uuid
from dataclasses import dataclass, field

import pandas as pd
from sqlalchemy.orm import Session

from .. import models, schemas
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
class PreparedImport:
    upload_token: str
    sheet_name: str | None
    windows_by_pianist: dict[int, list[AvailabilityWindow]] = field(default_factory=dict)
    slots_by_pianist: dict[int, list[AccompanistAvailabilitySlot]] = field(default_factory=dict)
    names_by_pianist: dict[int, str] = field(default_factory=dict)
    days_by_pianist: dict[int, set[str]] = field(default_factory=dict)
    errors: list[AvailabilityIssue] = field(default_factory=list)
    warnings: list[AvailabilityIssue] = field(default_factory=list)
    rows_processed: int = 0
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
    name: str,
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
    person_column: str | None,
    pianists_by_name: dict[str, list[models.Pianist]],
    prepared: PreparedImport,
) -> models.Pianist | None:
    name_value = row.get(person_column) if person_column else None
    if _is_blank(name_value):
        prepared.errors.append(_issue("error", "MISSING_PERSON", "Pianist name is missing.", row_number))
        return None
    name = str(name_value).strip()
    matches = pianists_by_name.get(name.casefold(), [])
    if not matches:
        prepared.errors.append(_issue("error", "UNKNOWN_PIANIST", f"No pianist exactly matches '{name}'.", row_number))
        return None
    if len(matches) > 1:
        prepared.errors.append(_issue("error", "AMBIGUOUS_PIANIST", f"More than one pianist exactly matches '{name}'.", row_number))
        return None
    pianist = matches[0]
    prepared.names_by_pianist[pianist.id] = pianist.name
    prepared.windows_by_pianist.setdefault(pianist.id, [])
    prepared.days_by_pianist.setdefault(pianist.id, set())
    return pianist


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
    pianists = db.query(models.Pianist).order_by(models.Pianist.id).all()
    pianists_by_name: dict[str, list[models.Pianist]] = {}
    for pianist in pianists:
        pianists_by_name.setdefault(pianist.name.strip().casefold(), []).append(pianist)

    mapping = request.mapping
    if not prepared.errors:
        for index, row in frame.iterrows():
            row_number = int(index) + 2 if isinstance(index, (int, float)) else 0
            if all(_is_blank(value) for value in row.values):
                continue
            pianist = _person_for_row(row, row_number, mapping.person_name_column, pianists_by_name, prepared)
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
                    name=pianist.name,
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
                            name=pianist.name,
                            errors=prepared.errors,
                        )
                        if window:
                            row_windows.append(window)
            prepared.windows_by_pianist[pianist.id].extend(row_windows)
            prepared.days_by_pianist[pianist.id].update(window.day for window in row_windows)

    for pianist_id, windows in prepared.windows_by_pianist.items():
        normalized, issues = normalize_windows(windows)
        prepared.windows_by_pianist[pianist_id] = normalized
        prepared.errors.extend(issue for issue in issues if issue.severity == "error")
        prepared.warnings.extend(issue for issue in issues if issue.severity == "warning")
        slots, adapter_issues = windows_to_accompanist_slots(normalized)
        prepared.slots_by_pianist[pianist_id] = slots
        prepared.errors.extend(issue for issue in adapter_issues if issue.severity == "error")

    matched_ids = list(prepared.names_by_pianist)
    if matched_ids:
        prepared.existing_slots_in_scope = db.query(models.AvailabilitySlot).filter(
            models.AvailabilitySlot.pianist_id.in_(matched_ids)
        ).count()

    absent_pianists = len(pianists) - len(prepared.names_by_pianist)
    if absent_pianists:
        prepared.warnings.append(_issue(
            "warning", "PIANISTS_ABSENT_FROM_IMPORT",
            f"{absent_pianists} pianist(s) are absent from the import and will remain unchanged.",
        ))
    preview_token = uuid.uuid4().hex
    _STAGED_PREVIEWS[preview_token] = prepared
    while len(_STAGED_PREVIEWS) > MAX_STAGED_UPLOADS:
        _STAGED_PREVIEWS.pop(next(iter(_STAGED_PREVIEWS)))

    preview_pianists = [
        schemas.AvailabilityImportPianistOut(
            pianist_id=pianist_id,
            pianist_name=prepared.names_by_pianist[pianist_id],
            days=sorted(prepared.days_by_pianist.get(pianist_id, set()), key=models.DAYS_ORDER.index),
        )
        for pianist_id in prepared.names_by_pianist
    ]
    preview_windows = [
        schemas.AvailabilityImportWindowOut(
            pianist_id=pianist_id,
            pianist_name=prepared.names_by_pianist[pianist_id],
            day=window.day,
            start_minute=window.start_minute,
            end_minute=window.end_minute,
            status=window.status,
        )
        for pianist_id, windows in prepared.windows_by_pianist.items()
        for window in windows
    ][:MAX_PREVIEW_WINDOWS]
    can_apply = not prepared.errors and bool(matched_ids)
    return schemas.AvailabilityImportPreviewOut(
        preview_token=preview_token,
        sheet_name=request.sheet_name,
        rows_processed=prepared.rows_processed,
        matched_pianist_count=len(matched_ids),
        absent_pianist_count=absent_pianists,
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
    if not prepared.names_by_pianist:
        raise AvailabilityImportError("NO_MATCHED_PIANISTS", "No existing pianists matched the import.")
    pianist_ids = list(prepared.names_by_pianist)
    try:
        current_pianists = {
            pianist.id: pianist.name
            for pianist in db.query(models.Pianist).filter(models.Pianist.id.in_(pianist_ids)).all()
        }
        for pianist_id, expected_name in prepared.names_by_pianist.items():
            if current_pianists.get(pianist_id) != expected_name:
                raise AvailabilityImportError(
                    "PIANISTS_CHANGED", "Pianist records changed after preview; review the availability import again."
                )

        slots_replaced = db.query(models.AvailabilitySlot).filter(
            models.AvailabilitySlot.pianist_id.in_(pianist_ids)
        ).delete(synchronize_session=False)
        days_replaced = len(pianist_ids) * len(models.DAYS_ORDER)

        new_slots = [
            models.AvailabilitySlot(
                pianist_id=pianist_id,
                day=slot.day,
                slot_start_minute=slot.slot_start_minute,
                status=slot.status,
            )
            for pianist_id, slots in prepared.slots_by_pianist.items()
            for slot in slots
        ]
        db.add_all(new_slots)
        for pianist_id in pianist_ids:
            availability_state = db.get(models.PianistAvailabilityState, pianist_id)
            if availability_state is None:
                db.add(models.PianistAvailabilityState(pianist_id=pianist_id, is_complete=True))
            else:
                availability_state.is_complete = True
        db.commit()
    except Exception as error:
        db.rollback()
        if isinstance(error, AvailabilityImportError):
            raise
        raise AvailabilityImportError("APPLY_FAILED", f"Availability was not applied: {error}") from error

    _STAGED_PREVIEWS.pop(request.preview_token, None)
    return schemas.AvailabilityImportApplyResult(
        pianists_updated=len(pianist_ids),
        slots_replaced=slots_replaced,
        slots_created=len(new_slots),
        days_replaced=days_replaced,
    )