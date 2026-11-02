import { useRef, useState } from "react";
import { api } from "../lib/api";
import { chooseLocalFile } from "../lib/platform";
import { useScrollToPreview } from "../lib/useScrollToPreview";
import { ColumnSelect } from "./imports/ColumnSelect";
import { ImportDialogShell } from "./imports/ImportDialogShell";
import { ImportTemplateLinks } from "./ImportTemplateLinks";
import { DAYS_ORDER, type AvailabilityImportInspection, type AvailabilityImportMapping, type AvailabilityImportPreview } from "../lib/types";

type Props = {
  onBeforeApply: () => Promise<boolean>;
  onApplied: () => Promise<void>;
};

function formatMinute(value: number) {
  if (value === 1440) return "24:00";
  const hour = Math.floor(value / 60).toString().padStart(2, "0");
  const minute = (value % 60).toString().padStart(2, "0");
  return `${hour}:${minute}`;
}

export function AvailabilityImportDialog({ onBeforeApply, onApplied }: Props) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [inspection, setInspection] = useState<AvailabilityImportInspection | null>(null);
  const [personColumn, setPersonColumn] = useState("");
  const [emailColumn, setEmailColumn] = useState("");
  const [hoursColumn, setHoursColumn] = useState("");
  const [dayColumns, setDayColumns] = useState<Record<string, string | null>>({});
  const [preview, setPreview] = useState<AvailabilityImportPreview | null>(null);
  const previewRef = useScrollToPreview<HTMLElement>(preview);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);

  function openDialog() {
    setError(null);
    setSuccess(null);
    dialogRef.current?.showModal();
  }

  function loadInspection(next: AvailabilityImportInspection) {
    setInspection(next);
    setPersonColumn(next.suggested_columns.person_name_column ?? "");
    setEmailColumn(next.suggested_columns.email_column ?? "");
    setHoursColumn(next.suggested_columns.max_hours_column ?? "");
    setDayColumns(next.suggested_days);
    setPreview(null);
  }

  async function selectFile() {
    setError(null);
    setSuccess(null);
    try {
      const file = await chooseLocalFile("Availability spreadsheet", ["xlsx", "xls", "csv"]);
      if (!file) return;
      setBusy(true);
      loadInspection(await api.inspectAvailabilityFile(file));
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not inspect the selected file.");
    } finally {
      setBusy(false);
    }
  }

  async function chooseSheet(sheetName: string) {
    if (!inspection) return;
    setBusy(true);
    setError(null);
    try {
      loadInspection(await api.inspectAvailabilitySheet(inspection.upload_token, sheetName));
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not read the selected worksheet.");
    } finally {
      setBusy(false);
    }
  }

  function updateMapping(update: () => void) {
    update();
    setPreview(null);
  }

  function mappingPayload(): AvailabilityImportMapping {
    return {
      person_name_column: personColumn || null,
      email_column: emailColumn || null,
      max_hours_column: hoursColumn || null,
      days: dayColumns,
    };
  }

  async function createPreview() {
    if (!inspection) return;
    setBusy(true);
    setError(null);
    setSuccess(null);
    try {
      setPreview(await api.previewAvailabilityImport(
        inspection.upload_token,
        inspection.selected_sheet,
        mappingPayload(),
      ));
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not validate availability.");
    } finally {
      setBusy(false);
    }
  }

  async function applyPreview() {
    if (!preview?.can_apply) return;
    const ready = await onBeforeApply();
    if (!ready) return;
    const detail = `Replace the current Pianist roster with ${preview.incoming_pianist_count} Pianist(s) from this file? This removes ${preview.existing_pianist_count} current Pianist(s), their Accompanist weekly availability, ${preview.existing_assignment_count} Lesson assignments, and all Jury Availability Windows. Student Lessons, Jury Required values, and Panel choices remain.`;
    if (!window.confirm(detail)) return;
    setBusy(true);
    setError(null);
    try {
      const result = await api.applyAvailabilityImport(preview.preview_token);
      await onApplied();
      setSuccess(`Replaced ${result.pianists_removed} Pianist(s) with ${result.pianists_created} new Pianist(s), cleared ${result.assignments_cleared} Lesson assignments and ${result.jury_availability_windows_removed} Jury Availability Windows, and applied ${result.slots_created} Accompanist availability slots.`);
      setPreview(null);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not apply availability.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <>
      <button type="button" className="secondary-btn" onClick={openDialog}>
        Import Pianist Roster
      </button>
      <ImportDialogShell title="Import Pianist Roster" dialogRef={dialogRef} error={error} success={success}>
        <div className="availability-import-file-row">
          <button type="button" disabled={busy} onClick={() => void selectFile()}>
            {inspection ? "Choose another file" : "Choose spreadsheet"}
          </button>
          <ImportTemplateLinks template="pianistAvailability" disabled={busy} />
        </div>

        {inspection && (
          <>
            {inspection.sheets.length > 1 && (
              <label className="availability-map-field availability-sheet-select">
                <span>Worksheet</span>
                <select
                  value={inspection.selected_sheet ?? inspection.sheets[0]}
                  disabled={busy}
                  onChange={(event) => void chooseSheet(event.target.value)}
                >
                  {inspection.sheets.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}
                </select>
              </label>
            )}

            <section className="availability-mapping-section">
              <h3>Column mapping</h3>
              <ColumnSelect
                variant="field"
                label="Pianist name"
                value={personColumn}
                columns={inspection.columns}
                onChange={(value) => updateMapping(() => setPersonColumn(value))}
              />
              <div className="availability-normalized-mapping">
                <ColumnSelect
                  variant="field"
                  label="Email (optional)"
                  value={emailColumn || null}
                  columns={inspection.columns}
                  onChange={(value) => updateMapping(() => setEmailColumn(value))}
                />
                <ColumnSelect
                  variant="field"
                  label="Max hours per week (optional)"
                  value={hoursColumn || null}
                  columns={inspection.columns}
                  onChange={(value) => updateMapping(() => setHoursColumn(value))}
                />
              </div>
              <h4>Weekly availability</h4>
              <div className="availability-normalized-mapping">
                {DAYS_ORDER.map((day) => (
                  <ColumnSelect
                    variant="field"
                    key={day}
                    label={day}
                    value={dayColumns[day] ?? null}
                    columns={inspection.columns}
                    onChange={(value) => updateMapping(() => setDayColumns((current) => ({ ...current, [day]: value || null })))}
                  />
                ))}
              </div>
              <p className="muted">
                Weekday cells list Availability Windows such as 8:00 AM-12:00 PM; prefix an entry with T for Tentative and separate entries with semicolons. Blank cells mean no availability. Blank Max Hours Per Week defaults to 40.
              </p>
            </section>

            <details className="availability-source-preview">
              <summary>Source preview ({inspection.sample_rows.length} rows)</summary>
              <div className="preview-table-wrap">
                <table className="preview-table">
                  <thead><tr>{inspection.columns.map((column) => <th key={column}>{column}</th>)}</tr></thead>
                  <tbody>
                    {inspection.sample_rows.map((row, index) => (
                      <tr key={index}>{inspection.columns.map((column) => <td key={column}>{row[column]}</td>)}</tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </details>

            <p className="muted availability-completeness-note">
              Rows with the same Pianist name are combined. Applying replaces the entire Pianist roster and weekly Availability Windows; all Lesson assignments are cleared. Missing mappings or malformed cells block the import.
            </p>

            <div className="availability-import-actions">
              <button type="button" disabled={busy} onClick={() => void createPreview()}>
                {busy ? "Reviewing…" : "Preview Import"}
              </button>
            </div>
          </>
        )}

        {preview && (
          <section className="availability-import-review" aria-live="polite" ref={previewRef}>
            <h3>Import review</h3>
            <p className="availability-import-summary">
              {preview.incoming_pianist_count} incoming Pianists · {preview.valid_window_count} Availability Windows · {preview.warnings.length} warnings · {preview.errors.length} errors
            </p>
            <p className="muted">
              This replaces {preview.existing_pianist_count} current Pianists and clears {preview.existing_assignment_count} Lesson assignments. Every imported name receives a fresh Pianist record. Sparse Available/Tentative windows are stored; other times derive Unavailable.
            </p>
            <ul className="availability-import-plan">
              {preview.pianists.map((pianist, index) => (
                <li key={`${pianist.pianist_name}-${index}`} className={`is-${pianist.action}`}>
                  <div>
                    <strong>{pianist.pianist_name || "Unnamed respondent"}</strong>
                    <span>
                      {pianist.action === "new" ? "New Pianist"
                        : `Invalid row${pianist.row_numbers.length ? ` · row ${pianist.row_numbers.join(", ")}` : ""}`}
                    </span>
                  </div>
                  {pianist.action === "new" && <small>Email: {pianist.email || "blank"} · Max hours/week: 40</small>}
                </li>
              ))}
            </ul>
            {preview.errors.length > 0 && (
              <div className="availability-issue-list availability-errors" role="alert">
                <strong>Resolve these errors before applying:</strong>
                <ul>{preview.errors.map((issue, index) => <li key={`${issue.code}-${index}`}>{issue.row_number ? `Row ${issue.row_number}: ` : ""}{issue.message}</li>)}</ul>
              </div>
            )}
            {preview.warnings.length > 0 && (
              <div className="availability-issue-list availability-warnings">
                <strong>Review warnings:</strong>
                <ul>{preview.warnings.map((issue, index) => <li key={`${issue.code}-${index}`}>{issue.row_number ? `Row ${issue.row_number}: ` : ""}{issue.message}</li>)}</ul>
              </div>
            )}
            {preview.windows.length > 0 && (
              <div className="preview-table-wrap availability-windows-table-wrap">
                <table className="preview-table">
                  <thead><tr><th>Pianist</th><th>Day</th><th>Start</th><th>End</th><th>Status</th></tr></thead>
                  <tbody>
                    {preview.windows.map((item, index) => (
                      <tr key={`${item.pianist_name}-${item.day}-${item.start_minute}-${index}`}>
                        <td>{item.pianist_name}</td><td>{item.day}</td><td>{formatMinute(item.start_minute)}</td><td>{formatMinute(item.end_minute)}</td><td>{item.status}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {preview.valid_window_count > preview.windows.length && (
                  <p className="muted">Showing {preview.windows.length} of {preview.valid_window_count} windows.</p>
                )}
              </div>
            )}
            <div className="availability-import-actions">
              <button type="button" className="danger-btn" disabled={busy || !preview.can_apply} onClick={() => void applyPreview()}>
                {busy ? "Applying…" : "Apply availability"}
              </button>
            </div>
          </section>
        )}
      </ImportDialogShell>
    </>
  );
}