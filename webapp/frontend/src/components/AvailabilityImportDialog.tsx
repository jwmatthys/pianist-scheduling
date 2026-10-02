import { useRef, useState } from "react";
import { api } from "../lib/api";
import { chooseLocalFile, saveTextFile } from "../lib/platform";
import { generateAvailabilityTemplateCsv } from "../lib/availabilityTemplate";
import { DAYS_ORDER, type AvailabilityImportInspection, type AvailabilityImportLayout, type AvailabilityImportMapping, type AvailabilityImportPreview, type AvailabilityStatus, type AvailabilityWindowColumns } from "../lib/types";

type Props = {
  onBeforeApply: () => Promise<boolean>;
  onApplied: () => Promise<void>;
};

const COLUMN_LABELS: Record<string, string> = {
  person_name_column: "Pianist name",
  email_column: "Email",
  day_column: "Weekday",
  start_column: "Start time",
  end_column: "End time",
  status_column: "Status",
};

function ColumnSelect({
  label,
  value,
  columns,
  onChange,
}: {
  label: string;
  value: string | null;
  columns: string[];
  onChange: (value: string | null) => void;
}) {
  return (
    <label className="availability-map-field">
      <span>{label}</span>
      <select aria-label={label} value={value ?? ""} onChange={(event) => onChange(event.target.value || null)}>
        <option value="">Choose column</option>
        {columns.map((column) => <option key={column} value={column}>{column}</option>)}
      </select>
    </label>
  );
}

function initialMappings(inspection: AvailabilityImportInspection) {
  const suggested = inspection.suggested_normalized;
  const hasNormalized = Boolean(
    suggested.person_name_column && suggested.day_column && suggested.start_column &&
    suggested.end_column && suggested.status_column
  );
  const hasWide = Object.values(inspection.suggested_wide_windows).some((pairs) => pairs.length > 0);
  return {
    layout: hasNormalized || !hasWide ? "normalized" as const : "wide" as const,
    person: suggested.person_name_column ?? "",
    email: suggested.email_column ?? "",
    day: suggested.day_column ?? "",
    start: suggested.start_column ?? "",
    end: suggested.end_column ?? "",
    status: suggested.status_column ?? "",
    wideWindows: inspection.suggested_wide_windows,
  };
}

function formatMinute(value: number) {
  if (value === 1440) return "24:00";
  const hour = Math.floor(value / 60).toString().padStart(2, "0");
  const minute = (value % 60).toString().padStart(2, "0");
  return `${hour}:${minute}`;
}

export function AvailabilityImportDialog({ onBeforeApply, onApplied }: Props) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [inspection, setInspection] = useState<AvailabilityImportInspection | null>(null);
  const [layout, setLayout] = useState<AvailabilityImportLayout>("normalized");
  const [personColumn, setPersonColumn] = useState("");
  const [emailColumn, setEmailColumn] = useState("");
  const [dayColumn, setDayColumn] = useState("");
  const [startColumn, setStartColumn] = useState("");
  const [endColumn, setEndColumn] = useState("");
  const [statusColumn, setStatusColumn] = useState("");
  const [wideStatus, setWideStatus] = useState<AvailabilityStatus>("Available");
  const [wideWindows, setWideWindows] = useState<Record<string, AvailabilityWindowColumns[]>>({});
  const [preview, setPreview] = useState<AvailabilityImportPreview | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);

  function openDialog() {
    setError(null);
    setSuccess(null);
    dialogRef.current?.showModal();
  }

  function loadInspection(next: AvailabilityImportInspection) {
    const mapping = initialMappings(next);
    setInspection(next);
    setLayout(mapping.layout);
    setPersonColumn(mapping.person);
    setEmailColumn(mapping.email);
    setDayColumn(mapping.day);
    setStartColumn(mapping.start);
    setEndColumn(mapping.end);
    setStatusColumn(mapping.status);
    setWideWindows(mapping.wideWindows);
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
    return layout === "normalized"
      ? {
          layout,
          person_name_column: personColumn || null,
          email_column: emailColumn || null,
          day_column: dayColumn || null,
          start_column: startColumn || null,
          end_column: endColumn || null,
          status_column: statusColumn || null,
        }
      : {
          layout,
          person_name_column: personColumn || null,
          email_column: emailColumn || null,
          wide_status: wideStatus,
          wide_windows: wideWindows,
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

  async function saveAvailabilityTemplate() {
    setError(null);
    setSuccess(null);
    try {
      const saved = await saveTextFile(
        generateAvailabilityTemplateCsv(),
        "pianist-availability-template.csv",
        "Pianist availability CSV",
        "csv",
      );
      if (saved) setSuccess("Availability template saved.");
    } catch {
      setError("Could not save the availability template. Check the selected destination and try again.");
    }
  }

  function changeWideWindow(day: string, index: number, property: keyof AvailabilityWindowColumns, value: string | null) {
    updateMapping(() => setWideWindows((current) => ({
      ...current,
      [day]: current[day].map((pair, pairIndex) => pairIndex === index ? { ...pair, [property]: value } : pair),
    })));
  }

  return (
    <>
      <button type="button" className="secondary-btn" onClick={openDialog}>
        Import Availability
      </button>
      <dialog className="availability-import-dialog" ref={dialogRef}>
        <div className="availability-import-content">
          <header className="availability-import-heading">
            <div>
              <h2>Import Pianist Availability</h2>
              <p className="muted">CSV, XLSX, or XLS</p>
            </div>
            <button type="button" className="dialog-close" onClick={() => dialogRef.current?.close()}>
              Close
            </button>
          </header>

          {error && <div className="error-banner" role="alert">{error}</div>}
          {success && <div className="availability-success" role="status">{success}</div>}

          <div className="availability-import-file-row">
            <button type="button" className="primary-btn" disabled={busy} onClick={() => void selectFile()}>
              {inspection ? "Choose another file" : "Choose spreadsheet"}
            </button>
            <button type="button" className="template-download-button" disabled={busy} onClick={() => void saveAvailabilityTemplate()}>
              Download CSV template
            </button>
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

              <label className="availability-map-field availability-layout-select">
                <span>Spreadsheet layout</span>
                <select value={layout} onChange={(event) => {
                  updateMapping(() => setLayout(event.target.value as AvailabilityImportLayout));
                }}>
                  <option value="normalized">One availability window per row</option>
                  <option value="wide">One respondent per row</option>
                </select>
              </label>

              <section className="availability-mapping-section">
                <h3>Column mapping</h3>
                <ColumnSelect
                  label="Pianist name"
                  value={personColumn}
                  columns={inspection.columns}
                  onChange={(value) => updateMapping(() => setPersonColumn(value ?? ""))}
                />
                <div className="availability-normalized-mapping">
                  <ColumnSelect
                    label="Email (optional)"
                    value={emailColumn || null}
                    columns={inspection.columns}
                    onChange={(value) => updateMapping(() => setEmailColumn(value ?? ""))}
                  />
                </div>
                {layout === "normalized" ? (
                  <div className="availability-normalized-mapping">
                    {(["day_column", "start_column", "end_column", "status_column"] as const).map((field) => {
                      const value = {
                        day_column: dayColumn,
                        start_column: startColumn,
                        end_column: endColumn,
                        status_column: statusColumn,
                      }[field];
                      const setter = {
                        day_column: setDayColumn,
                        start_column: setStartColumn,
                        end_column: setEndColumn,
                        status_column: setStatusColumn,
                      }[field];
                      return (
                        <ColumnSelect
                          key={field}
                          label={COLUMN_LABELS[field]}
                          value={value}
                          columns={inspection.columns}
                          onChange={(next) => updateMapping(() => setter(next ?? ""))}
                        />
                      );
                    })}
                  </div>
                ) : (
                  <>
                    <label className="availability-map-field">
                      <span>Status for mapped windows</span>
                      <select value={wideStatus} onChange={(event) => updateMapping(() => setWideStatus(event.target.value as AvailabilityStatus))}>
                        <option>Available</option>
                        <option>Tentative</option>
                        <option>Unavailable</option>
                      </select>
                    </label>
                    <div className="wide-day-mappings">
                      {DAYS_ORDER.map((day) => {
                        const pairs = wideWindows[day] ?? [];
                        return (
                          <details key={day} open={pairs.length > 0}>
                            <summary>{day}{pairs.length ? ` (${pairs.length} window${pairs.length === 1 ? "" : "s"})` : ""}</summary>
                            {pairs.map((pair, index) => (
                              <div className="wide-window-pair" key={`${day}-${index}`}>
                                <ColumnSelect
                                  label={`${day} window ${index + 1} start`}
                                  value={pair.start_column}
                                  columns={inspection.columns}
                                  onChange={(value) => changeWideWindow(day, index, "start_column", value)}
                                />
                                <ColumnSelect
                                  label={`${day} window ${index + 1} end`}
                                  value={pair.end_column}
                                  columns={inspection.columns}
                                  onChange={(value) => changeWideWindow(day, index, "end_column", value)}
                                />
                                <button type="button" className="danger-btn small" onClick={() => updateMapping(() => setWideWindows((current) => ({
                                  ...current,
                                  [day]: current[day].filter((_, pairIndex) => pairIndex !== index),
                                })))}>
                                  Remove
                                </button>
                              </div>
                            ))}
                            <button type="button" className="secondary-btn small-button" onClick={() => updateMapping(() => setWideWindows((current) => ({
                              ...current,
                              [day]: [...(current[day] ?? []), { start_column: null, end_column: null }],
                            })))}>
                              Add time window
                            </button>
                          </details>
                        );
                      })}
                    </div>
                  </>
                )}
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
                Names in this file group that respondent's rows. Applying replaces the whole Pianist roster and weekly availability; all Lesson-to-Pianist assignments are cleared. Blank mapped windows mean Unavailable. Missing mappings or malformed rows block the import.
              </p>

              <div className="availability-import-actions">
                <button type="button" className="primary-btn" disabled={busy} onClick={() => void createPreview()}>
                  {busy ? "Reviewing…" : "Preview and validate"}
                </button>
              </div>
            </>
          )}

          {preview && (
            <section className="availability-import-review" aria-live="polite">
              <h3>Import review</h3>
              <p className="availability-import-summary">
                {preview.incoming_pianist_count} Pianists · {preview.valid_window_count} Availability Windows · {preview.warnings.length} warnings · {preview.errors.length} errors
              </p>
              <p className="muted">
                This will replace {preview.existing_pianist_count} current Pianists and clear {preview.existing_assignment_count} Lesson assignments. Imported names receive fresh internal records. Sparse Available/Tentative windows are stored; other times derive Unavailable.
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
                <button type="button" className="primary-btn" disabled={busy || !preview.can_apply} onClick={() => void applyPreview()}>
                  {busy ? "Applying…" : "Apply availability"}
                </button>
              </div>
            </section>
          )}
        </div>
      </dialog>
    </>
  );
}