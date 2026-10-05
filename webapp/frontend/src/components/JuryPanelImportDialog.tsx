import { useRef, useState } from "react";
import { api } from "../lib/api";
import { chooseLocalFile } from "../lib/platform";
import type { JuryPanelImportInspection, JuryPanelImportMapping } from "../lib/types";

type Props = {
  onApplied: (message: string) => Promise<void>;
};

const FIELDS: { key: keyof JuryPanelImportMapping; label: string; required: boolean }[] = [
  { key: "schedule_date", label: "Schedule date", required: true },
  { key: "panel_name", label: "Panel name", required: true },
  { key: "room", label: "Room", required: false },
  { key: "earliest_start", label: "Earliest start", required: true },
  { key: "preferred_start", label: "Preferred start", required: false },
  { key: "jury_length", label: "Jury length (minutes)", required: true },
  { key: "break_needed", label: "Break needed (yes/no)", required: false },
  { key: "break_every", label: "Break every X juries", required: false },
  { key: "break_length", label: "Break length (minutes)", required: false },
  { key: "meal_break", label: "Meal break needed (yes/no)", required: false },
  { key: "meal_start", label: "Meal start", required: false },
  { key: "meal_end", label: "Meal end", required: false },
];

export function JuryPanelImportDialog({ onApplied }: Props) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [inspection, setInspection] = useState<JuryPanelImportInspection | null>(null);
  const [mapping, setMapping] = useState<JuryPanelImportMapping | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);

  function loadInspection(next: JuryPanelImportInspection) {
    setInspection(next);
    setMapping(next.suggested_mapping);
  }

  function openDialog() {
    setError(null);
    setSuccess(null);
    dialogRef.current?.showModal();
  }

  async function selectFile() {
    setError(null);
    setSuccess(null);
    try {
      const chosen = await chooseLocalFile("Jury Panel spreadsheet", ["xlsx", "xls", "csv"]);
      if (!chosen) return;
      setBusy(true);
      loadInspection(await api.inspectJuryPanelFile(chosen));
      setFile(chosen);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not inspect the selected file.");
    } finally {
      setBusy(false);
    }
  }

  async function chooseSheet(sheetName: string) {
    if (!file) return;
    setBusy(true);
    setError(null);
    try {
      loadInspection(await api.inspectJuryPanelFile(file, sheetName));
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not read the selected worksheet.");
    } finally {
      setBusy(false);
    }
  }

  const missingRequired = !mapping || FIELDS.some((field) => field.required && !mapping[field.key]);

  async function applyImport() {
    if (!file || !inspection || !mapping || missingRequired) return;
    const confirmed = window.confirm(
      "Importing Jury Panels will delete all current Jury Panels and clear all Panel assignments in the Lesson Roster. Continue?"
    );
    if (!confirmed) return;
    setBusy(true);
    setError(null);
    try {
      const result = await api.applyJuryPanelImport(file, inspection.selected_sheet, mapping);
      dialogRef.current?.close();
      await onApplied(`Removed ${result.panels_removed} Panel(s), cleared ${result.assignments_cleared} Panel assignment(s), and imported ${result.panels_created} Panel(s).`);
      setFile(null);
      setInspection(null);
      setMapping(null);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not import Jury Panels.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <>
      <button type="button" className="secondary-btn" onClick={openDialog}>Import jury panels</button>
      <dialog className="availability-import-dialog" ref={dialogRef}>
        <div className="availability-import-content">
          <header className="availability-import-heading">
            <div>
              <h2>Import Jury Panels</h2>
              <p className="muted">CSV, XLSX, or XLS</p>
            </div>
            <button type="button" className="dialog-close" onClick={() => dialogRef.current?.close()}>Close</button>
          </header>

          {error && <div className="error-banner" role="alert" style={{ whiteSpace: "pre-line" }}>{error}</div>}
          {success && <div className="availability-success" role="status">{success}</div>}

          <div className="availability-import-file-row">
            <button type="button" className="primary-btn" disabled={busy} onClick={() => void selectFile()}>
              {inspection ? "Choose another file" : "Choose spreadsheet"}
            </button>
          </div>

          {inspection && mapping && (
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
                <div className="availability-normalized-mapping">
                  {FIELDS.map((field) => (
                    <label className="availability-map-field" key={field.key}>
                      <span>{field.label}{field.required ? "" : " (optional)"}</span>
                      <select
                        aria-label={field.label}
                        value={mapping[field.key] ?? ""}
                        onChange={(event) => setMapping({ ...mapping, [field.key]: event.target.value || null })}
                      >
                        <option value="">{field.required ? "Choose column" : "Not in file"}</option>
                        {inspection.columns.map((column) => <option key={column} value={column}>{column}</option>)}
                      </select>
                    </label>
                  ))}
                </div>
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
                Importing deletes all current Jury Panels and clears every Panel assignment in the Lesson Roster. Break and meal values are used only when the matching "needed" column is Yes.
              </p>

              <div className="availability-import-actions">
                <button type="button" className="primary-btn" disabled={busy || missingRequired} onClick={() => void applyImport()}>
                  {busy ? "Importing…" : "Import jury panels"}
                </button>
              </div>
            </>
          )}
        </div>
      </dialog>
    </>
  );
}
