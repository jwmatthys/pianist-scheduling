import { useState } from "react";
import { api } from "../../lib/api";
import { chooseLocalFile } from "../../lib/platform";
import type {
  ClinicalStudentsImportField,
  ClinicalStudentsImportInspection,
  ClinicalStudentsImportMapping,
  ClinicalStudentsImportPreview,
} from "../../lib/types";
import { ConfirmDialog, Modal } from "./ClinicalDialogs";

const FIELDS: { key: ClinicalStudentsImportField; label: string; required: boolean }[] = [
  { key: "student_id", label: "Student ID", required: true },
  { key: "name", label: "Name", required: false },
  { key: "email", label: "Email", required: false },
];

const WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export function ClinicalStudentRosterImportDialog({
  onClose,
  onImported,
}: {
  onClose: () => void;
  onImported: (message: string) => Promise<void>;
}) {
  const [file, setFile] = useState<File | null>(null);
  const [inspection, setInspection] = useState<ClinicalStudentsImportInspection | null>(null);
  const [mapping, setMapping] = useState<ClinicalStudentsImportMapping | null>(null);
  const [preview, setPreview] = useState<ClinicalStudentsImportPreview | null>(null);
  const [confirming, setConfirming] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  function load(next: ClinicalStudentsImportInspection) {
    setInspection(next);
    setMapping(next.suggested_mapping);
    setPreview(null);
  }

  async function selectFile() {
    setError("");
    try {
      const chosen = await chooseLocalFile("Student roster spreadsheet", ["xlsx", "xls", "csv"]);
      if (!chosen) return;
      setBusy(true);
      load(await api.inspectClinicalStudentsFile(chosen));
      setFile(chosen);
    } catch (cause) {
      setError(errorText(cause));
    } finally {
      setBusy(false);
    }
  }

  async function chooseSheet(sheetName: string) {
    if (!file) return;
    setBusy(true);
    setError("");
    try {
      load(await api.inspectClinicalStudentsFile(file, sheetName));
    } catch (cause) {
      setError(errorText(cause));
    } finally {
      setBusy(false);
    }
  }

  function updateMapping(key: ClinicalStudentsImportField, value: string) {
    if (!mapping) return;
    setMapping({ ...mapping, [key]: value || null });
    setPreview(null);
  }

  function updateDay(day: string, value: string) {
    if (!mapping) return;
    setMapping({ ...mapping, days: { ...mapping.days, [day]: value || null } });
    setPreview(null);
  }

  const mappedDays = mapping ? WEEKDAYS.filter((day) => mapping.days[day]) : [];
  const missingRequired = !mapping || FIELDS.some((field) => field.required && !mapping[field.key]);

  async function runPreview() {
    if (!file || !inspection || !mapping || missingRequired) return;
    setBusy(true);
    setError("");
    try {
      setPreview(await api.previewClinicalStudentsImport(file, inspection.selected_sheet, mapping));
    } catch (cause) {
      setPreview(null);
      setError(errorText(cause));
    } finally {
      setBusy(false);
    }
  }

  async function commit() {
    if (!file || !inspection || !mapping) return;
    setConfirming(false);
    setBusy(true);
    setError("");
    try {
      const result = await api.applyClinicalStudentsImport(file, inspection.selected_sheet, mapping);
      const plural = (count: number, noun: string) => `${count} ${noun}${count === 1 ? "" : "s"}`;
      const parts = [
        `Imported ${plural(result.students_created + result.students_updated, "student")} (${result.students_created} created, ${result.students_updated} updated)`,
      ];
      if (mappedDays.length > 0) {
        parts.push(`replaced availability for ${plural(result.availability_updated, "student")}`);
      }
      await onImported(`${parts.join("; ")}.`);
      onClose();
    } catch (cause) {
      setError(errorText(cause));
      setBusy(false);
    }
  }

  return (
    <Modal title="Import Student Roster" onClose={onClose} className="clinical-wide-dialog clinical-import-dialog">
      <div className="clinical-tag-manager">
        <p className="muted">CSV, XLSX, or XLS</p>
        {error && (
          <div className="error-banner" role="alert" style={{ whiteSpace: "pre-line" }}>
            {error}
          </div>
        )}
        <div>
          <button type="button" className="primary-btn" disabled={busy} onClick={() => void selectFile()}>
            {inspection ? "Choose another file" : "Choose spreadsheet"}
          </button>
        </div>

        {inspection && mapping && (
          <>
            {inspection.sheets.length > 1 && (
              <label>
                Worksheet
                <select
                  value={inspection.selected_sheet ?? inspection.sheets[0]}
                  disabled={busy}
                  onChange={(event) => void chooseSheet(event.target.value)}
                >
                  {inspection.sheets.map((sheet) => (
                    <option key={sheet} value={sheet}>
                      {sheet}
                    </option>
                  ))}
                </select>
              </label>
            )}

            <section>
              <h3>Column mapping</h3>
              <div className="clinical-import-mapping">
                {FIELDS.map((field) => (
                  <label key={field.key}>
                    {field.label}
                    {field.required ? "" : " (optional)"}
                    <select
                      aria-label={field.label}
                      value={mapping[field.key] ?? ""}
                      onChange={(event) => updateMapping(field.key, event.target.value)}
                    >
                      <option value="">{field.required ? "Choose column" : "Not in file"}</option>
                      {inspection.columns.map((column) => (
                        <option key={column} value={column}>
                          {column}
                        </option>
                      ))}
                    </select>
                  </label>
                ))}
              </div>
              <h4>Weekly availability (optional)</h4>
              <div className="clinical-import-mapping">
                {WEEKDAYS.map((day) => (
                  <label key={day}>
                    {day}
                    <select
                      aria-label={day}
                      value={mapping.days[day] ?? ""}
                      onChange={(event) => updateDay(day, event.target.value)}
                    >
                      <option value="">Not in file</option>
                      {inspection.columns.map((column) => (
                        <option key={column} value={column}>
                          {column}
                        </option>
                      ))}
                    </select>
                  </label>
                ))}
              </div>
              <p className="dialog-note">
                Name is required only for new students. Map at least one weekday column to replace each listed
                student&apos;s Availability Windows (blank cells mean no availability); with none mapped, availability
                is left unchanged. All other columns are ignored.
              </p>
            </section>

            <div className="dialog-actions">
              <button
                type="button"
                className="primary-btn"
                disabled={busy || missingRequired}
                onClick={() => void runPreview()}
              >
                Preview Import
              </button>
            </div>

            {preview && (
              <section aria-label="Import preview">
                <h3>Preview</h3>
                <p role="status">
                  {preview.total_rows} {preview.total_rows === 1 ? "row" : "rows"}: {preview.valid_rows} valid,{" "}
                  {preview.invalid_rows} with errors.
                </p>
                {!preview.can_apply && (
                  <div className="error-banner" role="alert">
                    {preview.total_rows === 0
                      ? "The selected file contains no students."
                      : "Fix the rows marked below in the source file, then choose the file again."}{" "}
                    Nothing has been changed.
                  </div>
                )}
                <div className="clinical-import-preview-wrap">
                  <table className="clinical-table">
                    <thead>
                      <tr>
                        <th>Row</th>
                        <th>Student ID</th>
                        <th>Name</th>
                        <th>Student</th>
                        <th>Availability</th>
                        <th>Issues</th>
                      </tr>
                    </thead>
                    <tbody>
                      {preview.rows.map((row) => (
                        <tr key={row.row_number} className={row.errors.length ? "clinical-import-row-error" : ""}>
                          <td>{row.row_number}</td>
                          <td>{row.student_id}</td>
                          <td>{row.name}</td>
                          <td>{row.student_change}</td>
                          <td>{row.availability_change}</td>
                          <td>
                            {row.errors.map((message) => (
                              <div key={message}>{message}</div>
                            ))}
                            {row.warnings.map((message) => (
                              <div key={message} className="muted">
                                Warning: {message}
                              </div>
                            ))}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="dialog-actions">
                  <button
                    type="button"
                    className="primary-btn"
                    disabled={busy || !preview.can_apply}
                    onClick={() => setConfirming(true)}
                  >
                    Import Student Roster
                  </button>
                </div>
              </section>
            )}
          </>
        )}
      </div>
      {confirming && (
        <ConfirmDialog
          title="Import Student Roster?"
          message=""
          confirmLabel="Import Student Roster"
          onConfirm={() => void commit()}
          onCancel={() => setConfirming(false)}
        >
          <p className="dialog-note">Importing this file will:</p>
          <ul className="dialog-note">
            <li>Create {preview?.students_to_create ?? 0} new student(s)</li>
            <li>Update {preview?.students_to_update ?? 0} existing student(s) matched by Student ID</li>
            {preview?.availability_mapped ? (
              <li>Replace weekly availability for the students listed in the file</li>
            ) : (
              <li>Leave student availability unchanged</li>
            )}
            <li>Preserve students not listed in the file</li>
            <li>Preserve Clinical Sites and the Clinical Tag registry</li>
            <li>Preserve tag selections and placements</li>
            <li>Mark any generated schedule as out of date</li>
          </ul>
        </ConfirmDialog>
      )}
    </Modal>
  );
}
