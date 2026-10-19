import { useState } from "react";
import { api } from "../../lib/api";
import { chooseLocalFile } from "../../lib/platform";
import type {
  LessonStudentImportInspection,
  LessonStudentImportMapping,
  LessonStudentImportPreview,
} from "../../lib/types";
import { ConfirmDialog, Modal } from "../clinical/ClinicalDialogs";

type FieldKey = "student_id" | "name" | "email";
const FIELDS: { key: FieldKey; label: string; required: boolean }[] = [
  { key: "student_id", label: "Student ID", required: true },
  { key: "name", label: "Name", required: true },
  { key: "email", label: "Email", required: false },
];
const WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function ColumnSelect({
  label,
  value,
  columns,
  emptyLabel,
  onChange,
}: {
  label: string;
  value: string | null;
  columns: string[];
  emptyLabel: string;
  onChange: (value: string) => void;
}) {
  return (
    <label>
      {label}
      <select aria-label={label} value={value ?? ""} onChange={(event) => onChange(event.target.value)}>
        <option value="">{emptyLabel}</option>
        {columns.map((column) => (
          <option key={column} value={column}>
            {column}
          </option>
        ))}
      </select>
    </label>
  );
}

export function LessonStudentImportDialog({
  onClose,
  onImported,
}: {
  onClose: () => void;
  onImported: (message: string) => Promise<void>;
}) {
  const [file, setFile] = useState<File | null>(null);
  const [inspection, setInspection] = useState<LessonStudentImportInspection | null>(null);
  const [mapping, setMapping] = useState<LessonStudentImportMapping | null>(null);
  const [preview, setPreview] = useState<LessonStudentImportPreview | null>(null);
  const [confirming, setConfirming] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  function load(next: LessonStudentImportInspection) {
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
      load(await api.inspectLessonStudentFile(chosen));
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
      load(await api.inspectLessonStudentFile(file, sheetName));
    } catch (cause) {
      setError(errorText(cause));
    } finally {
      setBusy(false);
    }
  }

  function updateField(key: FieldKey, value: string) {
    if (!mapping) return;
    setMapping({ ...mapping, [key]: value || null });
    setPreview(null);
  }

  function updateDay(day: string, value: string) {
    if (!mapping) return;
    setMapping({ ...mapping, days: { ...mapping.days, [day]: value || null } });
    setPreview(null);
  }

  const missingRequired = !mapping || FIELDS.some((field) => field.required && !mapping[field.key]);

  async function runPreview() {
    if (!file || !inspection || !mapping || missingRequired) return;
    setBusy(true);
    setError("");
    try {
      setPreview(await api.previewLessonStudentImport(file, inspection.selected_sheet, mapping));
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
      const result = await api.applyLessonStudentImport(file, inspection.selected_sheet, mapping);
      await onImported(
        `Imported ${result.students_created} ${
          result.students_created === 1 ? "student" : "students"
        }; replaced ${result.students_removed} and preserved ${result.placements_preserved} placement decision(s) in the stale schedule.`,
      );
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
                  <ColumnSelect
                    key={field.key}
                    label={field.label}
                    value={mapping[field.key]}
                    columns={inspection.columns}
                    emptyLabel={field.required ? "Choose column" : "Not in file"}
                    onChange={(value) => updateField(field.key, value)}
                  />
                ))}
              </div>
              <h4>Weekly availability (optional)</h4>
              <div className="clinical-import-mapping">
                {WEEKDAYS.map((day) => (
                  <ColumnSelect
                    key={day}
                    label={day}
                    value={mapping.days[day] ?? null}
                    columns={inspection.columns}
                    emptyLabel="Not in file"
                    onChange={(value) => updateDay(day, value)}
                  />
                ))}
              </div>
              <p className="dialog-note">
                Weekday cells list Availability Windows such as 8:00 AM-12:00 PM; prefix an entry with T for
                Tentative and separate entries with semicolons. Blank cells mean no availability. Tag selections are
                not imported; set them on each student after importing.
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
                    Fix the rows marked below in the source file, then choose the file again. Nothing has been changed.
                  </div>
                )}
                <div className="clinical-import-preview-wrap">
                  <table className="clinical-table">
                    <thead>
                      <tr>
                        <th>Row</th>
                        <th>Student ID</th>
                        <th>Name</th>
                        <th>Email</th>
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
                          <td>{row.email}</td>
                          <td>
                            {row.availability.length === 0
                              ? "None"
                              : row.availability.map((entry) => <div key={entry}>{entry}</div>)}
                          </td>
                          <td>
                            {row.errors.map((message) => (
                              <div key={message}>{message}</div>
                            ))}
                            {row.warnings.map((message) => (
                              <div key={message} className="muted">
                                {message}
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
                    Import Students
                  </button>
                </div>
              </section>
            )}
          </>
        )}
      </div>
      {confirming && (
        <ConfirmDialog
          title="Replace Students?"
          message=""
          confirmLabel="Replace Students"
          onConfirm={() => void commit()}
          onCancel={() => setConfirming(false)}
        >
          <p className="dialog-note">Importing the Student Roster will:</p>
          <ul className="dialog-note">
            <li>Delete all current students ({preview?.existing_student_count ?? 0} currently)</li>
            <li>Replace all students, their availability, and their tag selections with imported data</li>
            <li>Preserve existing placement decisions ({preview?.existing_placement_count ?? 0}); the schedule will be marked stale</li>
            <li>Clear lesson schedules</li>
            <li>Preserve Instructor records and Instructor availability</li>
            <li>Preserve the Lesson Tag registry</li>
          </ul>
        </ConfirmDialog>
      )}
    </Modal>
  );
}
