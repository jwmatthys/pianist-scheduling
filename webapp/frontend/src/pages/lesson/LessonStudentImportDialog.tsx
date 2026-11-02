import { useState } from "react";
import { useScrollToPreview } from "../../lib/useScrollToPreview";
import { api } from "../../lib/api";
import { chooseLocalFile } from "../../lib/platform";
import type {
  LessonStudentImportInspection,
  LessonStudentImportMapping,
  LessonStudentImportPreview,
} from "../../lib/types";
import { ConfirmDialog } from "../../components/dialogs/ConfirmDialog";
import { ImportModalBody, ImportModalShell } from "../../components/imports/ImportDialogShell";
import { ImportTemplateLinks } from "../../components/ImportTemplateLinks";
import { ColumnSelect } from "../../components/imports/ColumnSelect";
import { DAYS_ORDER } from "../../lib/types";

type FieldKey = "student_id" | "name" | "email";
type TagFieldKey = "required_tags" | "preferred_tags" | "restricted_tags";
const FIELDS: { key: FieldKey; label: string; required: boolean }[] = [
  { key: "student_id", label: "Student ID", required: true },
  { key: "name", label: "Name", required: true },
  { key: "email", label: "Email", required: false },
];
const TAG_FIELDS: { key: TagFieldKey; label: string }[] = [
  { key: "required_tags", label: "Required Tags" },
  { key: "preferred_tags", label: "Preferred Tags" },
  { key: "restricted_tags", label: "Restricted Tags" },
];

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
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
  const previewRef = useScrollToPreview<HTMLElement>(preview);
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

  function updateField(key: FieldKey | TagFieldKey, value: string) {
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
        }; replaced ${result.students_removed} and preserved ${result.placements_preserved} placement decision(s) in the stale schedule and created ${result.tags_created} new tag(s).`,
      );
      onClose();
    } catch (cause) {
      setError(errorText(cause));
      setBusy(false);
    }
  }

  return (
    <ImportModalShell title="Import Student Roster" onClose={onClose}>
      <ImportModalBody error={error}>
        <div className="availability-import-file-row">
          <button type="button" disabled={busy} onClick={() => void selectFile()}>
            {inspection ? "Choose another file" : "Choose spreadsheet"}
          </button>
          <ImportTemplateLinks template="lessonStudents" disabled={busy} />
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
              <h4>Tags (optional)</h4>
              <div className="clinical-import-mapping">
                {TAG_FIELDS.map((field) => (
                  <ColumnSelect
                    key={field.key}
                    label={field.label}
                    value={mapping[field.key]}
                    columns={inspection.columns}
                    emptyLabel="Not in file"
                    onChange={(value) => updateField(field.key, value)}
                  />
                ))}
              </div>
              <h4>Weekly availability (optional)</h4>
              <div className="clinical-import-mapping">
                {DAYS_ORDER.map((day) => (
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
                Tentative and separate entries with semicolons. Blank cells mean no availability. Separate multiple
                tags in a cell with semicolons. Tags named in the Required, Preferred or Restricted column get that state
                for the student; tags not yet in the registry are added to the Tag Manager.
              </p>
            </section>

            <div className="dialog-actions">
              <button
                type="button"
                disabled={busy || missingRequired}
                onClick={() => void runPreview()}
              >
                Preview Import
              </button>
            </div>

            {preview && (
              <section aria-label="Import preview" ref={previewRef}>
                <h3>Preview</h3>
                <p role="status">
                  {preview.total_rows} {preview.total_rows === 1 ? "row" : "rows"}: {preview.valid_rows} valid,{" "}
                  {preview.invalid_rows} with errors.
                </p>
                <p className="muted">
                  Tag columns:{" "}
                  {preview.tag_columns.length ? preview.tag_columns.join(", ") : "none"}. New tags:{" "}
                  {preview.new_tags.length ? preview.new_tags.join(", ") : "none"}. Existing tags reused:{" "}
                  {preview.reused_tags.length ? preview.reused_tags.join(", ") : "none"}.
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
                        <th>Required Tags</th>
                        <th>Preferred Tags</th>
                        <th>Restricted Tags</th>
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
                          <td>{row.required_tags.join(", ")}</td>
                          <td>{row.preferred_tags.join(", ")}</td>
                          <td>{row.restricted_tags.join(", ")}</td>
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
      </ImportModalBody>
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
            <li>Preserve the Lesson Tag registry and add any new imported tags to it</li>
          </ul>
        </ConfirmDialog>
      )}
    </ImportModalShell>
  );
}
