import { useState } from "react";
import { useScrollToPreview } from "../../lib/useScrollToPreview";
import { api } from "../../lib/api";
import { chooseLocalFile } from "../../lib/platform";
import type {
  TutorTutorImportField,
  TutorTutorImportInspection,
  TutorTutorImportMapping,
  TutorTutorImportPreview,
} from "../../lib/types";
import { ConfirmDialog } from "../../components/dialogs/ConfirmDialog";
import { ImportModalBody, ImportModalShell } from "../../components/imports/ImportDialogShell";
import { ImportTemplateLinks } from "../../components/ImportTemplateLinks";
import { ColumnSelect } from "../../components/imports/ColumnSelect";
import { DAYS_ORDER } from "../../lib/types";

const FIELDS: { key: TutorTutorImportField; label: string; required: boolean }[] = [
  { key: "tutor_name", label: "Tutor Name", required: true },
  { key: "tutor_email", label: "Tutor Email", required: false },
  { key: "room", label: "Room", required: false },
  { key: "lesson_length", label: "Lesson Length", required: false },
  { key: "capacity", label: "Capacity", required: false },
  { key: "tags", label: "Tags", required: false },
];

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export function TutorTutorImportDialog({
  onClose,
  onImported,
}: {
  onClose: () => void;
  onImported: (message: string) => Promise<void>;
}) {
  const [file, setFile] = useState<File | null>(null);
  const [inspection, setInspection] = useState<TutorTutorImportInspection | null>(null);
  const [mapping, setMapping] = useState<TutorTutorImportMapping | null>(null);
  const [preview, setPreview] = useState<TutorTutorImportPreview | null>(null);
  const previewRef = useScrollToPreview<HTMLElement>(preview);
  const [confirming, setConfirming] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  function load(next: TutorTutorImportInspection) {
    setInspection(next);
    setMapping(next.suggested_mapping);
    setPreview(null);
  }

  async function selectFile() {
    setError("");
    try {
      const chosen = await chooseLocalFile("Tutor roster spreadsheet", ["xlsx", "xls", "csv"]);
      if (!chosen) return;
      setBusy(true);
      load(await api.inspectTutorTutorFile(chosen));
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
      load(await api.inspectTutorTutorFile(file, sheetName));
    } catch (cause) {
      setError(errorText(cause));
    } finally {
      setBusy(false);
    }
  }

  function updateField(key: TutorTutorImportField, value: string) {
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
      setPreview(await api.previewTutorTutorImport(file, inspection.selected_sheet, mapping));
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
      const result = await api.applyTutorTutorImport(file, inspection.selected_sheet, mapping);
      await onImported(
        `Imported ${result.tutors_created} ${
          result.tutors_created === 1 ? "tutor" : "tutors"
        }; replaced ${result.tutors_removed}, cleared ${result.placements_removed} placement(s), and created ${result.tags_created} new tag(s).`,
      );
      onClose();
    } catch (cause) {
      setError(errorText(cause));
      setBusy(false);
    }
  }

  return (
    <ImportModalShell title="Import Tutor Roster" onClose={onClose}>
      <ImportModalBody error={error}>
        <div className="availability-import-file-row">
          <button type="button" disabled={busy} onClick={() => void selectFile()}>
            {inspection ? "Choose another file" : "Choose spreadsheet"}
          </button>
          <ImportTemplateLinks template="tutorTutors" disabled={busy} />
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
                Tentative and separate entries with semicolons. Blank cells mean no availability. Map one column to Tags and separate multiple tags in a cell with semicolons. Lesson Length defaults to 60 minutes and Capacity to 1
                when omitted.
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
                  Tags column: {preview.tag_columns.length ? preview.tag_columns[0] : "none"}. New tags:{" "}
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
                        <th>Tutor Name</th>
                        <th>Email</th>
                        <th>Room</th>
                        <th>Lesson Length</th>
                        <th>Capacity</th>
                        <th>Availability</th>
                        <th>Tags</th>
                        <th>Issues</th>
                      </tr>
                    </thead>
                    <tbody>
                      {preview.rows.map((row) => (
                        <tr key={row.row_number} className={row.errors.length ? "clinical-import-row-error" : ""}>
                          <td>{row.row_number}</td>
                          <td>{row.tutor_name}</td>
                          <td>{row.tutor_email}</td>
                          <td>{row.room}</td>
                          <td>{row.lesson_length_minutes === null ? "—" : `${row.lesson_length_minutes} min`}</td>
                          <td>{row.capacity ?? "—"}</td>
                          <td>
                            {row.availability.length === 0
                              ? "None"
                              : row.availability.map((entry) => <div key={entry}>{entry}</div>)}
                          </td>
                          <td>{row.tags.join(", ")}</td>
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
                    Import Tutors
                  </button>
                </div>
              </section>
            )}
          </>
        )}
      </ImportModalBody>
      {confirming && (
        <ConfirmDialog
          title="Replace Tutors?"
          message=""
          confirmLabel="Replace Tutors"
          onConfirm={() => void commit()}
          onCancel={() => setConfirming(false)}
        >
          <p className="dialog-note">Importing the Tutor Roster will:</p>
          <ul className="dialog-note">
            <li>Delete all current tutors ({preview?.existing_tutor_count ?? 0} currently)</li>
            <li>Replace all tutors, their availability, and their tags with imported data</li>
            <li>Preserve the Tutor Tag registry</li>
            <li>Preserve Student records</li>
            <li>Delete all existing placements ({preview?.existing_placement_count ?? 0}) and the generated schedule</li>
            <li>Clear tutor schedules</li>
          </ul>
        </ConfirmDialog>
      )}
    </ImportModalShell>
  );
}
