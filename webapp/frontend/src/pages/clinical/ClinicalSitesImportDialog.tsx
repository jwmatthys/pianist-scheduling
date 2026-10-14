import { useState } from "react";
import { api } from "../../lib/api";
import { chooseLocalFile } from "../../lib/platform";
import { formatMinutes } from "../../lib/time";
import type {
  ClinicalSitesImportField,
  ClinicalSitesImportInspection,
  ClinicalSitesImportMapping,
  ClinicalSitesImportPreview,
} from "../../lib/types";
import { ConfirmDialog, Modal } from "./ClinicalDialogs";

const FIELDS: { key: ClinicalSitesImportField; label: string; required: boolean }[] = [
  { key: "site_name", label: "Site Name", required: true },
  { key: "site_address", label: "Site Address", required: false },
  { key: "contact_name", label: "Contact Name", required: false },
  { key: "contact_email", label: "Contact Email", required: false },
  { key: "contact_phone", label: "Contact Phone", required: false },
  { key: "day_of_week", label: "Day of Week", required: true },
  { key: "start_time", label: "Start Time", required: true },
  { key: "end_time", label: "End Time", required: true },
  { key: "session_length", label: "Session Length", required: true },
  { key: "travel_time", label: "Travel Time", required: false },
  { key: "capacity", label: "Capacity", required: false },
];

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export function ClinicalSitesImportDialog({
  onClose,
  onImported,
}: {
  onClose: () => void;
  onImported: (message: string) => Promise<void>;
}) {
  const [file, setFile] = useState<File | null>(null);
  const [inspection, setInspection] = useState<ClinicalSitesImportInspection | null>(null);
  const [mapping, setMapping] = useState<ClinicalSitesImportMapping | null>(null);
  const [preview, setPreview] = useState<ClinicalSitesImportPreview | null>(null);
  const [confirming, setConfirming] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  function load(next: ClinicalSitesImportInspection) {
    setInspection(next);
    setMapping(next.suggested_mapping);
    setPreview(null);
  }

  async function selectFile() {
    setError("");
    try {
      const chosen = await chooseLocalFile("Clinical Sites spreadsheet", ["xlsx", "xls", "csv"]);
      if (!chosen) return;
      setBusy(true);
      load(await api.inspectClinicalSitesFile(chosen));
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
      load(await api.inspectClinicalSitesFile(file, sheetName));
    } catch (cause) {
      setError(errorText(cause));
    } finally {
      setBusy(false);
    }
  }

  function updateMapping(key: ClinicalSitesImportField, value: string) {
    if (!mapping) return;
    setMapping({ ...mapping, [key]: value || null });
    setPreview(null);
  }

  const missingRequired = !mapping || FIELDS.some((field) => field.required && !mapping[field.key]);

  async function runPreview() {
    if (!file || !inspection || !mapping || missingRequired) return;
    setBusy(true);
    setError("");
    try {
      setPreview(await api.previewClinicalSitesImport(file, inspection.selected_sheet, mapping));
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
      const result = await api.applyClinicalSitesImport(file, inspection.selected_sheet, mapping);
      await onImported(
        `Imported ${result.opportunities_created} Placement ${
          result.opportunities_created === 1 ? "Opportunity" : "Opportunities"
        }; replaced ${result.opportunities_removed}, removed ${result.placements_removed} placement(s), and created ${result.tags_created} new tag(s).`,
      );
      onClose();
    } catch (cause) {
      setError(errorText(cause));
      setBusy(false);
    }
  }

  return (
    <Modal title="Import Clinical Sites" onClose={onClose} className="clinical-wide-dialog clinical-import-dialog">
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
              <p className="dialog-note">
                Every column not mapped above is read as tags: each nonblank cell is exactly one tag (not split on
                commas or semicolons), and the column heading is ignored. Capacity defaults to 1 when omitted.
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
                <p className="muted">
                  Tag columns: {preview.tag_columns.length ? preview.tag_columns.join(", ") : "none"}. New tags:{" "}
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
                        <th>Site Name</th>
                        <th>Day</th>
                        <th>Start Time</th>
                        <th>End Time</th>
                        <th>Session Length</th>
                        <th>Travel Time</th>
                        <th>Capacity</th>
                        <th>Tags</th>
                        <th>Issues</th>
                      </tr>
                    </thead>
                    <tbody>
                      {preview.rows.map((row) => (
                        <tr key={row.row_number} className={row.errors.length ? "clinical-import-row-error" : ""}>
                          <td>{row.row_number}</td>
                          <td>{row.site_name}</td>
                          <td>{row.day_of_week ?? "—"}</td>
                          <td>{row.start_minute === null ? "—" : formatMinutes(row.start_minute)}</td>
                          <td>{row.end_minute === null ? "—" : formatMinutes(row.end_minute)}</td>
                          <td>{row.session_length_minutes === null ? "—" : `${row.session_length_minutes} min`}</td>
                          <td>{row.travel_minutes === null ? "—" : `${row.travel_minutes} min`}</td>
                          <td>{row.capacity ?? "—"}</td>
                          <td>{row.tags.join(", ")}</td>
                          <td>
                            {row.errors.map((message) => (
                              <div key={message}>{message}</div>
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
                    Import Clinical Sites
                  </button>
                </div>
              </section>
            )}
          </>
        )}
      </div>
      {confirming && (
        <ConfirmDialog
          title="Replace Clinical Sites?"
          message=""
          confirmLabel="Replace Clinical Sites"
          onConfirm={() => void commit()}
          onCancel={() => setConfirming(false)}
        >
          <p className="dialog-note">Importing Clinical Sites will:</p>
          <ul className="dialog-note">
            <li>
              Delete all current Placement Opportunities ({preview?.existing_opportunity_count ?? 0} currently)
            </li>
            <li>Replace all Placement Opportunities with imported data</li>
            <li>Preserve the Clinical Tag registry</li>
            <li>Preserve Student records</li>
            <li>Preserve Student tag selections</li>
            <li>Remove any existing Clinical Placements ({preview?.existing_placement_count ?? 0} currently)</li>
            <li>Mark schedules as cleared/not generated</li>
          </ul>
        </ConfirmDialog>
      )}
    </Modal>
  );
}
