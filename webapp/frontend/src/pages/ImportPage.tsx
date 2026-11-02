import { useState } from "react";
import { api } from "../lib/api";
import type { ImportPreview } from "../lib/types";
import { ImportTemplateLinks } from "../components/ImportTemplateLinks";
import { chooseLocalFile } from "../lib/platform";
import { useScrollToPreview } from "../lib/useScrollToPreview";

const FIELD_LABELS: Record<string, string> = {
  teacher: "Teacher",
  teacher_email: "Instructor email (optional)",
  student: "Student",
  student_id: "Student ID (optional)",
  day: "Day",
  start_time: "Start time",
  end_time: "End time",
  room: "Room",
  instrument: "Instrument",
  need_pianist: "Needs pianist?",
  jury_required: "Jury required?",
  required_pianist_name: "Specific pianist (optional)",
};

const REQUIRED_FIELDS = new Set(["day", "start_time"]);

function normalizeHeader(value: string): string {
  return value.toLowerCase().replace(/\(optional\)/g, "").replace(/[^a-z0-9]+/g, " ").trim();
}

function guessColumn(field: string, columns: string[]): string {
  // Exact field-label matches first so downloaded templates map every column.
  const exactTargets = [normalizeHeader(field), normalizeHeader(FIELD_LABELS[field] ?? "")];
  const exact = columns.find((c) => exactTargets.includes(normalizeHeader(c)));
  if (exact) return exact;
  const normalized = field.replace(/_/g, " ");
  const found = columns.find((c) => c.toLowerCase().includes(normalized.split(" ")[0]));
  return found ?? "";
}

export function ImportPage({ onImported }: { onImported: () => void }) {
  const [preview, setPreview] = useState<ImportPreview | null>(null);
  const [mapping, setMapping] = useState<Record<string, string>>({});
  const [reviewing, setReviewing] = useState(false);
  const previewRef = useScrollToPreview<HTMLElement>(reviewing ? preview : null);
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<{ created: number; skipped: number; warnings: string[] } | null>(null);
  const [error, setError] = useState<string | null>(null);

  async function handleFile(file: File) {
    setError(null);
    setResult(null);
    setReviewing(false);
    try {
      const preview = await api.previewImport(file);
      setPreview(preview);
      const guessed: Record<string, string> = {};
      for (const field of Object.keys(FIELD_LABELS)) {
        guessed[field] = guessColumn(field, preview.columns);
      }
      setMapping(guessed);
    } catch (e: any) {
      setError(e.message ?? String(e));
    }
  }

  async function chooseFile() {
    try {
      const file = await chooseLocalFile("Lesson spreadsheet", ["xlsx", "xls", "csv"]);
      if (file) await handleFile(file);
    } catch (e: any) {
      setError(e.message ?? String(e));
    }
  }

  function showPreview() {
    const missing = Array.from(REQUIRED_FIELDS).filter((f) => !mapping[f]);
    if (missing.length) {
      setError(`Please map required fields: ${missing.map((f) => FIELD_LABELS[f]).join(", ")}`);
      return;
    }
    setError(null);
    setReviewing(true);
  }

  async function commit() {
    if (!preview) return;
    const missing = Array.from(REQUIRED_FIELDS).filter((f) => !mapping[f]);
    if (missing.length) {
      setError(`Please map required fields: ${missing.map((f) => FIELD_LABELS[f]).join(", ")}`);
      return;
    }
    const existingLessons = await api.listLessons();
    if (existingLessons.length && !confirm("Importing will delete all existing lessons. Continue?")) return;
    setBusy(true);
    setError(null);
    try {
      const res = await api.commitImport(preview.upload_token, mapping);
      setResult(res);
      onImported();
    } catch (e: any) {
      setError(e.message ?? String(e));
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="import-page availability-import-content">
      {error && <div className="error-banner" role="alert">{error}</div>}

      <div className="availability-import-file-row">
        <button type="button" disabled={busy} onClick={() => void chooseFile()}>
          {preview ? "Choose another file" : "Choose spreadsheet"}
        </button>
        <ImportTemplateLinks template="accompanistLessons" disabled={busy} />
      </div>

      {preview && (
        <>
          <section>
            <h3>Column mapping</h3>
            <div className="clinical-import-mapping">
              {Object.entries(FIELD_LABELS).map(([field, label]) => (
                <label key={field}>
                  {label}
                  {REQUIRED_FIELDS.has(field) && <span className="required-star">*</span>}
                  <select
                    aria-label={label}
                    value={mapping[field] ?? ""}
                    onChange={(e) => setMapping({ ...mapping, [field]: e.target.value })}
                  >
                    <option value="">{REQUIRED_FIELDS.has(field) ? "Choose column" : "Not in file"}</option>
                    {preview.columns.map((c) => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </label>
              ))}
            </div>
          </section>

          <div className="availability-import-actions">
            <button type="button" disabled={busy} onClick={showPreview}>
              Preview Import
            </button>
          </div>

          {reviewing && (
          <section aria-label="Import preview" ref={previewRef}>
            <h3>Preview (first {preview.rows.length} rows)</h3>
            <div className="preview-table-wrap">
              <table className="preview-table">
                <thead>
                  <tr>
                    {preview.columns.map((c) => (
                      <th key={c}>{c}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {preview.rows.map((row, i) => (
                    <tr key={i}>
                      {preview.columns.map((c) => (
                        <td key={c}>{row[c]}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="availability-import-actions">
              <span className="muted">Importing replaces all existing lessons.</span>
              <button className="danger-btn" onClick={commit} disabled={busy}>
                {busy ? "Importing\u2026" : "Import lessons"}
              </button>
            </div>
          </section>
          )}
        </>
      )}

      {result && (
        <div className="availability-success" role="status">
          <p>
            Imported <strong>{result.created}</strong> lessons
            {result.skipped ? `, skipped ${result.skipped}` : ""}.
          </p>
          {result.warnings.length > 0 && (
            <ul className="warnings-list">
              {result.warnings.map((w, i) => (
                <li key={i}>{w}</li>
              ))}
            </ul>
          )}
        </div>
      )}
    </div>
  );
}
