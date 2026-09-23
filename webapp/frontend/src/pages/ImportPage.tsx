import { useState } from "react";
import { api } from "../lib/api";
import type { ImportPreview } from "../lib/types";

const FIELD_LABELS: Record<string, string> = {
  teacher: "Teacher",
  student: "Student",
  day: "Day",
  start_time: "Start time",
  end_time: "End time",
  room: "Room",
  instrument: "Instrument",
  required_pianist_name: "Required pianist (optional)",
  need_pianist: "Needs pianist? (optional)",
};

const REQUIRED_FIELDS = new Set(["day", "start_time"]);

function guessColumn(field: string, columns: string[]): string {
  const normalized = field.replace(/_/g, " ");
  const found = columns.find((c) => c.toLowerCase().includes(normalized.split(" ")[0]));
  return found ?? "";
}

export function ImportPage({ onImported }: { onImported: () => void }) {
  const [preview, setPreview] = useState<ImportPreview | null>(null);
  const [mapping, setMapping] = useState<Record<string, string>>({});
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<{ created: number; skipped: number; warnings: string[] } | null>(null);
  const [error, setError] = useState<string | null>(null);

  async function handleFile(file: File) {
    setError(null);
    setResult(null);
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

  async function commit() {
    if (!preview) return;
    const missing = Array.from(REQUIRED_FIELDS).filter((f) => !mapping[f]);
    if (missing.length) {
      setError(`Please map required fields: ${missing.map((f) => FIELD_LABELS[f]).join(", ")}`);
      return;
    }
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
    <div className="page import-page">
      <h2>Import lessons</h2>
      <p className="muted">
        Upload a CSV or XLSX file of lessons, then tell us which column maps to which field.
      </p>
      <input
        type="file"
        accept=".csv,.xlsx,.xls"
        onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])}
      />
      {error && <p className="error-banner">{error}</p>}

      {preview && (
        <>
          <h3>Column mapping</h3>
          <table className="mapping-table">
            <tbody>
              {Object.entries(FIELD_LABELS).map(([field, label]) => (
                <tr key={field}>
                  <td>
                    {label}
                    {REQUIRED_FIELDS.has(field) && <span className="required-star">*</span>}
                  </td>
                  <td>
                    <select
                      value={mapping[field] ?? ""}
                      onChange={(e) => setMapping({ ...mapping, [field]: e.target.value })}
                    >
                      <option value="">-- not mapped --</option>
                      {preview.columns.map((c) => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>

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

          <button className="primary-btn" onClick={commit} disabled={busy}>
            {busy ? "Importing\u2026" : "Import lessons"}
          </button>
        </>
      )}

      {result && (
        <div className="import-result">
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
