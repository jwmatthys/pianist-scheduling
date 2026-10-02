import { useState } from "react";
import ReactMarkdown from "react-markdown";
import { api } from "../lib/api";
import { saveBinaryFile, saveTextFile } from "../lib/platform";
import { createMarkdownPdf } from "../lib/markdownPdf";

export function ReportsPage() {
  const [markdown, setMarkdown] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [saving, setSaving] = useState<"pdf" | "markdown" | null>(null);
  const [exportError, setExportError] = useState<string | null>(null);
  const [exportMessage, setExportMessage] = useState<string | null>(null);

  async function generate() {
    setBusy(true);
    try {
      const text = await api.markdownReport();
      setMarkdown(text);
    } finally {
      setBusy(false);
    }
  }

  async function downloadMarkdown() {
    if (!markdown) return;
    setSaving("markdown");
    setExportError(null);
    setExportMessage(null);
    try {
      const saved = await saveTextFile(markdown, "lesson_pianists.md", "Markdown report", "md");
      setExportMessage(saved ? "Markdown report saved." : "Markdown save canceled.");
    } catch (error) {
      setExportError(error instanceof Error ? error.message : "Could not save the Markdown report.");
    } finally {
      setSaving(null);
    }
  }

  async function downloadPdf() {
    if (!markdown) return;
    setSaving("pdf");
    setExportError(null);
    setExportMessage(null);
    try {
      const pdf = await createMarkdownPdf(markdown);
      const saved = await saveBinaryFile(pdf, "lesson_pianists.pdf", "PDF report", "pdf");
      setExportMessage(saved ? "PDF report saved." : "PDF save canceled.");
    } catch (error) {
      setExportError(error instanceof Error ? error.message : "Could not generate or save the PDF report.");
    } finally {
      setSaving(null);
    }
  }

  return (
    <div className="page reports-page">
      <h2>Reports</h2>
      <p className="muted">
        Generate the same pianist / instructor / student Markdown schedule as
        <code> generate_lesson_markdown.py</code>, based on the current assignments.
      </p>
      <div className="reports-actions">
        <button className="primary-btn" onClick={generate} disabled={busy}>
          {busy ? "Generating\u2026" : "Generate report"}
        </button>
        <button className="secondary-btn" onClick={() => void downloadPdf()} disabled={!markdown || saving !== null}>
          {saving === "pdf" ? "Saving PDF…" : "Download .pdf"}
        </button>
        <button className="secondary-btn" onClick={() => void downloadMarkdown()} disabled={!markdown || saving !== null}>
          {saving === "markdown" ? "Saving Markdown…" : "Download .md"}
        </button>
      </div>
      {exportError && <p className="error-banner" role="alert">{exportError}</p>}
      {exportMessage && <p className="availability-success" role="status">{exportMessage}</p>}
      {markdown && (
        <div className="markdown-preview">
          <ReactMarkdown>{markdown}</ReactMarkdown>
        </div>
      )}
    </div>
  );
}
