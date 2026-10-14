import { useState } from "react";
import ReactMarkdown from "react-markdown";
import { api } from "../../lib/api";
import { buildClinicalReport } from "../../lib/clinicalReport";
import { createMarkdownPdf } from "../../lib/markdownPdf";
import { saveBinaryFile, saveTextFile } from "../../lib/platform";

function errorText(error: unknown) {
  return error instanceof Error ? error.message : "The request could not be completed.";
}

export function ClinicalReportsPage() {
  const [markdown, setMarkdown] = useState<string | null>(null);
  const [pageBreaks, setPageBreaks] = useState<string[]>([]);
  const [busy, setBusy] = useState(false);
  const [saving, setSaving] = useState<"pdf" | "markdown" | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);

  async function generate() {
    setBusy(true);
    setError(null);
    setNotice(null);
    try {
      const [schedule, opportunities, students] = await Promise.all([
        api.getClinicalSchedule(),
        api.listClinicalOpportunities(),
        api.listClinicalStudents(),
      ]);
      const report = buildClinicalReport(schedule, opportunities, students, new Date());
      setMarkdown(report.markdown);
      setPageBreaks(report.pageBreakBeforeHeadings);
    } catch (cause) {
      setError(errorText(cause));
    } finally {
      setBusy(false);
    }
  }

  async function saveMarkdown() {
    if (!markdown) return;
    setSaving("markdown");
    setError(null);
    setNotice(null);
    try {
      const saved = await saveTextFile(markdown, "clinical_placements.md", "Markdown report", "md");
      setNotice(saved ? "Clinical Markdown report saved." : "Markdown save canceled.");
    } catch (cause) {
      setError(errorText(cause));
    } finally {
      setSaving(null);
    }
  }

  async function savePdf() {
    if (!markdown) return;
    setSaving("pdf");
    setError(null);
    setNotice(null);
    try {
      const pdf = await createMarkdownPdf(markdown, { pageBreakBeforeHeadings: pageBreaks });
      const saved = await saveBinaryFile(pdf, "clinical_placements.pdf", "PDF report", "pdf");
      setNotice(saved ? "Clinical PDF report saved." : "PDF save canceled.");
    } catch (cause) {
      setError(errorText(cause));
    } finally {
      setSaving(null);
    }
  }

  return (
    <section className="page reports-page clinical-reports-page" aria-labelledby="clinical-reports-heading">
      <h2 id="clinical-reports-heading">Clinical Reports</h2>
      <p className="muted">Generate a report from the current saved Clinical placement schedule.</p>
      <div className="reports-actions">
        <button className="primary-btn" type="button" onClick={() => void generate()} disabled={busy || saving !== null}>
          {busy ? "Generating…" : "Generate report"}
        </button>
        <button className="secondary-btn" type="button" onClick={() => void savePdf()} disabled={!markdown || saving !== null}>
          {saving === "pdf" ? "Saving PDF…" : "Download .pdf"}
        </button>
        <button className="secondary-btn" type="button" onClick={() => void saveMarkdown()} disabled={!markdown || saving !== null}>
          {saving === "markdown" ? "Saving Markdown…" : "Download .md"}
        </button>
      </div>
      {error && <p className="error-banner" role="alert">{error}</p>}
      {notice && <p className="availability-success" role="status">{notice}</p>}
      {markdown && <div className="markdown-preview"><ReactMarkdown>{markdown}</ReactMarkdown></div>}
    </section>
  );
}
