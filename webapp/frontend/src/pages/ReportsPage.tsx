import { useRef, useState } from "react";
import ReactMarkdown from "react-markdown";
import { api } from "../lib/api";

export function ReportsPage() {
  const [markdown, setMarkdown] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [downloadingPdf, setDownloadingPdf] = useState(false);
  const reportRef = useRef<HTMLDivElement>(null);

  async function generate() {
    setBusy(true);
    try {
      const text = await api.markdownReport();
      setMarkdown(text);
    } finally {
      setBusy(false);
    }
  }

  function download() {
    if (!markdown) return;
    const blob = new Blob([markdown], { type: "text/markdown" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "lesson_pianists.md";
    a.click();
    URL.revokeObjectURL(url);
  }

  async function downloadPdf() {
    if (!markdown || !reportRef.current) return;
    setDownloadingPdf(true);
    try {
      const { jsPDF } = await import("jspdf");
      const pdf = new jsPDF({ format: "letter", unit: "pt" });
      await pdf.html(reportRef.current, {
        autoPaging: "text",
        margin: [42, 42, 42, 42],
        width: 528,
        windowWidth: reportRef.current.scrollWidth,
      });
      pdf.save("lesson_pianists.pdf");
    } finally {
      setDownloadingPdf(false);
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
        <button className="secondary-btn" onClick={downloadPdf} disabled={!markdown || downloadingPdf}>
          {downloadingPdf ? "Preparing PDF\u2026" : "Download .pdf"}
        </button>
        <button className="secondary-btn" onClick={download} disabled={!markdown}>
          Download .md
        </button>
      </div>
      {markdown && (
        <div ref={reportRef} className="markdown-preview">
          <ReactMarkdown>{markdown}</ReactMarkdown>
        </div>
      )}
    </div>
  );
}
