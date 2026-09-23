import { useState } from "react";
import { api } from "../lib/api";

export function ReportsPage() {
  const [markdown, setMarkdown] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

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
        <button className="secondary-btn" onClick={download} disabled={!markdown}>
          Download .md
        </button>
      </div>
      {markdown && <pre className="markdown-preview">{markdown}</pre>}
    </div>
  );
}
