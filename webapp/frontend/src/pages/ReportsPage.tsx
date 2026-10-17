import { useState } from "react";
import { useSavedReport, type ReportSourceState } from "../lib/savedReport";
import { StaleReportBanner } from "../components/StaleReportBanner";
import ReactMarkdown from "react-markdown";
import { api } from "../lib/api";
import { buildAccompanistReportTable } from "../lib/accompanistReportTable";
import { ReportExportButton, type ExportStatus } from "../components/ReportExport";

async function fetchSourceState(): Promise<ReportSourceState> {
  const state = await api.getAccompanistFinalizationState();
  return { stamp: String(state.source_revision), sourceStale: false };
}

export function ReportsPage() {
  const { report, generate: saveGenerated } = useSavedReport("accompanists", fetchSourceState);
  const [busy, setBusy] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);
  const [exportMessage, setExportMessage] = useState<string | null>(null);

  async function generate() {
    setBusy(true);
    setExportError(null);
    try {
      await saveGenerated(async () => {
        const [markdown, lessons, pianists] = await Promise.all([
          api.markdownReport(),
          api.listLessons(),
          api.listPianists(),
        ]);
        return { markdown, table: buildAccompanistReportTable(lessons, pianists) };
      });
    } catch (cause) {
      setExportError(cause instanceof Error ? cause.message : "The report could not be generated.");
    } finally {
      setBusy(false);
    }
  }

  function handleExportStatus(status: ExportStatus) {
    setExportError(status.error ?? null);
    setExportMessage(status.notice ?? null);
  }

  return (
    <div className="page reports-page">
      <h2>Reports</h2>
      {report && <StaleReportBanner stale={report.stale} generatedAt={report.generatedAt} />}
      <p className="muted">
        Generate a report from the current saved lesson roster.
      </p>
      <div className="reports-actions">
        <button className="primary-btn" onClick={generate} disabled={busy}>
          {busy ? "Generating\u2026" : "Generate report"}
        </button>
        <ReportExportButton
          reportId="accompanist.lesson-pianists"
          title="Accompanist Schedule"
          payload={report?.payload ?? null}
          onStatus={handleExportStatus}
        />
      </div>
      {exportError && <p className="error-banner" role="alert">{exportError}</p>}
      {exportMessage && <p className="availability-success" role="status">{exportMessage}</p>}
      {report && (
        <div className="markdown-preview">
          <ReactMarkdown>{report.payload.markdown}</ReactMarkdown>
        </div>
      )}
    </div>
  );
}
