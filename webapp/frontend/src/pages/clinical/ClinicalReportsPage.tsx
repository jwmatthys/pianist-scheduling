import { useState } from "react";
import { useSavedReport, type ReportSourceState } from "../../lib/savedReport";
import { StaleReportBanner } from "../../components/StaleReportBanner";
import ReactMarkdown from "react-markdown";
import { api } from "../../lib/api";
import { buildClinicalReportTable } from "../../lib/clinicalReportTable";
import { buildClinicalReport } from "../../lib/clinicalReport";
import { ReportExportButton, type ExportStatus } from "../../components/ReportExport";

function errorText(error: unknown) {
  return error instanceof Error ? error.message : "The request could not be completed.";
}

async function fetchSourceState(): Promise<ReportSourceState> {
  const schedule = await api.getClinicalSchedule();
  return {
    stamp: `${schedule.schedule_revision}:${schedule.source_data_revision}`,
    sourceStale: schedule.is_schedule_stale,
  };
}

export function ClinicalReportsPage() {
  const { report, generate: saveGenerated } = useSavedReport("clinical", fetchSourceState);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);

  async function generate() {
    setBusy(true);
    setError(null);
    setNotice(null);
    try {
      await saveGenerated(async () => {
        const [schedule, opportunities, students] = await Promise.all([
          api.getClinicalSchedule(),
          api.listClinicalOpportunities(),
          api.listClinicalStudents(),
        ]);
        const built = buildClinicalReport(schedule, opportunities, students, new Date());
        return {
          markdown: built.markdown,
          pageBreakBeforeHeadings: built.pageBreakBeforeHeadings,
          table: buildClinicalReportTable(schedule, opportunities, students),
        };
      });
    } catch (cause) {
      setError(errorText(cause));
    } finally {
      setBusy(false);
    }
  }

  function handleExportStatus(status: ExportStatus) {
    setError(status.error ?? null);
    setNotice(status.notice ?? null);
  }

  return (
    <section className="page reports-page clinical-reports-page" aria-labelledby="clinical-reports-heading">
      <h2 id="clinical-reports-heading">Clinical Reports</h2>
      {report && <StaleReportBanner stale={report.stale} generatedAt={report.generatedAt} />}
      <p className="muted">Generate a report from the current saved Clinical placement schedule.</p>
      <div className="reports-actions">
        <button className="primary-btn" type="button" onClick={() => void generate()} disabled={busy}>
          {busy ? "Generating…" : "Generate report"}
        </button>
        <ReportExportButton
          reportId="clinical.placements"
          title="Clinical Placements"
          payload={report?.payload ?? null}
          onStatus={handleExportStatus}
        />
      </div>
      {error && <p className="error-banner" role="alert">{error}</p>}
      {notice && <p className="availability-success" role="status">{notice}</p>}
      {report && <div className="markdown-preview"><ReactMarkdown>{report.payload.markdown}</ReactMarkdown></div>}
    </section>
  );
}
