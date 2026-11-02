import { useState } from "react";
import ReactMarkdown from "react-markdown";
import { api } from "../../lib/api";
import { useSavedReport, type ReportSourceState } from "../../lib/savedReport";
import { StaleReportBanner } from "../../components/StaleReportBanner";
import { ReportExportButton, type ExportStatus } from "../../components/ReportExport";
import { REPORT_TITLE, buildLessonReport, buildLessonReportTable } from "../../lib/lessonReport";

function errorText(error: unknown) {
  return error instanceof Error ? error.message : "The request could not be completed.";
}

async function fetchSourceState(): Promise<ReportSourceState> {
  const [schedule, moduleState] = await Promise.all([api.getLessonSchedule(), api.getLessonModuleState()]);
  return {
    stamp: `${schedule.schedule_revision}:${schedule.source_data_revision}`,
    sourceStale: schedule.is_schedule_stale || schedule.source_data_revision !== moduleState.source_data_revision,
  };
}

export function LessonReportsPage() {
  const { report, generate: saveGenerated } = useSavedReport("lesson", fetchSourceState);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);

  async function generate() {
    setBusy(true);
    setError(null);
    setNotice(null);
    try {
      await saveGenerated(async () => {
        const [schedule, moduleState, instructors] = await Promise.all([
          api.getLessonSchedule(),
          api.getLessonModuleState(),
          api.listLessonInstructors(),
        ]);
        const stale = schedule.is_schedule_stale || schedule.source_data_revision !== moduleState.source_data_revision;
        const built = buildLessonReport(schedule, instructors, { generatedAt: new Date(), stale });
        return {
          markdown: built.markdown,
          pageBreakBeforeHeadings: built.pageBreakBeforeHeadings,
          table: buildLessonReportTable(schedule, instructors, { stale }),
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
    <section className="page reports-page lesson-reports-page" aria-labelledby="lesson-reports-heading">
      <h2 id="lesson-reports-heading">Lesson Reports</h2>
      {report && <StaleReportBanner stale={report.stale} generatedAt={report.generatedAt} />}
      <p className="muted">Generate a report from the current saved Lesson schedule.</p>
      <div className="reports-actions">
        <button className="primary-btn" type="button" onClick={() => void generate()} disabled={busy}>
          {busy ? "Generating…" : "Generate report"}
        </button>
        <ReportExportButton
          reportId="lesson.placements"
          title={REPORT_TITLE}
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
