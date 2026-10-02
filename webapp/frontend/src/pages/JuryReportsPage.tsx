import { useState } from "react";
import ReactMarkdown from "react-markdown";
import { api } from "../lib/api";
import { createMarkdownPdf } from "../lib/markdownPdf";
import { saveBinaryFile, saveTextFile } from "../lib/platform";
import { formatUtcTimestampLocally } from "../lib/time";
import type { JuryScheduleEvent, JuryScheduleResult } from "../lib/types";

function formatTime(value: number) {
  if (value === 0 || value === 1440) return "12:00 AM";
  const hour = Math.floor(value / 60);
  const minute = value % 60;
  return `${hour % 12 || 12}:${String(minute).padStart(2, "0")} ${hour < 12 ? "AM" : "PM"}`;
}

function formatDate(value: string) {
  return new Date(`${value}T12:00:00`).toLocaleDateString(undefined, {
    month: "long",
    day: "numeric",
    year: "numeric",
  });
}

function eventMarkdown(event: JuryScheduleEvent, entriesByLesson: Map<string, { teacher: string }>) {
  const time = `${formatTime(event.start_minute)}-${formatTime(event.end_minute)}`;
  if (event.kind === "meal_break") return `- **${time}** - MEAL BREAK`;
  if (event.kind === "periodic_break") return `- **${time}** - BREAK`;

  const lesson = entriesByLesson.get(event.source_lesson_uuid);
  const instructor = lesson?.teacher || "Not supplied";
  return `- **${time}** - ${event.student_display_name || "Unnamed student"} - ${event.instrument || "Instrument not supplied"} - Pianist: ${event.pianist_display_name || "None"} - Instructor: ${instructor}`;
}

function reportMarkdown(
  result: JuryScheduleResult,
  entries: Awaited<ReturnType<typeof api.getJuryEntries>>,
) {
  const entriesByLesson = new Map(entries.map((entry) => [entry.source_lesson_uuid, entry]));
  const lines = [
    "# Jury Schedule",
    "",
    `Generated: ${formatUtcTimestampLocally(result.created_at)}`,
  ];
  const pageBreakBeforeHeadings: string[] = [];

  lines.push("");

  for (const [index, panel] of result.panel_timelines.entries()) {
    const heading = `${panel.panel_name} - ${formatDate(panel.jury_date)}`;
    lines.push(`## ${heading}`, "");
    if (index > 0) pageBreakBeforeHeadings.push(heading);
    if (panel.events.length === 0) {
      lines.push("No scheduled entries.", "");
      continue;
    }
    for (const event of panel.events) {
      lines.push(eventMarkdown(event, entriesByLesson));
    }
    lines.push("");
  }

  const pianistSchedules = new Map<string, string[]>();
  for (const panel of result.panel_timelines) {
    for (const event of panel.events) {
      if (event.kind !== "jury") continue;
      const pianistName = event.pianist_display_name || "No Pianist Assigned";
      const eventLine = `- **${formatTime(event.start_minute)}-${formatTime(event.end_minute)}** - ${event.student_display_name || "Unnamed student"} - ${event.instrument || "Instrument not supplied"} - ${panel.panel_name} - ${formatDate(panel.jury_date)}`;
      const events = pianistSchedules.get(pianistName) ?? [];
      events.push(eventLine);
      pianistSchedules.set(pianistName, events);
    }
  }

  const pianistHeading = "Schedule by Pianist";
  pageBreakBeforeHeadings.push(pianistHeading);
  lines.push(`## ${pianistHeading}`, "");
  if (pianistSchedules.size === 0) {
    lines.push("No scheduled Jury entries.", "");
  } else {
    for (const [pianistName, events] of [...pianistSchedules.entries()].sort(([left], [right]) => left.localeCompare(right))) {
      lines.push(`### ${pianistName}`, ...events, "");
    }
  }

  if (result.unscheduled_lessons.length > 0) {
    lines.push("## Unscheduled Lessons", "");
    for (const lesson of result.unscheduled_lessons) {
      const panel = result.panel_timelines.find((item) => item.panel_uuid === lesson.panel_uuid);
      lines.push(
        `- ${lesson.student_display_name || "Unnamed student"} - ${lesson.instrument} - ${panel?.panel_name ?? "Panel unavailable"} - ${lesson.reason_code}: ${lesson.explanation}`,
      );
    }
  }

  return { markdown: lines.join("\n").trimEnd() + "\n", pageBreakBeforeHeadings };
}

function errorText(error: unknown) {
  return error instanceof Error ? error.message : "The request could not be completed.";
}

export function JuryReportsPage() {
  const [markdown, setMarkdown] = useState<string | null>(null);
  const [panelPageBreaks, setPanelPageBreaks] = useState<string[]>([]);
  const [busy, setBusy] = useState(false);
  const [saving, setSaving] = useState<"pdf" | "markdown" | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);

  async function generate() {
    setBusy(true);
    setError(null);
    setNotice(null);
    try {
      const [result, entries] = await Promise.all([
        api.getCurrentJurySchedule(),
        api.getJuryEntries(),
      ]);
      const report = reportMarkdown(result, entries);
      setMarkdown(report.markdown);
      setPanelPageBreaks(report.pageBreakBeforeHeadings);
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
      const saved = await saveTextFile(markdown, "jury_schedule.md", "Markdown report", "md");
      setNotice(saved ? "Jury Markdown report saved." : "Markdown save canceled.");
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
      const pdf = await createMarkdownPdf(markdown, { pageBreakBeforeHeadings: panelPageBreaks });
      const saved = await saveBinaryFile(pdf, "jury_schedule.pdf", "PDF report", "pdf");
      setNotice(saved ? "Jury PDF report saved." : "PDF save canceled.");
    } catch (cause) {
      setError(errorText(cause));
    } finally {
      setSaving(null);
    }
  }

  return (
    <section className="page reports-page jury-reports-page" aria-labelledby="jury-reports-title">
      <h2 id="jury-reports-title">Jury Reports</h2>
      <p className="muted">Generate a report from the current saved Jury schedule.</p>
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
