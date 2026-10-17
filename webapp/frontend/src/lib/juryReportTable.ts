import { formatMinutes } from "./time";
import type { ReportCell, ReportTable } from "./reportTable";
import type { JuryLessonEntry, JuryScheduleResult } from "./types";

function titleCase(value: string) {
  return value.toLowerCase().replaceAll("_", " ").replace(/\b\w/g, (letter) => letter.toUpperCase());
}

/**
 * Jury-owned: the generated Schedule tab, one row per Panel entry (jury or break) followed by
 * one row per unscheduled lesson, using the tab's columns.
 */
export function buildJuryReportTable(result: JuryScheduleResult, entries: JuryLessonEntry[]): ReportTable {
  const entriesByLesson = new Map(entries.map((entry) => [entry.source_lesson_uuid, entry]));
  const rows: ReportCell[][] = [];
  const panelNames = new Map(result.panel_timelines.map((panel) => [panel.panel_uuid, panel.panel_name]));
  for (const panel of result.panel_timelines) {
    for (const event of panel.events) {
      const times = [formatMinutes(event.start_minute), formatMinutes(event.end_minute)];
      if (event.kind === "jury") {
        const entry = entriesByLesson.get(event.source_lesson_uuid);
        rows.push([
          panel.panel_name, panel.jury_date, ...times, event.student_display_name || "Unnamed student",
          event.instrument, entry?.teacher ?? "", event.pianist_display_name ?? "", entry?.pianist_required ?? null,
          "Scheduled", "", "",
        ]);
      } else {
        rows.push([
          panel.panel_name, panel.jury_date, ...times, event.kind === "meal_break" ? "MEAL BREAK" : "BREAK",
          "", "", "", null, "Scheduled", "", "",
        ]);
      }
    }
  }
  for (const lesson of result.unscheduled_lessons) {
    const entry = entriesByLesson.get(lesson.source_lesson_uuid);
    rows.push([
      panelNames.get(lesson.panel_uuid) ?? "Panel unavailable", lesson.jury_date, "", "",
      lesson.student_display_name || "Unnamed student", lesson.instrument, entry?.teacher ?? "",
      entry?.assigned_pianist?.display_name ?? "", entry?.pianist_required ?? null, "Unscheduled",
      titleCase(lesson.reason_code), lesson.explanation,
    ]);
  }
  return {
    name: "Jury Schedule",
    columns: [
      "Panel", "Jury Date", "Start", "End", "Student / entry", "Instrument", "Instructor", "Assigned Pianist",
      "Pianist Required", "Status", "Reason", "Explanation",
    ],
    rows,
  };
}
