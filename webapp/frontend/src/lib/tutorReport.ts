import { formatMinutes } from "./time";
import { isTutorEntryPlaced, tutorPlacementState } from "./tutorPlacement";
import type { ReportTable } from "./reportTable";
import type { TutorTutor, TutorSchedule, TutorScheduleEntry } from "./types";

const WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

export const UNPLACED_REASONS: Record<string, string> = {
  NO_FEASIBLE_CANDIDATE: "No feasible tutor",
  CAPACITY_LIMIT: "Capacity limit",
  TUTOR_OVERLAP: "Tutor overlap",
  MANUALLY_UNPLACED: "Manually unplaced",
};

const PLACEMENT_TYPES: Record<string, string> = {
  automatic: "Automatic",
  manual: "Manual",
  manuallyUnplaced: "Manually Unplaced",
};

export const REPORT_TITLE = "Tutor Schedule";
export const STUDENT_ROSTER_HEADING = "Student Placement Roster";
export const TUTOR_ROSTER_HEADING = "Tutor Placement Roster";
export const UNPLACED_HEADING = "Unplaced Students";
export const WARNINGS_HEADING = "Placement Warnings";
export const STALE_NOTE =
  "**Note:** Student or Tutor data has changed since this schedule was generated. Regenerate before relying on it.";

export interface TutorReport {
  markdown: string;
  pageBreakBeforeHeadings: string[];
}

export interface TutorReportOptions {
  generatedAt?: Date;
  /** True when the module state shows source data newer than the schedule. */
  stale?: boolean;
}

function weekdayIndex(day: string | null): number {
  const index = WEEKDAYS.indexOf(day ?? "");
  return index === -1 ? WEEKDAYS.length : index;
}

function compareText(left: string, right: string) {
  return left.localeCompare(right, undefined, { sensitivity: "base" }) || (left < right ? -1 : left > right ? 1 : 0);
}

function compareByStudent(left: TutorScheduleEntry, right: TutorScheduleEntry) {
  return compareText(left.student_name, right.student_name);
}

function compareByTime(left: TutorScheduleEntry, right: TutorScheduleEntry) {
  return (
    weekdayIndex(left.day) - weekdayIndex(right.day) ||
    (left.start_minute ?? 0) - (right.start_minute ?? 0) ||
    compareByStudent(left, right)
  );
}

function isPlaced(entry: TutorScheduleEntry) {
  return isTutorEntryPlaced(entry);
}

function timeRange(entry: TutorScheduleEntry) {
  if (entry.start_minute === null || entry.end_minute === null) return "Time not set";
  return `${formatMinutes(entry.start_minute)}-${formatMinutes(entry.end_minute)}`;
}

function unplacedReason(entry: TutorScheduleEntry) {
  const code = tutorPlacementState(entry) === "manuallyUnplaced"
    ? "MANUALLY_UNPLACED"
    : entry.unplaced_reason ?? "NO_FEASIBLE_CANDIDATE";
  return UNPLACED_REASONS[code] ?? code;
}

function roomOf(tutor: TutorTutor | undefined) {
  return tutor?.room || "";
}

export function buildTutorReport(
  schedule: TutorSchedule,
  tutors: TutorTutor[],
  options: TutorReportOptions = {},
): TutorReport {
  const tutorsById = new Map(tutors.map((item) => [item.id, item]));
  const placed = schedule.entries.filter(isPlaced);
  const unplaced = schedule.entries.filter((entry) => !isPlaced(entry));
  const stale = schedule.is_schedule_stale || options.stale === true;
  const lines: string[] = [`# ${REPORT_TITLE}`, ""];

  if (options.generatedAt) lines.push(`Generated: ${options.generatedAt.toLocaleString()}`, "");
  lines.push(
    `${placed.length} ${placed.length === 1 ? "student" : "students"} placed, ${unplaced.length} unplaced.`,
    "",
  );
  if (stale) lines.push(STALE_NOTE, "");

  lines.push(`## ${STUDENT_ROSTER_HEADING}`, "");
  if (placed.length === 0) {
    lines.push("No students have been placed.", "");
  } else {
    for (const entry of [...placed].sort(compareByStudent)) {
      const tutor = entry.tutor_id ? tutorsById.get(entry.tutor_id) : undefined;
      const room = roomOf(tutor);
      const manual = entry.placement_type === "manual" ? " (Manual)" : "";
      lines.push(
        `- **${entry.student_name}** - ${entry.tutor_name ?? "Tutor removed"}` +
          `${room ? ` - ${room}` : ""} - ${entry.day ?? ""} ${timeRange(entry)}${manual}`,
      );
    }
    lines.push("");
  }

  lines.push(`## ${TUTOR_ROSTER_HEADING}`, "");
  const byTutor = new Map<string, { name: string; tutor?: TutorTutor; entries: TutorScheduleEntry[] }>();
  for (const tutor of tutors) {
    byTutor.set(tutor.id, { name: tutor.name, tutor, entries: [] });
  }
  for (const entry of placed) {
    const key = entry.tutor_id ?? entry.tutor_name ?? "";
    const group = byTutor.get(key) ?? { name: entry.tutor_name ?? "Tutor removed", entries: [] };
    group.entries.push(entry);
    byTutor.set(key, group);
  }
  if (byTutor.size === 0) {
    lines.push("No tutors have been entered.", "");
  } else {
    for (const group of [...byTutor.values()].sort((a, b) => compareText(a.name, b.name))) {
      const details = [
        group.tutor?.room ? `Room ${group.tutor.room}` : "",
        group.tutor ? `${group.tutor.lesson_length_minutes}-minute sessions` : "",
      ].filter(Boolean);
      lines.push(`### ${group.name}${details.length ? ` - ${details.join(", ")}` : ""}`, "");
      if (group.entries.length === 0) {
        lines.push("No students assigned.", "");
        continue;
      }
      for (const entry of [...group.entries].sort(compareByTime)) {
        const manual = entry.placement_type === "manual" ? " (Manual)" : "";
        lines.push(`- **${entry.day ?? ""} ${timeRange(entry)}** - ${entry.student_name}${manual}`);
      }
      lines.push("");
    }
  }

  lines.push(`## ${UNPLACED_HEADING}`, "");
  if (unplaced.length === 0) {
    lines.push("All students have a placement.", "");
  } else {
    for (const entry of [...unplaced].sort(compareByStudent)) {
      lines.push(`- **${entry.student_name}** - ${unplacedReason(entry)}`);
    }
    lines.push("");
  }

  lines.push(`## ${WARNINGS_HEADING}`, "");
  if (stale) lines.push("The schedule is stale: source data has changed since it was generated.", "");
  const warned = schedule.verification.filter((item) => item.warnings.length > 0);
  if (warned.length === 0) {
    lines.push("No placement warnings.", "");
  } else {
    const manualByStudent = new Map(placed.map((entry) => [entry.student_id, entry.placement_type === "manual"]));
    for (const item of [...warned].sort((a, b) => compareText(a.student_name, b.student_name))) {
      const where = [item.tutor_name ?? "", item.day ?? "", item.start_minute === null ? "" : formatMinutes(item.start_minute)]
        .filter(Boolean)
        .join(" ");
      const manual = manualByStudent.get(item.student_id) ? " (Manual override)" : "";
      lines.push(`- **${item.student_name}**${where ? ` - ${where}` : ""}${manual}`);
      for (const warning of item.warnings) lines.push(`  - ${warning.message}`);
    }
    lines.push("");
  }

  return {
    markdown: lines.join("\n").trimEnd() + "\n",
    pageBreakBeforeHeadings: [TUTOR_ROSTER_HEADING, UNPLACED_HEADING],
  };
}

/** Tutor-owned: one row per student with the Schedule tab's fields, for CSV and Excel. */
export function buildTutorReportTable(
  schedule: TutorSchedule,
  tutors: TutorTutor[],
  options: { stale?: boolean } = {},
): ReportTable {
  const tutorsById = new Map(tutors.map((item) => [item.id, item]));
  const scheduleStatus = schedule.is_schedule_stale || options.stale === true
    ? "Stale: regenerate before relying on this schedule"
    : "Current";
  return {
    name: REPORT_TITLE,
    columns: [
      "Student", "Tutor", "Room", "Lesson Length (minutes)", "Day", "Start Time", "End Time",
      "Placement Type", "Preferred Tag Matches", "Tentative Minutes", "Status", "Warnings", "Unplaced Reason",
      "Schedule Status",
    ],
    rows: [...schedule.entries].sort(compareByStudent).map((entry) => {
      const placed = isPlaced(entry);
      const tutor = entry.tutor_id ? tutorsById.get(entry.tutor_id) : undefined;
      let status = "";
      if (placed) {
        if (entry.warnings.length > 0) {
          status = entry.warnings.every((warning) => warning.code === "TENTATIVE_AVAILABILITY")
            ? "Placement warning"
            : "Placement conflict";
        } else {
          status = entry.verified ? "Verified" : "Needs review";
        }
      }
      return [
        entry.student_name,
        placed ? entry.tutor_name ?? "" : "",
        placed ? roomOf(tutor) : "",
        placed ? tutor?.lesson_length_minutes ?? null : null,
        placed ? entry.day ?? "" : "",
        placed && entry.start_minute !== null ? formatMinutes(entry.start_minute) : "",
        placed && entry.end_minute !== null ? formatMinutes(entry.end_minute) : "",
        PLACEMENT_TYPES[(placed || tutorPlacementState(entry) === "manuallyUnplaced" ? entry.placement_type : null) ?? ""] ?? "Unassigned",
        placed ? entry.preferred_tag_match_count : null,
        placed ? entry.tentative_minutes : null,
        status,
        entry.warnings.map((warning) => warning.message).join("; "),
        placed ? "" : unplacedReason(entry),
        scheduleStatus,
      ];
    }),
  };
}
