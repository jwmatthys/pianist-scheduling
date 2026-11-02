import { formatMinutes } from "./time";
import { isLessonEntryPlaced, lessonPlacementState } from "./lessonPlacement";
import type { ReportTable } from "./reportTable";
import type { LessonInstructor, LessonSchedule, LessonScheduleEntry } from "./types";

const WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

export const UNPLACED_REASONS: Record<string, string> = {
  NO_FEASIBLE_CANDIDATE: "No feasible instructor",
  CAPACITY_LIMIT: "Capacity limit",
  INSTRUCTOR_OVERLAP: "Instructor overlap",
  MANUALLY_UNPLACED: "Manually unplaced",
};

const PLACEMENT_TYPES: Record<string, string> = {
  automatic: "Automatic",
  manual: "Manual",
  manuallyUnplaced: "Manually Unplaced",
};

export const REPORT_TITLE = "Lesson Schedule";
export const STUDENT_ROSTER_HEADING = "Student Placement Roster";
export const INSTRUCTOR_ROSTER_HEADING = "Instructor Placement Roster";
export const UNPLACED_HEADING = "Unplaced Students";
export const WARNINGS_HEADING = "Placement Warnings";
export const STALE_NOTE =
  "**Note:** Student or Instructor data has changed since this schedule was generated. Regenerate before relying on it.";

export interface LessonReport {
  markdown: string;
  pageBreakBeforeHeadings: string[];
}

export interface LessonReportOptions {
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

function compareByStudent(left: LessonScheduleEntry, right: LessonScheduleEntry) {
  return compareText(left.student_name, right.student_name);
}

function compareByTime(left: LessonScheduleEntry, right: LessonScheduleEntry) {
  return (
    weekdayIndex(left.day) - weekdayIndex(right.day) ||
    (left.start_minute ?? 0) - (right.start_minute ?? 0) ||
    compareByStudent(left, right)
  );
}

function isPlaced(entry: LessonScheduleEntry) {
  return isLessonEntryPlaced(entry);
}

function timeRange(entry: LessonScheduleEntry) {
  if (entry.start_minute === null || entry.end_minute === null) return "Time not set";
  return `${formatMinutes(entry.start_minute)}-${formatMinutes(entry.end_minute)}`;
}

function unplacedReason(entry: LessonScheduleEntry) {
  const code = lessonPlacementState(entry) === "manuallyUnplaced"
    ? "MANUALLY_UNPLACED"
    : entry.unplaced_reason ?? "NO_FEASIBLE_CANDIDATE";
  return UNPLACED_REASONS[code] ?? code;
}

function roomOf(instructor: LessonInstructor | undefined) {
  return instructor?.room || "";
}

export function buildLessonReport(
  schedule: LessonSchedule,
  instructors: LessonInstructor[],
  options: LessonReportOptions = {},
): LessonReport {
  const instructorsById = new Map(instructors.map((item) => [item.id, item]));
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
      const instructor = entry.instructor_id ? instructorsById.get(entry.instructor_id) : undefined;
      const room = roomOf(instructor);
      const manual = entry.placement_type === "manual" ? " (Manual)" : "";
      lines.push(
        `- **${entry.student_name}** - ${entry.instructor_name ?? "Instructor removed"}` +
          `${room ? ` - ${room}` : ""} - ${entry.day ?? ""} ${timeRange(entry)}${manual}`,
      );
    }
    lines.push("");
  }

  lines.push(`## ${INSTRUCTOR_ROSTER_HEADING}`, "");
  const byInstructor = new Map<string, { name: string; instructor?: LessonInstructor; entries: LessonScheduleEntry[] }>();
  for (const instructor of instructors) {
    byInstructor.set(instructor.id, { name: instructor.name, instructor, entries: [] });
  }
  for (const entry of placed) {
    const key = entry.instructor_id ?? entry.instructor_name ?? "";
    const group = byInstructor.get(key) ?? { name: entry.instructor_name ?? "Instructor removed", entries: [] };
    group.entries.push(entry);
    byInstructor.set(key, group);
  }
  if (byInstructor.size === 0) {
    lines.push("No instructors have been entered.", "");
  } else {
    for (const group of [...byInstructor.values()].sort((a, b) => compareText(a.name, b.name))) {
      const details = [
        group.instructor?.room ? `Room ${group.instructor.room}` : "",
        group.instructor ? `${group.instructor.lesson_length_minutes}-minute lessons` : "",
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
      const where = [item.instructor_name ?? "", item.day ?? "", item.start_minute === null ? "" : formatMinutes(item.start_minute)]
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
    pageBreakBeforeHeadings: [INSTRUCTOR_ROSTER_HEADING, UNPLACED_HEADING],
  };
}

/** Lesson-owned: one row per student with the Schedule tab's fields, for CSV and Excel. */
export function buildLessonReportTable(
  schedule: LessonSchedule,
  instructors: LessonInstructor[],
  options: { stale?: boolean } = {},
): ReportTable {
  const instructorsById = new Map(instructors.map((item) => [item.id, item]));
  const scheduleStatus = schedule.is_schedule_stale || options.stale === true
    ? "Stale: regenerate before relying on this schedule"
    : "Current";
  return {
    name: REPORT_TITLE,
    columns: [
      "Student", "Instructor", "Room", "Lesson Length (minutes)", "Day", "Start Time", "End Time",
      "Placement Type", "Preferred Tag Matches", "Tentative Minutes", "Status", "Warnings", "Unplaced Reason",
      "Schedule Status",
    ],
    rows: [...schedule.entries].sort(compareByStudent).map((entry) => {
      const placed = isPlaced(entry);
      const instructor = entry.instructor_id ? instructorsById.get(entry.instructor_id) : undefined;
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
        placed ? entry.instructor_name ?? "" : "",
        placed ? roomOf(instructor) : "",
        placed ? instructor?.lesson_length_minutes ?? null : null,
        placed ? entry.day ?? "" : "",
        placed && entry.start_minute !== null ? formatMinutes(entry.start_minute) : "",
        placed && entry.end_minute !== null ? formatMinutes(entry.end_minute) : "",
        PLACEMENT_TYPES[(placed || lessonPlacementState(entry) === "manuallyUnplaced" ? entry.placement_type : null) ?? ""] ?? "Unassigned",
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
