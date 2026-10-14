import { formatMinutes } from "./time";
import type {
  ClinicalSchedule,
  ClinicalScheduleEntry,
  ClinicalStudent,
  PlacementOpportunity,
} from "./types";

const WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const UNPLACED_REASONS: Record<string, string> = {
  NO_FEASIBLE_CANDIDATE: "No feasible opportunity",
  CAPACITY_LIMIT: "Capacity limit",
  MANUALLY_UNPLACED: "Manually unplaced",
};

export const STUDENT_ROSTER_HEADING = "Student Placement Roster";
export const SITE_ROSTER_HEADING = "Site Placement Roster";
export const UNPLACED_HEADING = "Unplaced Students";
export const WARNINGS_HEADING = "Placement Warnings";

export interface ClinicalReport {
  markdown: string;
  pageBreakBeforeHeadings: string[];
}

function weekdayIndex(day: string | null): number {
  const index = WEEKDAYS.indexOf(day ?? "");
  return index === -1 ? WEEKDAYS.length : index;
}

function compareText(left: string, right: string) {
  return left.localeCompare(right, undefined, { sensitivity: "base" }) || (left < right ? -1 : left > right ? 1 : 0);
}

function compareByStudent(left: ClinicalScheduleEntry, right: ClinicalScheduleEntry) {
  return compareText(left.student_name, right.student_name) || compareText(left.student_number, right.student_number);
}

function timeRange(entry: ClinicalScheduleEntry) {
  if (entry.start_minute === null || entry.end_minute === null) return "Time not set";
  return `${formatMinutes(entry.start_minute)}-${formatMinutes(entry.end_minute)}`;
}

function studentLabel(entry: ClinicalScheduleEntry, email: string | undefined, includeId = true) {
  const parts = [entry.student_name];
  if (includeId && entry.student_number) parts.push(`ID ${entry.student_number}`);
  if (email) parts.push(email);
  return parts.join(" - ");
}

function siteContact(opportunity: PlacementOpportunity | undefined) {
  if (!opportunity) return "";
  const parts = [opportunity.contact_name, opportunity.contact_email, opportunity.contact_phone].filter(Boolean);
  return parts.length ? `Contact: ${parts.join(", ")}` : "";
}

export function buildClinicalReport(
  schedule: ClinicalSchedule,
  opportunities: PlacementOpportunity[],
  students: ClinicalStudent[],
  generatedAt?: Date,
): ClinicalReport {
  const opportunitiesById = new Map(opportunities.map((item) => [item.id, item]));
  const emailByStudent = new Map(students.map((student) => [student.id, student.email]));
  const placed = schedule.entries.filter((entry) => entry.kind === "automatic" || entry.kind === "manual");
  const unplaced = schedule.entries.filter((entry) => entry.kind === null || entry.kind === "manuallyUnplaced");
  const lines: string[] = ["# Clinical Placement Schedule", ""];

  if (generatedAt) lines.push(`Generated: ${generatedAt.toLocaleString()}`, "");
  lines.push(
    `${placed.length} ${placed.length === 1 ? "student" : "students"} placed, ` +
      `${unplaced.length} unplaced.`,
    "",
  );
  if (schedule.is_schedule_stale) {
    lines.push(
      "**Note:** Student or Clinical Site data has changed since this schedule was generated. Regenerate before relying on it.",
      "",
    );
  }

  lines.push(`## ${STUDENT_ROSTER_HEADING}`, "");
  if (placed.length === 0) {
    lines.push("No students have been placed.", "");
  } else {
    for (const entry of [...placed].sort(compareByStudent)) {
      const opportunity = entry.opportunity_id ? opportunitiesById.get(entry.opportunity_id) : undefined;
      lines.push(
        `- **${studentLabel(entry, emailByStudent.get(entry.student_id))}** - ${entry.day ?? ""} ${timeRange(entry)}`,
      );
      const site = [
        entry.site_name ?? "Site unavailable",
        opportunity?.site_address,
        opportunity?.contact_name,
        opportunity?.contact_email,
        opportunity?.contact_phone,
      ].filter(Boolean);
      lines.push(`  - ${site.join(", ")}`);
    }
    lines.push("");
  }

  lines.push(`# ${SITE_ROSTER_HEADING}`, "");
  if (placed.length === 0) {
    lines.push("No students have been placed.", "");
  } else {
    const bySite = new Map<string, Map<string, ClinicalScheduleEntry[]>>();
    for (const entry of placed) {
      const siteName = entry.site_name ?? "Site unavailable";
      const siteGroups = bySite.get(siteName) ?? new Map<string, ClinicalScheduleEntry[]>();
      const key = entry.opportunity_id ?? "";
      siteGroups.set(key, [...(siteGroups.get(key) ?? []), entry]);
      bySite.set(siteName, siteGroups);
    }
    for (const siteName of [...bySite.keys()].sort(compareText)) {
      lines.push(`## ${siteName}`, "");
      const groups = [...(bySite.get(siteName) ?? new Map()).entries()].sort(([leftId, left], [rightId, right]) => {
        const a = opportunitiesById.get(leftId);
        const b = opportunitiesById.get(rightId);
        return (
          weekdayIndex(a?.day_of_week ?? left[0].day) - weekdayIndex(b?.day_of_week ?? right[0].day) ||
          (a?.start_minute ?? 0) - (b?.start_minute ?? 0) ||
          compareText(leftId, rightId)
        );
      });
      for (const [opportunityId, group] of groups) {
        const opportunity = opportunitiesById.get(opportunityId);
        const day = opportunity?.day_of_week ?? group[0].day ?? "";
        const window = opportunity
          ? `${formatMinutes(opportunity.start_minute)}-${formatMinutes(opportunity.end_minute)}`
          : "";
        const capacity = opportunity ? ` - ${group.length} of ${opportunity.capacity} seats filled` : "";
        lines.push(`### ${day} ${window}${capacity}`.replace(/\s+/g, " ").trim(), "");
        const contact = siteContact(opportunity);
        if (opportunity?.site_address) lines.push(`Address: ${opportunity.site_address}`, "");
        if (contact) lines.push(contact, "");
        const ordered = [...group].sort(
          (left, right) => (left.start_minute ?? 0) - (right.start_minute ?? 0) || compareByStudent(left, right),
        );
        for (const entry of ordered) {
          lines.push(
            `- **${timeRange(entry)}** - ${studentLabel(entry, emailByStudent.get(entry.student_id), false)}`,
          );
        }
        lines.push("");
      }
    }
  }

  lines.push(`# ${UNPLACED_HEADING}`, "");
  if (unplaced.length === 0) {
    lines.push("All students have a placement.", "");
  } else {
    for (const entry of [...unplaced].sort(compareByStudent)) {
      const reason = UNPLACED_REASONS[entry.unplaced_reason ?? "NO_FEASIBLE_CANDIDATE"];
      lines.push(`- **${studentLabel(entry, emailByStudent.get(entry.student_id))}** - ${reason}`);
    }
    lines.push("");
  }

  const warned = placed.filter((entry) => entry.warnings.length > 0).sort(compareByStudent);
  lines.push(`## ${WARNINGS_HEADING}`, "");
  if (warned.length === 0) {
    lines.push("No placement warnings.", "");
  } else {
    for (const entry of warned) {
      lines.push(`- **${entry.student_name}** - ${entry.site_name ?? ""} ${entry.day ?? ""} ${timeRange(entry)}`);
      for (const warning of entry.warnings) lines.push(`  - ${warning.message}`);
    }
    lines.push("");
  }

  return {
    markdown: lines.join("\n").trimEnd() + "\n",
    pageBreakBeforeHeadings: [SITE_ROSTER_HEADING, UNPLACED_HEADING],
  };
}
