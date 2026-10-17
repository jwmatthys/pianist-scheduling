import { formatMinutes } from "./time";
import type { ReportTable } from "./reportTable";
import type { ClinicalSchedule, ClinicalScheduleEntry, ClinicalStudent, PlacementOpportunity } from "./types";

const UNPLACED_REASONS: Record<string, string> = {
  NO_FEASIBLE_CANDIDATE: "No feasible opportunity",
  CAPACITY_LIMIT: "Capacity limit",
  MANUALLY_UNPLACED: "Manually unplaced",
};

const PLACEMENT_TYPES: Record<string, string> = {
  automatic: "Automatic",
  manual: "Manual",
  manuallyUnplaced: "Manually Unplaced",
};

function status(entry: ClinicalScheduleEntry) {
  if (entry.warnings.length > 0) {
    return entry.warnings.every((warning) => warning.code === "TENTATIVE_AVAILABILITY") ? "Placement warning" : "Placement conflict";
  }
  return entry.verified ? "Verified" : "Needs review";
}

/**
 * Clinical-owned: the Schedule tab's Placements and Unplaced Students tables as one table,
 * one row per student, with the tab's columns.
 */
export function buildClinicalReportTable(
  schedule: ClinicalSchedule,
  opportunities: PlacementOpportunity[],
  students: ClinicalStudent[],
): ReportTable {
  const emailByStudent = new Map(students.map((student) => [student.id, student.email]));
  const opportunitiesById = new Map(opportunities.map((item) => [item.id, item]));
  return {
    name: "Clinical Placements",
    columns: [
      "Student", "Student ID", "Student Email", "Placement Site", "Site Address", "Contact Name", "Contact Email",
      "Contact Phone", "Opportunity Label", "Tags", "Day", "Start Time", "End Time", "Session Length (minutes)",
      "Travel (minutes)", "Placement Type", "Preferred Tag Matches", "Tentative Minutes", "Status", "Warnings",
      "Unplaced Reason",
    ],
    rows: schedule.entries.map((entry) => {
      const placed = entry.kind === "automatic" || entry.kind === "manual";
      const opportunity = entry.opportunity_id ? opportunitiesById.get(entry.opportunity_id) : undefined;
      const end = entry.start_minute !== null && opportunity ? entry.start_minute + opportunity.session_length_minutes : null;
      return [
        entry.student_name, entry.student_number, emailByStudent.get(entry.student_id) ?? "",
        entry.site_name ?? "", opportunity?.site_address ?? "", opportunity?.contact_name ?? "",
        opportunity?.contact_email ?? "", opportunity?.contact_phone ?? "", entry.opportunity_label ?? "",
        entry.tags.join(", "),
        opportunity?.day_of_week ?? entry.day ?? "",
        entry.start_minute === null ? "" : formatMinutes(entry.start_minute),
        end === null ? "" : `${formatMinutes(end % 1440)}${end >= 1440 ? ` (+${Math.floor(end / 1440)} day)` : ""}`,
        opportunity?.session_length_minutes ?? null, opportunity?.travel_minutes ?? null,
        PLACEMENT_TYPES[entry.kind ?? ""] ?? "Unplaced",
        placed ? entry.preferred_tag_match_count : null, entry.tentative_minutes,
        placed ? status(entry) : "",
        entry.warnings.map((warning) => warning.message).join("; "),
        placed ? "" : UNPLACED_REASONS[entry.unplaced_reason ?? "NO_FEASIBLE_CANDIDATE"],
      ];
    }),
  };
}
