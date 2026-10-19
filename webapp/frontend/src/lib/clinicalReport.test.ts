import { describe, expect, it } from "vitest";
import { buildClinicalReport } from "./clinicalReport";
import type { ClinicalSchedule, ClinicalScheduleEntry, ClinicalStudent, PlacementOpportunity } from "./types";

function opportunity(id: string, site: string, day: string, start: number, capacity = 2): PlacementOpportunity {
  return {
    id,
    site_name: site,
    site_address: `1 ${site} Way`,
    contact_name: "Casey Contact",
    contact_email: "contact@example.test",
    contact_phone: "555-0100",
    day_of_week: day,
    start_minute: start,
    end_minute: start + 180,
    session_length_minutes: 60,
    travel_minutes: 15,
    capacity,
    tags: [],
    placement_count: 0,
  };
}

function entry(id: string, name: string, overrides: Partial<ClinicalScheduleEntry> = {}): ClinicalScheduleEntry {
  return {
    student_id: id,
    student_name: name,
    student_number: "",
    kind: null,
    placement_state: "unassigned",
    opportunity_id: null,
    site_name: null,
    opportunity_label: null,
    day: null,
    start_minute: null,
    end_minute: null,
    tags: [],
    preferred_tag_match_count: 0,
    tentative_minutes: null,
    verified: null,
    warnings: [],
    unplaced_reason: "NO_FEASIBLE_CANDIDATE",
    ...overrides,
  };
}

function placed(id: string, name: string, opp: PlacementOpportunity, start: number, kind: "automatic" | "manual" = "automatic") {
  return entry(id, name, {
    kind,
    opportunity_id: opp.id,
    site_name: opp.site_name,
    day: opp.day_of_week,
    start_minute: start,
    end_minute: start + 60,
    unplaced_reason: null,
    verified: kind === "automatic",
  });
}

const students = ["s1", "s2", "s3", "s4"].map((id) => ({ id, email: `${id}@example.test` })) as ClinicalStudent[];
const mercy = opportunity("o-mercy-tue", "Mercy Clinic", "Tuesday", 9 * 60);
const mercyMon = opportunity("o-mercy-mon", "Mercy Clinic", "Monday", 13 * 60);
const adams = opportunity("o-adams", "Adams Center", "Friday", 8 * 60);

function schedule(entries: ClinicalScheduleEntry[], overrides: Partial<ClinicalSchedule> = {}): ClinicalSchedule {
  return { is_schedule_stale: false, entries, ...overrides } as ClinicalSchedule;
}

describe("buildClinicalReport", () => {
  const entries = [
    placed("s1", "Zed Student", mercy, 9 * 60 + 60),
    { ...placed("s2", "Amy Student", mercy, 9 * 60), student_number: "A-77" },
    placed("s3", "Bo Student", mercyMon, 13 * 60, "manual"),
    entry("s4", "Dee Student", { unplaced_reason: "CAPACITY_LIMIT" }),
  ];
  entries[2].warnings = [{ code: "REQUIRED_TAG_MISSING", message: "Missing a Required tag." }];
  const report = buildClinicalReport(schedule(entries), [mercy, mercyMon, adams], students);

  it("has student, site, unplaced and warning sections with contact details", () => {
    expect(report.markdown).toContain("## Student Placement Roster");
    expect(report.markdown).toContain("# Site Placement Roster");
    expect(report.markdown).toContain("# Unplaced Students");
    expect(report.markdown).toContain("## Placement Warnings");
    expect(report.markdown).toContain("s2@example.test");
    expect(report.markdown).toContain("Contact: Casey Contact, contact@example.test, 555-0100");
    expect(report.markdown).toContain("Dee Student - s4@example.test** - Capacity limit");
    expect(report.markdown).toContain("Missing a Required tag.");
    expect(report.markdown).toContain("3 students placed, 1 unplaced.");
  });

  it("lists student, ID, email and placement time, then site, address and contact details", () => {
    const lines = report.markdown.split("\n");
    const index = lines.findIndex((line) => line.startsWith("- **Amy Student"));
    expect(lines[index]).toBe("- **Amy Student - ID A-77 - s2@example.test** - Tuesday 9:00 AM-10:00 AM");
    expect(lines[index + 1]).toBe(
      "  - Mercy Clinic, 1 Mercy Clinic Way, Casey Contact, contact@example.test, 555-0100",
    );
  });

  it("orders students by name and sites by site, weekday, start, then session start", () => {
    const md = report.markdown;
    const studentSection = md.slice(md.indexOf("## Student Placement Roster"), md.indexOf("# Site Placement Roster"));
    expect(studentSection.indexOf("Amy Student")).toBeLessThan(studentSection.indexOf("Bo Student"));
    expect(studentSection.indexOf("Bo Student")).toBeLessThan(studentSection.indexOf("Zed Student"));
    const siteSection = md.slice(md.indexOf("# Site Placement Roster"));
    expect(siteSection.indexOf("## Mercy Clinic")).toBeGreaterThan(-1);
    expect(siteSection.indexOf("Monday")).toBeLessThan(siteSection.indexOf("Tuesday"));
    expect(siteSection.indexOf("Amy Student")).toBeLessThan(siteSection.indexOf("Zed Student"));
    expect(siteSection).toContain("2 of 2 seats filled");
    expect(siteSection).toContain("Amy Student - s2@example.test");
    expect(siteSection).not.toContain("A-77");
  });

  it("is deterministic and requests page breaks before site and unplaced sections", () => {
    expect(buildClinicalReport(schedule(entries), [mercy, mercyMon, adams], students).markdown).toBe(report.markdown);
    expect(report.pageBreakBeforeHeadings).toEqual(["Site Placement Roster", "Unplaced Students"]);
  });

  it("handles an empty schedule and flags a stale schedule", () => {
    const empty = buildClinicalReport(schedule([entry("s4", "Dee Student")], { is_schedule_stale: true }), [], students);
    expect(empty.markdown).toContain("No students have been placed.");
    expect(empty.markdown).toContain("No placement warnings.");
    expect(empty.markdown).toContain("Regenerate before relying on it");
  });
});
