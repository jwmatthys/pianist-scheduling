import { render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../../lib/api";
import type { ClinicalReadiness, ClinicalSchedule, ClinicalScheduleEntry, PlacementOpportunity } from "../../lib/types";
import { ClinicalSchedulePage } from "./ClinicalSchedulePage";

vi.mock("../../lib/api", () => ({
  api: {
    getClinicalReadiness: vi.fn(),
    getClinicalSchedule: vi.fn(),
    listClinicalOpportunities: vi.fn(),
    validateClinicalManualPlacement: vi.fn(),
    generateClinicalSchedule: vi.fn(),
    clearClinicalSchedule: vi.fn(),
    createClinicalManualPlacement: vi.fn(),
    moveClinicalManualPlacement: vi.fn(),
    markClinicalStudentManuallyUnplaced: vi.fn(),
    removeClinicalPlacement: vi.fn(),
    removeClinicalPlacementOverride: vi.fn(),
  },
}));

const emptySchedule: ClinicalSchedule = {
  schedule_revision: 0,
  source_data_revision: 3,
  is_schedule_stale: false,
  has_placements: false,
  student_count: 42,
  placed_count: 0,
  unplaced_count: 42,
  used_capacity: 0,
  total_capacity: 50,
  utilization: 0,
  preferred_tag_matches: 0,
  tentative_minutes: 0,
  unverified_placement_count: 0,
  unplaced: [],
  entries: [],
  verification: [],
  placements: [],
};
const generated: ClinicalSchedule = {
  ...emptySchedule,
  schedule_revision: 1,
  has_placements: true,
  placed_count: 40,
  unplaced_count: 2,
  used_capacity: 40,
  utilization: 0.8,
  preferred_tag_matches: 7,
  tentative_minutes: 120,
  unplaced: [{ student_id: "x", student_name: "Dee Demo", reason: "CAPACITY_LIMIT" }],
  entries: [{
    student_id: "x",
    student_name: "Dee Demo",
    student_number: "DEMO-1",
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
    unplaced_reason: "CAPACITY_LIMIT",
  }],
};
const opportunity: PlacementOpportunity = {
  id: "op-1",
  site_name: "Mercy Clinic",
  site_address: "",
  contact_name: "",
  contact_email: "",
  contact_phone: "",
  day_of_week: "Tuesday",
  start_minute: 8 * 60,
  end_minute: 11 * 60,
  session_length_minutes: 60,
  travel_minutes: 15,
  capacity: 1,
  tags: [{ id: "tag-1", name: "Mental Health" }],
  placement_count: 0,
};
const studentEntry: ClinicalScheduleEntry = {
  student_id: "student-1",
  student_name: "Ada Example",
  student_number: "A-001",
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
};
const manualEntry: ClinicalScheduleEntry = {
  ...studentEntry,
  kind: "manual",
  placement_state: "assigned",
  opportunity_id: opportunity.id,
  site_name: opportunity.site_name,
  opportunity_label: "Mercy Clinic, Tue 8:00 AM-11:00 AM",
  day: "Tuesday",
  start_minute: 8 * 60 + 10,
  end_minute: 9 * 60 + 10,
  tags: ["Mental Health"],
  verified: false,
  warnings: [{ code: "OUTSIDE_AVAILABILITY", message: "Outside student availability." }],
  unplaced_reason: null,
};

const base: ClinicalReadiness = {
  source_data_revision: 3,
  is_ready: true,
  student_count: 42,
  opportunity_count: 18,
  feasible_candidate_count: 300,
  total_candidate_count: 756,
  students_with_candidates: 40,
  warnings: [],
};

const studentWarning = (code: string, id: string, label: string) => ({
  code,
  severity: "warning" as const,
  category: "students" as const,
  subject: { kind: "student" as const, id, label },
});

describe("ClinicalSchedulePage readiness", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(api.getClinicalSchedule).mockResolvedValue(emptySchedule);
    vi.mocked(api.listClinicalOpportunities).mockResolvedValue([]);
  });

  it("shows no warnings and enables generation when ready", async () => {
    vi.mocked(api.getClinicalReadiness).mockResolvedValue(base);
    render(<ClinicalSchedulePage />);
    expect(await screen.findByText(/No readiness warnings/)).toBeTruthy();
    expect((screen.getByRole("button", { name: "Generate Schedule" }) as HTMLButtonElement).disabled).toBe(false);
    expect(screen.queryByText(/students loaded/)).toBeNull();
    expect(screen.queryByRole("button", { name: "Refresh" })).toBeNull();
  });

  it("groups warnings by category with counts and subjects", async () => {
    vi.mocked(api.getClinicalReadiness).mockResolvedValue({
      ...base,
      warnings: [
        studentWarning("STUDENT_WITHOUT_AVAILABILITY", "a", "Ada Example"),
        studentWarning("STUDENT_WITHOUT_AVAILABILITY", "b", "Ben Sample"),
        studentWarning("STUDENT_WITHOUT_AVAILABILITY", "c", "Cy Tester"),
        studentWarning("STUDENT_WITHOUT_CANDIDATES", "d", "Dee Demo"),
        studentWarning("STUDENT_WITHOUT_CANDIDATES", "e", "Eli Mock"),
        {
          code: "REQUIRED_TAG_WITHOUT_MATCH",
          severity: "warning",
          category: "tags",
          subject: { kind: "tag", id: "t", label: "MENTAL HEALTH" },
        },
      ],
    });
    render(<ClinicalSchedulePage />);
    expect(await screen.findByText(/3 students have no availability/)).toBeTruthy();
    expect(screen.getByText(/2 students have no feasible opportunities/)).toBeTruthy();
    expect(screen.getByText("Ada Example")).toBeTruthy();
    expect(screen.getByText(/Required tag “MENTAL HEALTH” has no matching site/)).toBeTruthy();
    for (const heading of ["Students", "Tags"]) {
      expect(screen.getByRole("heading", { name: heading })).toBeTruthy();
    }
    expect(screen.queryByRole("heading", { name: "Placement Opportunities" })).toBeNull();
  });

  it("disables generation for blocking errors and shows the warning", async () => {
    vi.mocked(api.getClinicalReadiness).mockResolvedValue({
      ...base,
      is_ready: false,
      student_count: 0,
      warnings: [{ code: "NO_STUDENTS", severity: "error", category: "data", subject: null }],
    });
    render(<ClinicalSchedulePage />);
    expect(await screen.findByText(/No students have been entered/)).toBeTruthy();
    expect((screen.getByRole("button", { name: "Generate Schedule" }) as HTMLButtonElement).disabled).toBe(true);
    expect(screen.queryByRole("status", { name: "Readiness status" })).toBeNull();
  });

  it("reports load errors", async () => {
    vi.mocked(api.getClinicalReadiness).mockRejectedValueOnce(new Error("boom"));
    render(<ClinicalSchedulePage />);
    await waitFor(() => expect(within(screen.getByRole("alert")).getByText("boom")).toBeTruthy());
  });
});

describe("ClinicalSchedulePage generation", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(api.getClinicalReadiness).mockResolvedValue(base);
    vi.mocked(api.getClinicalSchedule).mockResolvedValue(emptySchedule);
    vi.mocked(api.listClinicalOpportunities).mockResolvedValue([]);
  });

  it("generates a schedule, shows the summary, and keeps readiness visible", async () => {
    vi.mocked(api.generateClinicalSchedule).mockResolvedValue(generated);
    render(<ClinicalSchedulePage />);
    const generate = await screen.findByRole("button", { name: "Generate Schedule" });
    expect((screen.getByRole("button", { name: "Clear Schedule" }) as HTMLButtonElement).disabled).toBe(true);
    await userEvent.click(generate);
    expect(await screen.findByText(/40 students placed/)).toBeTruthy();
    expect(screen.getByText(/2 students unplaced/)).toBeTruthy();
    expect(screen.getByText(/Utilization: 80% \(40 of 50 seats\)/)).toBeTruthy();
    expect(screen.getByText(/Preferred-tag matches: 7/)).toBeTruthy();
    expect(screen.getByText(/Tentative minutes: 120/)).toBeTruthy();
    expect(screen.getByText("Dee Demo")).toBeTruthy();
    expect(screen.getByText("Capacity limit")).toBeTruthy();
    expect(screen.getByText(/No readiness warnings/)).toBeTruthy();
  });

  it("disables generation when not ready", async () => {
    vi.mocked(api.getClinicalReadiness).mockResolvedValue({ ...base, is_ready: false });
    render(<ClinicalSchedulePage />);
    expect((await screen.findByRole("button", { name: "Generate Schedule" }) as HTMLButtonElement).disabled).toBe(true);
  });

  it("shows a stale notice and clears only after confirmation", async () => {
    vi.mocked(api.getClinicalSchedule).mockResolvedValue({ ...generated, is_schedule_stale: true });
    vi.mocked(api.clearClinicalSchedule).mockResolvedValue(emptySchedule);
    render(<ClinicalSchedulePage />);
    expect(await screen.findByText(/Schedule is out of date/)).toBeTruthy();
    await userEvent.click(screen.getByRole("button", { name: "Clear Schedule" }));
    expect(api.clearClinicalSchedule).not.toHaveBeenCalled();
    const dialog = screen.getByRole("dialog", { name: "Clear Schedule?", hidden: true });
    await userEvent.click(within(dialog).getByRole("button", { name: "Clear Schedule", hidden: true }));
    await waitFor(() => expect(api.clearClinicalSchedule).toHaveBeenCalledTimes(1));
    await waitFor(() => expect(screen.queryByText(/students placed/)).toBeNull());
  });

  it("shows a structured generation error", async () => {
    vi.mocked(api.generateClinicalSchedule).mockRejectedValue(new Error("Placements cannot be generated"));
    render(<ClinicalSchedulePage />);
    await userEvent.click(await screen.findByRole("button", { name: "Generate Schedule" }));
    expect((await screen.findByRole("alert")).textContent).toContain("cannot be generated");
  });

  it("assigns an unplaced student from the drop-down and shows manual warnings", async () => {
    const initial: ClinicalSchedule = { ...emptySchedule, student_count: 1, unplaced_count: 1, entries: [studentEntry] };
    const saved: ClinicalSchedule = {
      ...initial,
      has_placements: true,
      placed_count: 1,
      unplaced_count: 0,
      entries: [manualEntry],
      placements: [{ student_id: "student-1", kind: "manual", opportunity_id: opportunity.id, start_minute: 490 }],
    };
    vi.mocked(api.getClinicalSchedule).mockResolvedValueOnce(initial);
    vi.mocked(api.listClinicalOpportunities).mockResolvedValue([opportunity]);
    vi.mocked(api.createClinicalManualPlacement).mockResolvedValue(saved);
    render(<ClinicalSchedulePage />);
    expect(screen.queryByRole("button", { name: "Assign" })).toBeNull();
    await userEvent.selectOptions(await screen.findByLabelText(/Placement Site for/), opportunity.id);
    await waitFor(() => expect(api.createClinicalManualPlacement).toHaveBeenCalledWith({
      student_id: "student-1",
      opportunity_id: opportunity.id,
      start_minute: opportunity.start_minute,
    }));
    expect((await screen.findAllByText("Outside student availability.")).length).toBeGreaterThan(0);
    expect(screen.getByText("Manual")).toBeTruthy();
  });

  it("marks manually unplaced from the drop-down and returns to automatic with the placeholder option", async () => {
    const manuallyUnplaced: ClinicalScheduleEntry = {
      ...studentEntry,
      kind: "manuallyUnplaced",
      placement_state: "manuallyUnplaced",
      unplaced_reason: "MANUALLY_UNPLACED",
    };
    const marked: ClinicalSchedule = {
      ...emptySchedule,
      has_placements: true,
      unplaced_count: 1,
      entries: [manuallyUnplaced],
      placements: [{ student_id: studentEntry.student_id, kind: "manuallyUnplaced", opportunity_id: null, start_minute: null }],
    };
    vi.mocked(api.getClinicalSchedule).mockResolvedValueOnce({ ...emptySchedule, entries: [studentEntry], student_count: 1, unplaced_count: 1 });
    vi.mocked(api.markClinicalStudentManuallyUnplaced).mockResolvedValue(marked);
    vi.mocked(api.removeClinicalPlacement).mockResolvedValue({
      ...emptySchedule,
      entries: [studentEntry],
      student_count: 1,
      unplaced_count: 1,
    });
    render(<ClinicalSchedulePage />);
    expect(screen.queryByRole("button", { name: /Mark Manually Unplaced|Remove Manual Unplaced/ })).toBeNull();
    await userEvent.selectOptions(await screen.findByLabelText(/Placement Site for/), "__manually_unplaced__");
    await waitFor(() => expect(api.markClinicalStudentManuallyUnplaced).toHaveBeenCalledWith(studentEntry.student_id));
    await waitFor(() => expect(
      (screen.getByLabelText(/Placement Site for/) as HTMLSelectElement).value,
    ).toBe("__manually_unplaced__"));
    await userEvent.selectOptions(screen.getByLabelText(/Placement Site for/), "__unassigned__");
    await waitFor(() => expect(api.removeClinicalPlacement).toHaveBeenCalledWith(studentEntry.student_id));
  });

  it("lists Unassigned first and Manually Unplaced last in every placement dropdown", async () => {
    vi.mocked(api.getClinicalSchedule).mockResolvedValue({ ...generated, entries: [manualEntry, { ...studentEntry, student_id: "other" }] });
    vi.mocked(api.listClinicalOpportunities).mockResolvedValue([opportunity]);
    render(<ClinicalSchedulePage />);
    const selects = (await screen.findAllByLabelText(/Placement Site for/)) as HTMLSelectElement[];
    expect(selects).toHaveLength(2);
    for (const select of selects) {
      const labels = Array.from(select.options).filter((option) => !option.disabled).map((option) => option.textContent);
      expect(labels[0]).toBe("Unassigned");
      expect(labels[labels.length - 1]).toBe("Manually Unplaced");
      expect(labels).toHaveLength(3);
    }
  });

  it("returns an assigned placement to Unassigned from its dropdown", async () => {
    vi.mocked(api.getClinicalSchedule).mockResolvedValue({ ...generated, entries: [manualEntry] });
    vi.mocked(api.listClinicalOpportunities).mockResolvedValue([opportunity]);
    vi.mocked(api.removeClinicalPlacement).mockResolvedValue({ ...generated, entries: [] });
    render(<ClinicalSchedulePage />);
    await userEvent.selectOptions(await screen.findByLabelText(/Placement Site for/), "__unassigned__");
    await waitFor(() => expect(api.removeClinicalPlacement).toHaveBeenCalledWith(manualEntry.student_id));
    expect(api.markClinicalStudentManuallyUnplaced).not.toHaveBeenCalled();
  });

  it("no longer shows a Placement Verification section and keeps the summary at the bottom", async () => {
    vi.mocked(api.getClinicalSchedule).mockResolvedValue({ ...generated, entries: [manualEntry] });
    vi.mocked(api.listClinicalOpportunities).mockResolvedValue([opportunity]);
    render(<ClinicalSchedulePage />);
    await screen.findByText("Schedule Summary");
    expect(screen.queryByText("Placement Verification")).toBeNull();
    const headings = screen.getAllByRole("heading", { level: 3 }).map((node) => node.textContent);
    expect(headings[headings.length - 1]).toBe("Schedule Summary");
  });

  it("edits an existing placement into a manual move", async () => {
    const initial: ClinicalSchedule = {
      ...generated,
      entries: [manualEntry],
      placements: [{
        student_id: manualEntry.student_id,
        kind: "manual",
        opportunity_id: opportunity.id,
        start_minute: manualEntry.start_minute,
      }],
    };
    const moved: ClinicalSchedule = {
      ...initial,
      schedule_revision: 2,
      entries: [{ ...manualEntry, start_minute: 8 * 60 + 30, end_minute: 9 * 60 + 30 }],
      placements: [{
        student_id: manualEntry.student_id,
        kind: "manual",
        opportunity_id: opportunity.id,
        start_minute: 8 * 60 + 30,
      }],
    };
    vi.mocked(api.getClinicalSchedule).mockResolvedValue(initial);
    vi.mocked(api.listClinicalOpportunities).mockResolvedValue([opportunity]);
    vi.mocked(api.validateClinicalManualPlacement).mockResolvedValue({ warnings: [] });
    vi.mocked(api.moveClinicalManualPlacement).mockResolvedValue(moved);
    render(<ClinicalSchedulePage />);
    const timeInput = await screen.findByLabelText(`Start time for ${manualEntry.student_name}`);
    await userEvent.clear(timeInput);
    await userEvent.type(timeInput, "8:30 AM");
    await userEvent.tab();
    await waitFor(() => expect(api.moveClinicalManualPlacement).toHaveBeenCalledWith(manualEntry.student_id, {
      student_id: manualEntry.student_id,
      opportunity_id: opportunity.id,
      start_minute: 8 * 60 + 30,
    }));
    expect(await screen.findByDisplayValue("8:30 AM")).toBeTruthy();
  });

  it("changes the site from the row drop-down and marks conflicts in amber", async () => {
    const other = { ...opportunity, id: "opp-2", site_name: "Other Clinic" };
    const initial: ClinicalSchedule = {
      ...generated,
      entries: [manualEntry],
      placements: [{
        student_id: manualEntry.student_id,
        kind: "manual",
        opportunity_id: opportunity.id,
        start_minute: manualEntry.start_minute,
      }],
    };
    const conflicted: ClinicalSchedule = {
      ...initial,
      entries: [{
        ...manualEntry,
        opportunity_id: other.id,
        site_name: other.site_name,
        warnings: [{ code: "REQUIRED_TAG_MISSING", message: "Missing required tag MENTAL HEALTH." }],
      }],
    };
    vi.mocked(api.getClinicalSchedule).mockResolvedValue(initial);
    vi.mocked(api.listClinicalOpportunities).mockResolvedValue([opportunity, other]);
    vi.mocked(api.moveClinicalManualPlacement).mockResolvedValue(conflicted);
    render(<ClinicalSchedulePage />);
    const select = await screen.findByLabelText(`Placement Site for ${manualEntry.student_name}`);
    await userEvent.selectOptions(select, other.id);
    await waitFor(() => expect(api.moveClinicalManualPlacement).toHaveBeenCalledWith(manualEntry.student_id, {
      student_id: manualEntry.student_id,
      opportunity_id: other.id,
      start_minute: manualEntry.start_minute,
    }));
    const warning = await screen.findByText("Missing required tag MENTAL HEALTH.");
    expect(warning.closest("tr")?.className).toContain("clinical-placement-row-warning");
  });
});
