import { render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../../lib/api";
import type { ClinicalReadiness, ClinicalSchedule } from "../../lib/types";
import { ClinicalSchedulePage } from "./ClinicalSchedulePage";

vi.mock("../../lib/api", () => ({
  api: {
    getClinicalReadiness: vi.fn(),
    getClinicalSchedule: vi.fn(),
    generateClinicalSchedule: vi.fn(),
    clearClinicalSchedule: vi.fn(),
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
  });

  it("shows the readiness summary and candidate counts without any placements", async () => {
    vi.mocked(api.getClinicalReadiness).mockResolvedValue(base);
    render(<ClinicalSchedulePage />);
    expect(await screen.findByText("Ready for scheduling")).toBeTruthy();
    expect(screen.getByText(/42 students loaded/)).toBeTruthy();
    expect(screen.getByText(/18 placement opportunities loaded/)).toBeTruthy();
    expect(screen.getByText(/300 feasible student\/opportunity pairings of 756 evaluated/)).toBeTruthy();
    expect(screen.getByText(/40 of 42 students have at least one feasible opportunity/)).toBeTruthy();
    expect(screen.getByText(/No readiness warnings/)).toBeTruthy();
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

  it("shows a not-ready status for blocking errors", async () => {
    vi.mocked(api.getClinicalReadiness).mockResolvedValue({
      ...base,
      is_ready: false,
      student_count: 0,
      warnings: [{ code: "NO_STUDENTS", severity: "error", category: "data", subject: null }],
    });
    render(<ClinicalSchedulePage />);
    const status = await screen.findByRole("status", { name: "Readiness status" });
    expect(status.textContent).toBe("Not ready for scheduling");
    expect(screen.getByText(/No students have been entered/)).toBeTruthy();
  });

  it("refreshes readiness on demand and reports errors", async () => {
    vi.mocked(api.getClinicalReadiness).mockResolvedValueOnce(base).mockRejectedValueOnce(new Error("boom"));
    const user = userEvent.setup();
    render(<ClinicalSchedulePage />);
    await screen.findByText("Ready for scheduling");
    await user.click(screen.getByRole("button", { name: "Refresh" }));
    await waitFor(() => expect(within(screen.getByRole("alert")).getByText("boom")).toBeTruthy());
    expect(api.getClinicalReadiness).toHaveBeenCalledTimes(2);
  });
});

describe("ClinicalSchedulePage generation", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(api.getClinicalReadiness).mockResolvedValue(base);
    vi.mocked(api.getClinicalSchedule).mockResolvedValue(emptySchedule);
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
    expect(screen.getByText(/Dee Demo: Capacity limit/)).toBeTruthy();
    expect(screen.getByText("Ready for scheduling")).toBeTruthy();
  });

  it("disables generation when not ready", async () => {
    vi.mocked(api.getClinicalReadiness).mockResolvedValue({ ...base, is_ready: false });
    render(<ClinicalSchedulePage />);
    await screen.findByText("Not ready for scheduling");
    expect((screen.getByRole("button", { name: "Generate Schedule" }) as HTMLButtonElement).disabled).toBe(true);
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
});
