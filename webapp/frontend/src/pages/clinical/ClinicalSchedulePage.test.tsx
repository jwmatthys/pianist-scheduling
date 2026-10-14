import { render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../../lib/api";
import type { ClinicalReadiness } from "../../lib/types";
import { ClinicalSchedulePage } from "./ClinicalSchedulePage";

vi.mock("../../lib/api", () => ({ api: { getClinicalReadiness: vi.fn() } }));

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
  beforeEach(() => vi.clearAllMocks());

  it("shows the readiness summary and candidate counts without any placements", async () => {
    vi.mocked(api.getClinicalReadiness).mockResolvedValue(base);
    render(<ClinicalSchedulePage />);
    expect(await screen.findByText("Ready for scheduling")).toBeTruthy();
    expect(screen.getByText(/42 students loaded/)).toBeTruthy();
    expect(screen.getByText(/18 placement opportunities loaded/)).toBeTruthy();
    expect(screen.getByText(/300 feasible student\/opportunity pairings of 756 evaluated/)).toBeTruthy();
    expect(screen.getByText(/40 of 42 students have at least one feasible opportunity/)).toBeTruthy();
    expect(screen.getByText(/No readiness warnings/)).toBeTruthy();
    expect(screen.getByText(/schedule generation is not implemented/i)).toBeTruthy();
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
