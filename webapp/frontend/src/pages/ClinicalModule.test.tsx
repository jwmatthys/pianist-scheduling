import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it, vi } from "vitest";
import { ClinicalModule } from "./ClinicalModule";

vi.mock("../lib/api", () => ({
  api: {
    listClinicalOpportunities: vi.fn().mockResolvedValue([]),
    listClinicalTags: vi.fn().mockResolvedValue([]),
    getClinicalReadiness: vi.fn().mockResolvedValue({
      source_data_revision: 0,
      is_ready: false,
      student_count: 0,
      opportunity_count: 0,
      feasible_candidate_count: 0,
      total_candidate_count: 0,
      students_with_candidates: 0,
      warnings: [],
    }),
  },
}));

describe("ClinicalModule", () => {
  it("shows the four tabs in order and starts on Clinical Sites", () => {
    render(<ClinicalModule />);
    const labels = screen
      .getAllByRole("button")
      .map((button) => button.textContent)
      .slice(0, 4);
    expect(labels).toEqual(["Clinical Sites", "Student Roster", "Schedule", "Reports"]);
    expect(screen.getByRole("heading", { name: "Clinical Sites" })).toBeTruthy();
  });

  it("shows readiness on Schedule and a placeholder on Reports", async () => {
    const user = userEvent.setup();
    render(<ClinicalModule />);

    await user.click(screen.getByRole("button", { name: "Schedule" }));
    expect(screen.getByText(/schedule generation is not implemented/i)).toBeTruthy();

    await user.click(screen.getByRole("button", { name: "Reports" }));
    expect(screen.getByText(/reports and exports are not implemented/i)).toBeTruthy();
  });
});
