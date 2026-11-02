import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../lib/api";
import { SCHEDULING_MODULES } from "../moduleRegistry";
import { TutorModule } from "./TutorModule";

vi.mock("../lib/api", () => ({
  api: {
    listTutorTutors: vi.fn(),
    listTutorTags: vi.fn(),
    listTutorStudents: vi.fn(),
    getTutorReadiness: vi.fn(),
    getTutorModuleState: vi.fn(),
    getTutorSchedule: vi.fn(),
    generateTutorSchedule: vi.fn(),
    clearTutorSchedule: vi.fn(),
  },
}));
vi.mock("../lib/platform", () => ({ chooseLocalFile: vi.fn() }));

beforeEach(() => {
  vi.mocked(api.listTutorTutors).mockResolvedValue([]);
  vi.mocked(api.listTutorTags).mockResolvedValue([]);
  vi.mocked(api.listTutorStudents).mockResolvedValue([]);
  vi.mocked(api.getTutorReadiness).mockResolvedValue({
    source_data_revision: 0,
    is_ready: false,
    student_count: 0,
    tutor_count: 0,
    feasible_candidate_count: 0,
    total_candidate_count: 0,
    students_with_candidates: 0,
    tutors_with_schedulable_tutors: 0,
    warnings: [],
  });
  vi.mocked(api.getTutorModuleState).mockResolvedValue({
    schedule_revision: 0,
    source_data_revision: 0,
    is_schedule_stale: false,
  });
  vi.mocked(api.getTutorSchedule).mockResolvedValue({
    schedule_revision: 0,
    source_data_revision: 0,
    is_schedule_stale: false,
    has_schedule: false,
    students_scheduled: 0,
    students_unscheduled: 0,
    automatic_count: 0,
    manual_count: 0,
    manually_unplaced_count: 0,
    verified_count: 0,
    unverified_count: 0,
    tutor_utilization: 0,
    preferred_tag_matches: 0,
    tentative_minutes: 0,
    gap_cost_minutes: 0,
    entries: [],
    verification: [],
  });
});

describe("Tutor Scheduling module", () => {
  it("is registered as an available module independent of Clinical", () => {
    const module = SCHEDULING_MODULES.find((item) => item.key === "tutor");
    expect(module).toMatchObject({ name: "Tutor Scheduling", availability: "available" });
    expect(SCHEDULING_MODULES.filter((item) => item.key === "tutor")).toHaveLength(1);
    expect(SCHEDULING_MODULES.map((item) => item.key)).toContain("clinical");
  });

  it("shows the four tabs in order and starts on the Tutor Roster", () => {
    render(<TutorModule />);
    const labels = screen
      .getAllByRole("button")
      .map((button) => button.textContent)
      .slice(0, 4);
    expect(labels).toEqual(["Tutor Roster", "Student Roster", "Schedule", "Reports"]);
    expect(screen.getByRole("heading", { name: "Tutor Roster" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Tutor Roster" }).getAttribute("aria-current")).toBe("page");
  });

  it("renders the empty Tutor Roster with its actions", async () => {
    render(<TutorModule />);
    expect(await screen.findByText("No tutors yet.")).toBeTruthy();
    expect(screen.getByRole("button", { name: "Add Tutor" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Import Tutor Roster" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Manage Tags" })).toBeTruthy();
  });

  it("renders the Student Roster, Schedule readiness, and Reports", async () => {
    const user = userEvent.setup();
    render(<TutorModule />);

    await user.click(screen.getByRole("button", { name: "Student Roster" }));
    expect(screen.getByRole("heading", { name: "Student Roster" })).toBeTruthy();
    expect(screen.getByText("No students yet.")).toBeTruthy();
    expect(screen.getByRole("button", { name: "Add Student" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Import Student Roster" })).toBeTruthy();

    await user.click(screen.getByRole("button", { name: "Schedule" }));
    expect(await screen.findByRole("heading", { name: "Warnings" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Generate Schedule" })).toBeTruthy();

    await user.click(screen.getByRole("button", { name: "Reports" }));
    expect(screen.getByRole("heading", { name: "Tutor Reports" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Generate report" })).toBeTruthy();
  });
});
