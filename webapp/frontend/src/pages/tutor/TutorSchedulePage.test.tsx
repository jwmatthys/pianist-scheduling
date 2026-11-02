import { render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../../lib/api";
import type { TutorReadiness, TutorModuleState, TutorSchedule } from "../../lib/types";
import { TutorSchedulePage } from "./TutorSchedulePage";

vi.mock("../../lib/api", () => ({
  api: {
    getTutorReadiness: vi.fn(),
    getTutorModuleState: vi.fn(),
    getTutorSchedule: vi.fn(),
    generateTutorSchedule: vi.fn(),
    clearTutorSchedule: vi.fn(),
    listTutorTutors: vi.fn(),
    createTutorManualPlacement: vi.fn(),
    moveTutorManualPlacement: vi.fn(),
    manuallyUnplaceTutorStudent: vi.fn(),
    removeTutorPlacementOverride: vi.fn(),
    removeTutorPlacement: vi.fn(),
  },
}));

const readiness: TutorReadiness = {
  source_data_revision: 4,
  is_ready: false,
  student_count: 3,
  tutor_count: 2,
  feasible_candidate_count: 5,
  total_candidate_count: 7,
  students_with_candidates: 2,
  tutors_with_schedulable_tutors: 2,
  warnings: [
    {
      code: "STUDENT_WITHOUT_AVAILABILITY",
      severity: "warning",
      category: "students",
      subjects: [
        { kind: "student", id: "s-1", label: "Ada Example" },
        { kind: "student", id: "s-2", label: "Bea Example" },
      ],
    },
    {
      code: "REQUIRED_TAG_WITHOUT_MATCH",
      severity: "warning",
      category: "tags",
      subjects: [{ kind: "tag", id: "tag-1", label: "VIOLIN" }],
    },
  ],
};

const state: TutorModuleState = {
  schedule_revision: 0,
  source_data_revision: 5,
  is_schedule_stale: false,
};

const generatedSchedule: TutorSchedule = {
  schedule_revision: 3,
  source_data_revision: 4,
  is_schedule_stale: false,
  has_schedule: true,
  students_scheduled: 120,
  students_unscheduled: 5,
  tutor_utilization: 0.8,
  preferred_tag_matches: 72,
  tentative_minutes: 45,
  gap_cost_minutes: 90,
  automatic_count: 1,
  manual_count: 0,
  manually_unplaced_count: 0,
  verified_count: 1,
  unverified_count: 0,
  entries: [
    {
      student_id: "student-1",
      student_name: "Ari Example",
      placement_type: "automatic",
      placement_state: "assigned",
      tutor_id: "tutor-1",
      tutor_name: "Mina Teacher",
      day: "Monday",
      start_minute: 540,
      end_minute: 600,
      preferred_tag_match_count: 1,
      tentative_minutes: 0,
      verified: true,
      warnings: [],
      unplaced_reason: null,
    },
  ],
  verification: [],
};

describe("TutorSchedulePage readiness", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(api.getTutorReadiness).mockResolvedValue(readiness);
    vi.mocked(api.getTutorModuleState).mockResolvedValue(state);
    vi.mocked(api.listTutorTutors).mockResolvedValue([
      {
        id: "tutor-1",
        name: "Mina Teacher",
        email: "",
        room: "",
        lesson_length_minutes: 60,
        capacity: 1,
        tags: [],
        availability_window_count: 1,
        created_at: "",
        updated_at: "",
      },
    ]);
    vi.mocked(api.getTutorSchedule).mockResolvedValue({
      ...generatedSchedule,
      has_schedule: false,
      students_scheduled: 0,
      students_unscheduled: 0,
      entries: [],
    });
    vi.mocked(api.generateTutorSchedule).mockResolvedValue(generatedSchedule);
    vi.mocked(api.clearTutorSchedule).mockResolvedValue({
      ...generatedSchedule,
      has_schedule: false,
      students_scheduled: 0,
      students_unscheduled: 0,
      entries: [],
    });
    vi.mocked(api.manuallyUnplaceTutorStudent).mockResolvedValue(generatedSchedule);
    vi.mocked(api.createTutorManualPlacement).mockResolvedValue(generatedSchedule);
    vi.mocked(api.moveTutorManualPlacement).mockResolvedValue(generatedSchedule);
    vi.mocked(api.removeTutorPlacementOverride).mockResolvedValue(generatedSchedule);
    vi.mocked(api.removeTutorPlacement).mockResolvedValue(generatedSchedule);
  });

  it("shows roster counts, grouped warnings, tag details, and stale revision state", async () => {
    render(<TutorSchedulePage />);

    await screen.findByText("Warnings");
    expect(screen.queryByText("Readiness Summary")).toBeNull();
    expect(screen.getByText("2 students have no availability")).toBeTruthy();
    expect(screen.getByText("Ada Example")).toBeTruthy();
    expect(screen.getByText("Bea Example")).toBeTruthy();
    expect(screen.getByText('Required tag "VIOLIN" has no matching tutor')).toBeTruthy();
    expect(screen.queryByText("No placements assigned.")).toBeNull();
    expect(screen.getByRole("button", { name: "Generate Schedule" }).hasAttribute("disabled")).toBe(true);
  });

  it("treats an automatically unplaced student with no time as unplaced", async () => {
    vi.mocked(api.getTutorSchedule).mockResolvedValue({
      ...generatedSchedule,
      students_scheduled: 1,
      students_unscheduled: 1,
      entries: [
        ...generatedSchedule.entries,
        {
          student_id: "student-2",
          student_name: "Bea Unplaced",
          placement_type: "automatic",
          placement_state: "assigned",
          tutor_id: null,
          tutor_name: null,
          day: null,
          start_minute: null,
          end_minute: null,
          preferred_tag_match_count: 0,
          tentative_minutes: 0,
          verified: false,
          warnings: [],
          unplaced_reason: "CAPACITY_LIMIT",
        },
      ],
    });
    render(<TutorSchedulePage />);
    const row = (await screen.findByText("Bea Unplaced")).closest("tr")!;
    const unplacedHeading = screen.getByRole("heading", { name: /Unplaced/ });
    const placedRow = screen.getByText("Ari Example").closest("tr")!;
    expect(row.closest("table")).not.toBe(placedRow.closest("table"));
    expect(unplacedHeading).toBeTruthy();
  });

  it("has no Refresh Readiness button and loads readiness on mount", async () => {
    render(<TutorSchedulePage />);
    await screen.findByText("Warnings");
    expect(screen.queryByRole("button", { name: "Refresh Readiness" })).toBeNull();
    expect(api.getTutorReadiness).toHaveBeenCalledTimes(1);
    expect(api.getTutorModuleState).toHaveBeenCalledTimes(1);
    expect(api.getTutorSchedule).toHaveBeenCalledTimes(1);
  });

  it("shows readiness API errors explicitly", async () => {
    vi.mocked(api.getTutorReadiness).mockRejectedValueOnce(new Error("Readiness unavailable"));
    render(<TutorSchedulePage />);
    expect((await screen.findByRole("alert")).textContent).toContain("Readiness unavailable");
  });

  it("generates and clears schedules while keeping readiness visible", async () => {
    const user = userEvent.setup();
    vi.mocked(api.getTutorReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.getTutorSchedule).mockResolvedValue(generatedSchedule);
    render(<TutorSchedulePage />);

    await screen.findByText("Schedule Summary");
    expect(screen.getByText("120")).toBeTruthy();
    expect(screen.getByText("80.0%")).toBeTruthy();
    expect(screen.getByText("72")).toBeTruthy();
    expect(screen.getByText("45")).toBeTruthy();
    expect(screen.getByText("90 minutes")).toBeTruthy();
    expect(screen.getByText("Warnings")).toBeTruthy();
    expect(screen.getByText("Ari Example")).toBeTruthy();
    expect(screen.getByText("Automatic")).toBeTruthy();

    await user.click(screen.getByRole("button", { name: "Generate Schedule" }));
    expect(api.generateTutorSchedule).toHaveBeenCalledOnce();
    await user.click(await screen.findByRole("button", { name: "Clear Schedule" }));
    const clearDialog = screen.getByRole("dialog");
    await user.click(within(clearDialog).getByRole("button", { name: "Clear Schedule" }));
    expect(api.clearTutorSchedule).toHaveBeenCalledOnce();
  });

  it("shows lifecycle terminology and supports manually unplacing a student", async () => {
    vi.mocked(api.getTutorReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.getTutorSchedule).mockResolvedValue(generatedSchedule);
    vi.mocked(api.manuallyUnplaceTutorStudent).mockResolvedValue({
      ...generatedSchedule,
      manual_count: 0,
      manually_unplaced_count: 1,
    });
    const user = userEvent.setup();
    render(<TutorSchedulePage />);

    await screen.findByText("Placements");
    expect(screen.getByRole("list", { name: "Schedule lifecycle status" }).textContent)
      .toContain("Verified: 1");
    expect(screen.queryByRole("button", { name: "Manually Unplace" })).toBeNull();
    expect(screen.queryByRole("button", { name: /Assign Manually|Reassign/ })).toBeNull();
    const select = screen.getByLabelText("Tutor for Ari Example") as HTMLSelectElement;
    expect(select.options[select.options.length - 1].textContent).toBe("Manually Unplaced");
    await user.selectOptions(select, "Manually Unplaced");

    expect(api.manuallyUnplaceTutorStudent).toHaveBeenCalledWith("student-1");
  });

  it("shows schedule API errors explicitly", async () => {
    vi.mocked(api.getTutorReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.generateTutorSchedule).mockRejectedValueOnce(new Error("Schedule generation failed"));
    const user = userEvent.setup();
    render(<TutorSchedulePage />);

    await screen.findByRole("button", { name: "Generate Schedule" });
    await user.click(screen.getByRole("button", { name: "Generate Schedule" }));

    expect((await screen.findByRole("alert")).textContent).toContain("Schedule generation failed");
  });

  it("treats a changed day or start time as a manual adjustment", async () => {
    vi.mocked(api.getTutorReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.getTutorSchedule).mockResolvedValue(generatedSchedule);
    const user = userEvent.setup();
    render(<TutorSchedulePage />);

    await screen.findByText("Placements");
    await user.selectOptions(screen.getByLabelText("Day for Ari Example"), "Tuesday");
    expect(api.moveTutorManualPlacement).toHaveBeenCalledWith("student-1", {
      student_id: "student-1",
      tutor_id: "tutor-1",
      day: "Tuesday",
      start_minute: 540,
    });

    const time = screen.getByLabelText("Start time for Ari Example");
    await user.clear(time);
    await user.type(time, "10:00 AM");
    await user.tab();
    expect(api.moveTutorManualPlacement).toHaveBeenLastCalledWith("student-1", {
      student_id: "student-1",
      tutor_id: "tutor-1",
      day: "Tuesday",
      start_minute: 600,
    });
  });

  it("removes an override from the unplaced list", async () => {
    vi.mocked(api.getTutorReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.getTutorSchedule).mockResolvedValue({
      ...generatedSchedule,
      entries: [{
        ...generatedSchedule.entries[0],
        placement_type: "manuallyUnplaced",
        placement_state: "manuallyUnplaced",
        tutor_id: null,
        tutor_name: null,
        day: null,
        start_minute: null,
        end_minute: null,
        unplaced_reason: "MANUALLY_UNPLACED",
      }],
    });
    const user = userEvent.setup();
    render(<TutorSchedulePage />);

    await screen.findByText("Unplaced Students");
    await user.selectOptions(screen.getByLabelText("Tutor for Ari Example"), "Unassigned");
    expect(api.removeTutorPlacement).toHaveBeenCalledWith("student-1");
  });

  it("lists Unassigned first and Manually Unplaced last in every dropdown", async () => {
    vi.mocked(api.getTutorReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.getTutorSchedule).mockResolvedValue({
      ...generatedSchedule,
      entries: [
        generatedSchedule.entries[0],
        {
          ...generatedSchedule.entries[0],
          student_id: "student-2",
          student_name: "Bea Example",
          placement_type: null,
          placement_state: "unassigned",
          tutor_id: null,
          tutor_name: null,
          day: null,
          start_minute: null,
          end_minute: null,
          unplaced_reason: "NO_FEASIBLE_CANDIDATE",
        },
      ],
    });
    render(<TutorSchedulePage />);
    await screen.findByText("Unplaced Students");
    for (const name of ["Ari Example", "Bea Example"]) {
      const select = screen.getByLabelText(`Tutor for ${name}`) as HTMLSelectElement;
      const labels = Array.from(select.options).filter((option) => !option.disabled).map((option) => option.textContent);
      expect(labels).toEqual(["Unassigned", "Mina Teacher", "Manually Unplaced"]);
    }
    expect((screen.getByLabelText("Tutor for Bea Example") as HTMLSelectElement).value).toBe("__unassigned__");
  });

  it("returns an assigned tutor to Unassigned from its dropdown", async () => {
    vi.mocked(api.getTutorReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.getTutorSchedule).mockResolvedValue(generatedSchedule);
    const user = userEvent.setup();
    render(<TutorSchedulePage />);
    await screen.findByText("Placements");
    await user.selectOptions(screen.getByLabelText("Tutor for Ari Example"), "Unassigned");
    expect(api.removeTutorPlacement).toHaveBeenCalledWith("student-1");
    expect(api.manuallyUnplaceTutorStudent).not.toHaveBeenCalled();
  });
});
