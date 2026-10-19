import { render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../../lib/api";
import type { LessonReadiness, LessonModuleState, LessonSchedule } from "../../lib/types";
import { LessonSchedulePage } from "./LessonSchedulePage";

vi.mock("../../lib/api", () => ({
  api: {
    getLessonReadiness: vi.fn(),
    getLessonModuleState: vi.fn(),
    getLessonSchedule: vi.fn(),
    generateLessonSchedule: vi.fn(),
    clearLessonSchedule: vi.fn(),
    listLessonInstructors: vi.fn(),
    createLessonManualPlacement: vi.fn(),
    moveLessonManualPlacement: vi.fn(),
    manuallyUnplaceLessonStudent: vi.fn(),
    removeLessonPlacementOverride: vi.fn(),
    removeLessonPlacement: vi.fn(),
  },
}));

const readiness: LessonReadiness = {
  source_data_revision: 4,
  is_ready: false,
  student_count: 3,
  instructor_count: 2,
  feasible_candidate_count: 5,
  total_candidate_count: 7,
  students_with_candidates: 2,
  instructors_with_schedulable_lessons: 2,
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

const state: LessonModuleState = {
  schedule_revision: 0,
  source_data_revision: 5,
  is_schedule_stale: false,
};

const generatedSchedule: LessonSchedule = {
  schedule_revision: 3,
  source_data_revision: 4,
  is_schedule_stale: false,
  has_schedule: true,
  students_scheduled: 120,
  students_unscheduled: 5,
  instructor_utilization: 0.8,
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
      instructor_id: "instructor-1",
      instructor_name: "Mina Teacher",
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

describe("LessonSchedulePage readiness", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(api.getLessonReadiness).mockResolvedValue(readiness);
    vi.mocked(api.getLessonModuleState).mockResolvedValue(state);
    vi.mocked(api.listLessonInstructors).mockResolvedValue([
      {
        id: "instructor-1",
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
    vi.mocked(api.getLessonSchedule).mockResolvedValue({
      ...generatedSchedule,
      has_schedule: false,
      students_scheduled: 0,
      students_unscheduled: 0,
      entries: [],
    });
    vi.mocked(api.generateLessonSchedule).mockResolvedValue(generatedSchedule);
    vi.mocked(api.clearLessonSchedule).mockResolvedValue({
      ...generatedSchedule,
      has_schedule: false,
      students_scheduled: 0,
      students_unscheduled: 0,
      entries: [],
    });
    vi.mocked(api.manuallyUnplaceLessonStudent).mockResolvedValue(generatedSchedule);
    vi.mocked(api.createLessonManualPlacement).mockResolvedValue(generatedSchedule);
    vi.mocked(api.moveLessonManualPlacement).mockResolvedValue(generatedSchedule);
    vi.mocked(api.removeLessonPlacementOverride).mockResolvedValue(generatedSchedule);
    vi.mocked(api.removeLessonPlacement).mockResolvedValue(generatedSchedule);
  });

  it("shows roster counts, grouped warnings, tag details, and stale revision state", async () => {
    render(<LessonSchedulePage />);

    await screen.findByText("Warnings");
    expect(screen.queryByText("Readiness Summary")).toBeNull();
    expect(screen.getByText("2 students have no availability")).toBeTruthy();
    expect(screen.getByText("Ada Example")).toBeTruthy();
    expect(screen.getByText("Bea Example")).toBeTruthy();
    expect(screen.getByText('Required tag "VIOLIN" has no matching instructor')).toBeTruthy();
    expect(screen.queryByText("No placements assigned.")).toBeNull();
    expect(screen.getByRole("button", { name: "Generate Schedule" }).hasAttribute("disabled")).toBe(true);
  });

  it("has no Refresh Readiness button and loads readiness on mount", async () => {
    render(<LessonSchedulePage />);
    await screen.findByText("Warnings");
    expect(screen.queryByRole("button", { name: "Refresh Readiness" })).toBeNull();
    expect(api.getLessonReadiness).toHaveBeenCalledTimes(1);
    expect(api.getLessonModuleState).toHaveBeenCalledTimes(1);
    expect(api.getLessonSchedule).toHaveBeenCalledTimes(1);
  });

  it("shows readiness API errors explicitly", async () => {
    vi.mocked(api.getLessonReadiness).mockRejectedValueOnce(new Error("Readiness unavailable"));
    render(<LessonSchedulePage />);
    expect((await screen.findByRole("alert")).textContent).toContain("Readiness unavailable");
  });

  it("generates and clears schedules while keeping readiness visible", async () => {
    const user = userEvent.setup();
    vi.mocked(api.getLessonReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.getLessonSchedule).mockResolvedValue(generatedSchedule);
    render(<LessonSchedulePage />);

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
    expect(api.generateLessonSchedule).toHaveBeenCalledOnce();
    await user.click(await screen.findByRole("button", { name: "Clear Schedule" }));
    const clearDialog = screen.getByRole("dialog");
    await user.click(within(clearDialog).getByRole("button", { name: "Clear Schedule" }));
    expect(api.clearLessonSchedule).toHaveBeenCalledOnce();
  });

  it("shows lifecycle terminology and supports manually unplacing a student", async () => {
    vi.mocked(api.getLessonReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.getLessonSchedule).mockResolvedValue(generatedSchedule);
    vi.mocked(api.manuallyUnplaceLessonStudent).mockResolvedValue({
      ...generatedSchedule,
      manual_count: 0,
      manually_unplaced_count: 1,
    });
    const user = userEvent.setup();
    render(<LessonSchedulePage />);

    await screen.findByText("Placements");
    expect(screen.getByRole("list", { name: "Schedule lifecycle status" }).textContent)
      .toContain("Verified: 1");
    expect(screen.queryByRole("button", { name: "Manually Unplace" })).toBeNull();
    expect(screen.queryByRole("button", { name: /Assign Manually|Reassign/ })).toBeNull();
    const select = screen.getByLabelText("Instructor for Ari Example") as HTMLSelectElement;
    expect(select.options[select.options.length - 1].textContent).toBe("Manually Unplaced");
    await user.selectOptions(select, "Manually Unplaced");

    expect(api.manuallyUnplaceLessonStudent).toHaveBeenCalledWith("student-1");
  });

  it("shows schedule API errors explicitly", async () => {
    vi.mocked(api.getLessonReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.generateLessonSchedule).mockRejectedValueOnce(new Error("Schedule generation failed"));
    const user = userEvent.setup();
    render(<LessonSchedulePage />);

    await screen.findByRole("button", { name: "Generate Schedule" });
    await user.click(screen.getByRole("button", { name: "Generate Schedule" }));

    expect((await screen.findByRole("alert")).textContent).toContain("Schedule generation failed");
  });

  it("treats a changed day or start time as a manual adjustment", async () => {
    vi.mocked(api.getLessonReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.getLessonSchedule).mockResolvedValue(generatedSchedule);
    const user = userEvent.setup();
    render(<LessonSchedulePage />);

    await screen.findByText("Placements");
    await user.selectOptions(screen.getByLabelText("Day for Ari Example"), "Tuesday");
    expect(api.moveLessonManualPlacement).toHaveBeenCalledWith("student-1", {
      student_id: "student-1",
      instructor_id: "instructor-1",
      day: "Tuesday",
      start_minute: 540,
    });

    const time = screen.getByLabelText("Start time for Ari Example");
    await user.clear(time);
    await user.type(time, "10:00 AM");
    await user.tab();
    expect(api.moveLessonManualPlacement).toHaveBeenLastCalledWith("student-1", {
      student_id: "student-1",
      instructor_id: "instructor-1",
      day: "Tuesday",
      start_minute: 600,
    });
  });

  it("removes an override from the unplaced list", async () => {
    vi.mocked(api.getLessonReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.getLessonSchedule).mockResolvedValue({
      ...generatedSchedule,
      entries: [{
        ...generatedSchedule.entries[0],
        placement_type: "manuallyUnplaced",
        placement_state: "manuallyUnplaced",
        instructor_id: null,
        instructor_name: null,
        day: null,
        start_minute: null,
        end_minute: null,
        unplaced_reason: "MANUALLY_UNPLACED",
      }],
    });
    const user = userEvent.setup();
    render(<LessonSchedulePage />);

    await screen.findByText("Unplaced Students");
    await user.selectOptions(screen.getByLabelText("Instructor for Ari Example"), "Unassigned");
    expect(api.removeLessonPlacement).toHaveBeenCalledWith("student-1");
  });

  it("lists Unassigned first and Manually Unplaced last in every dropdown", async () => {
    vi.mocked(api.getLessonReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.getLessonSchedule).mockResolvedValue({
      ...generatedSchedule,
      entries: [
        generatedSchedule.entries[0],
        {
          ...generatedSchedule.entries[0],
          student_id: "student-2",
          student_name: "Bea Example",
          placement_type: null,
          placement_state: "unassigned",
          instructor_id: null,
          instructor_name: null,
          day: null,
          start_minute: null,
          end_minute: null,
          unplaced_reason: "NO_FEASIBLE_CANDIDATE",
        },
      ],
    });
    render(<LessonSchedulePage />);
    await screen.findByText("Unplaced Students");
    for (const name of ["Ari Example", "Bea Example"]) {
      const select = screen.getByLabelText(`Instructor for ${name}`) as HTMLSelectElement;
      const labels = Array.from(select.options).filter((option) => !option.disabled).map((option) => option.textContent);
      expect(labels).toEqual(["Unassigned", "Mina Teacher", "Manually Unplaced"]);
    }
    expect((screen.getByLabelText("Instructor for Bea Example") as HTMLSelectElement).value).toBe("__unassigned__");
  });

  it("returns an assigned lesson to Unassigned from its dropdown", async () => {
    vi.mocked(api.getLessonReadiness).mockResolvedValue({ ...readiness, is_ready: true });
    vi.mocked(api.getLessonSchedule).mockResolvedValue(generatedSchedule);
    const user = userEvent.setup();
    render(<LessonSchedulePage />);
    await screen.findByText("Placements");
    await user.selectOptions(screen.getByLabelText("Instructor for Ari Example"), "Unassigned");
    expect(api.removeLessonPlacement).toHaveBeenCalledWith("student-1");
    expect(api.manuallyUnplaceLessonStudent).not.toHaveBeenCalled();
  });
});
