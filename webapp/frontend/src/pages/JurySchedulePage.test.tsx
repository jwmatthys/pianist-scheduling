import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { JuryReadiness, JuryScheduleHistoryItem, JuryScheduleResult } from "../lib/types";

const apiMocks = vi.hoisted(() => ({
  getJuryReadiness: vi.fn(),
  generateJurySchedule: vi.fn(),
  getCurrentJurySchedule: vi.fn(),
  getJuryScheduleHistory: vi.fn(),
  getJuryScheduleResult: vi.fn(),
}));

vi.mock("../lib/api", () => ({ api: apiMocks }));

import { JurySchedulePage } from "./JurySchedulePage";

function readyReadiness(): JuryReadiness {
  return {
    ready: true,
    source_result_uuid: "source-result-1",
    source_revision: 8,
    jury_input_revision: 4,
    issues: [],
  };
}

function blockedReadiness(): JuryReadiness {
  return {
    ready: false,
    source_result_uuid: "source-result-1",
    source_revision: 8,
    jury_input_revision: 4,
    issues: [
      {
        code: "PANEL_REQUIRED",
        severity: "error",
        message: "Select a Jury Panel for this lesson.",
        entity_uuids: ["lesson-1"],
      },
      {
        code: "PIANIST_AVAILABILITY_INCOMPLETE",
        severity: "error",
        message: "Assigned Pianist has no Jury Availability Windows for the Panel date.",
        entity_uuids: ["lesson-1", "pianist-1"],
      },
      {
        code: "UNUSED_PANEL",
        severity: "warning",
        message: "Panel B has no Jury-required lessons.",
        entity_uuids: ["panel-2"],
      },
      {
        code: "NO_JURY_REQUIRED_ENTRIES",
        severity: "warning",
        message: "No Jury lesson entries are currently required.",
        entity_uuids: [],
      },
    ],
  };
}

function scheduleResult(overrides: Partial<JuryScheduleResult> = {}): JuryScheduleResult {
  return {
    result_uuid: "jury-result-2",
    session_uuid: "session-1",
    contract_id: "jury.schedule-result",
    contract_version: 1,
    result_version: 2,
    state: "draft",
    created_at: "2026-10-02T14:00:00Z",
    source_result_uuid: "source-result-1",
    source_contract_id: "accompanist.assignment-result",
    source_contract_version: 2,
    source_revision: 8,
    jury_input_revision: 4,
    stale: false,
    stale_reasons: [],
    panel_timelines: [{
      panel_uuid: "panel-1",
      panel_name: "Voice Panel",
      jury_date: "2026-10-15",
      events: [{
        kind: "jury",
        source_lesson_uuid: "lesson-1",
        student_person_uuid: "student-1",
        student_display_name: "Synthetic Student",
        instrument: "Voice",
        panel_uuid: "panel-1",
        jury_date: "2026-10-15",
        pianist_person_uuid: "pianist-1",
        pianist_display_name: "Synthetic Pianist",
        start_minute: 540,
        end_minute: 600,
      }],
    }],
    unscheduled_lessons: [],
    warnings: [],
    ...overrides,
  };
}

function historyItem(overrides: Partial<JuryScheduleHistoryItem> = {}): JuryScheduleHistoryItem {
  return {
    result_uuid: "jury-result-1",
    session_uuid: "session-1",
    result_version: 1,
    state: "superseded",
    created_at: "2026-10-01T14:00:00Z",
    source_result_uuid: "source-result-1",
    source_revision: 7,
    jury_input_revision: 3,
    stale: true,
    stale_reasons: ["accompanist_source_revision_changed"],
    scheduled_count: 0,
    unscheduled_count: 1,
    ...overrides,
  };
}

function renderSchedule(readiness = readyReadiness()) {
  return render(<JurySchedulePage readiness={readiness} onReadinessChange={vi.fn()} />);
}

describe("Jury Schedule view", () => {
  beforeEach(() => {
    vi.resetAllMocks();
    apiMocks.getCurrentJurySchedule.mockResolvedValue(null);
    apiMocks.getJuryScheduleHistory.mockResolvedValue([]);
    apiMocks.getJuryReadiness.mockResolvedValue(readyReadiness());
    apiMocks.generateJurySchedule.mockResolvedValue(scheduleResult());
    apiMocks.getJuryScheduleResult.mockResolvedValue(scheduleResult());
  });

  it("displays readiness blockers and warnings", async () => {
    renderSchedule(blockedReadiness());

    expect(await screen.findByText("Select a Jury Panel for this lesson.")).toBeTruthy();
    expect(screen.getByText("Assigned Pianist has no Jury Availability Windows for the Panel date.")).toBeTruthy();
    expect(screen.getByText("Panel Required")).toBeTruthy();
    expect(screen.getByText("Panel B has no Jury-required lessons.")).toBeTruthy();
    expect(screen.getByText("No Jury lesson entries are currently required.")).toBeTruthy();
  });

  it("disables generation while readiness blockers exist", async () => {
    renderSchedule(blockedReadiness());

    const button = await screen.findByRole("button", { name: "Generate Schedule" });
    expect(button).toHaveProperty("disabled", true);
    expect(apiMocks.generateJurySchedule).not.toHaveBeenCalled();
  });

  it("rechecks readiness and aborts generation if blockers appear before POST", async () => {
    const onReadinessChange = vi.fn();
    apiMocks.getJuryReadiness.mockResolvedValueOnce(blockedReadiness());

    render(<JurySchedulePage readiness={readyReadiness()} onReadinessChange={onReadinessChange} />);
    fireEvent.click(await screen.findByRole("button", { name: "Generate Schedule" }));

    await waitFor(() => expect(onReadinessChange).toHaveBeenCalled());
    expect(onReadinessChange).toHaveBeenCalledWith(blockedReadiness());
    expect(apiMocks.generateJurySchedule).not.toHaveBeenCalled();
  });

  it("generates a schedule and refreshes the current result and history", async () => {
    const generated = scheduleResult();
    apiMocks.getCurrentJurySchedule
      .mockResolvedValueOnce(null)
      .mockResolvedValue(generated);
    apiMocks.getJuryScheduleHistory
      .mockResolvedValueOnce([])
      .mockResolvedValue([historyItem({
        result_uuid: generated.result_uuid,
        result_version: generated.result_version,
        state: generated.state,
        stale: false,
        stale_reasons: [],
        scheduled_count: 1,
        unscheduled_count: 0,
      })]);
    apiMocks.generateJurySchedule.mockResolvedValue(generated);

    renderSchedule();
    fireEvent.click(await screen.findByRole("button", { name: "Generate Schedule" }));

    await waitFor(() => expect(apiMocks.generateJurySchedule).toHaveBeenCalledWith(4));
    expect(await screen.findByText("Synthetic Student")).toBeTruthy();
    expect(screen.getByRole("region", { name: "Current generated schedule" })).toBeTruthy();
    expect(screen.getAllByText("Draft", { selector: ".jury-schedule-status" }).length).toBeGreaterThan(0);
    expect(apiMocks.getCurrentJurySchedule).toHaveBeenCalledTimes(2);
    expect(apiMocks.getJuryScheduleHistory).toHaveBeenCalledTimes(2);
  });

  it("prominently displays stale status returned by the backend", async () => {
    apiMocks.getCurrentJurySchedule.mockResolvedValue(scheduleResult({
      stale: true,
      stale_reasons: ["accompanist_result_changed"],
    }));

    renderSchedule();

    expect(await screen.findByText("Stale", { selector: ".jury-schedule-status" })).toBeTruthy();
    expect(screen.getByRole("status").textContent).toContain("This schedule uses older source data.");
  });

  it("renders a backend Finalized state without inferring Current lifecycle state", async () => {
    apiMocks.getCurrentJurySchedule.mockResolvedValue(scheduleResult({ state: "finalized" }));

    renderSchedule();

    expect(await screen.findByRole("region", { name: "Current generated schedule" })).toBeTruthy();
    expect(screen.getByText("Finalized", { selector: ".jury-schedule-status" })).toBeTruthy();
    expect(screen.queryByText("Current", { selector: ".jury-schedule-status" })).toBeNull();
  });

  it("displays unscheduled lessons separately with reason and explanation", async () => {
    apiMocks.getCurrentJurySchedule.mockResolvedValue(scheduleResult({
      unscheduled_lessons: [{
        source_lesson_uuid: "lesson-unscheduled",
        student_person_uuid: "student-2",
        student_display_name: "Unscheduled Student",
        instrument: "Cello",
        panel_uuid: "panel-1",
        jury_date: "2026-10-15",
        reason_code: "fixed_pianist_conflict",
        explanation: "The assigned Pianist is occupied for the remaining interval.",
      }],
    }));

    renderSchedule();

    expect(await screen.findByRole("region", { name: "Unscheduled lessons" })).toBeTruthy();
    expect(screen.getByText("Unscheduled Student")).toBeTruthy();
    expect(screen.getByText("Fixed Pianist Conflict")).toBeTruthy();
    expect(screen.getByText("The assigned Pianist is occupied for the remaining interval.")).toBeTruthy();
  });

  it("renders Meal Break, Periodic Break, and optimizer warnings from the result", async () => {
    apiMocks.getCurrentJurySchedule.mockResolvedValue(scheduleResult({
      panel_timelines: [{
        panel_uuid: "panel-1",
        panel_name: "Voice Panel",
        jury_date: "2026-10-15",
        events: [
          { kind: "meal_break", panel_uuid: "panel-1", jury_date: "2026-10-15", start_minute: 720, end_minute: 750 },
          { kind: "periodic_break", panel_uuid: "panel-1", jury_date: "2026-10-15", start_minute: 800, end_minute: 810, after_jury_count: 3 },
        ],
      }],
      warnings: [{
        code: "SCHEDULE_WARNING",
        severity: "warning",
        message: "Backend schedule warning for this Panel.",
      }],
    }));

    renderSchedule();

    expect(await screen.findByText("Meal Break")).toBeTruthy();
    expect(screen.getByText("Periodic Break · after 3 juries")).toBeTruthy();
    expect(screen.getByText("Backend schedule warning for this Panel.")).toBeTruthy();
  });

  it("loads and displays historical results without changing the current result", async () => {
    const historical = scheduleResult({
      result_uuid: "jury-result-1",
      result_version: 1,
      state: "superseded",
      stale: true,
      stale_reasons: ["jury_inputs_changed"],
      panel_timelines: [{
        panel_uuid: "panel-1",
        panel_name: "Historical Panel",
        jury_date: "2026-10-10",
        events: [{
          kind: "jury",
          source_lesson_uuid: "lesson-old",
          student_person_uuid: "student-old",
          student_display_name: "Archived Student",
          instrument: "Violin",
          panel_uuid: "panel-1",
          jury_date: "2026-10-10",
          pianist_person_uuid: null,
          pianist_display_name: null,
          start_minute: 600,
          end_minute: 630,
        }],
      }],
    });
    apiMocks.getCurrentJurySchedule.mockResolvedValue(scheduleResult());
    apiMocks.getJuryScheduleHistory.mockResolvedValue([historyItem({ state: "superseded" })]);
    apiMocks.getJuryScheduleResult.mockResolvedValue(historical);

    renderSchedule();
    fireEvent.click(await screen.findByRole("button", { name: /Result v1/ }));

    expect(await screen.findByRole("region", { name: "Saved schedule v1" })).toBeTruthy();
    expect(screen.getByText("Archived Student")).toBeTruthy();
    expect(screen.getAllByText("Superseded", { selector: ".jury-schedule-status" }).length).toBeGreaterThan(0);
    expect(screen.getAllByText("Stale", { selector: ".jury-schedule-status" }).length).toBeGreaterThan(0);
    expect(screen.getByRole("region", { name: "Current generated schedule" })).toBeTruthy();
    expect(apiMocks.getJuryScheduleResult).toHaveBeenCalledWith("jury-result-1");
    expect(apiMocks.generateJurySchedule).not.toHaveBeenCalled();
    expect(apiMocks.getCurrentJurySchedule).toHaveBeenCalledTimes(1);
  });

  it("renders a non-stale Finalized history state exactly as returned by the backend", async () => {
    apiMocks.getJuryScheduleHistory.mockResolvedValue([historyItem({
      state: "finalized",
      stale: false,
      stale_reasons: [],
    })]);

    renderSchedule();

    expect(await screen.findByText("Finalized", { selector: ".jury-schedule-status" })).toBeTruthy();
    expect(screen.queryByText("Current", { selector: ".jury-schedule-status" })).toBeNull();
  });

  it("shows explicit empty-history and empty-current-result states", async () => {
    renderSchedule();

    expect(await screen.findByText("No Jury schedule generated yet")).toBeTruthy();
    expect(screen.getByText("No generated schedules in this session.")).toBeTruthy();
  });

  it("keeps the current result visible when history retrieval fails", async () => {
    apiMocks.getCurrentJurySchedule.mockResolvedValue(scheduleResult());
    apiMocks.getJuryScheduleHistory.mockRejectedValue(new Error("503: History unavailable."));

    renderSchedule();

    expect(await screen.findByText("Synthetic Student")).toBeTruthy();
    expect(await screen.findByRole("alert")).toBeTruthy();
    expect(screen.getByRole("alert").textContent).toContain("503: History unavailable.");
    expect(screen.getByRole("region", { name: "Current generated schedule" })).toBeTruthy();
  });

  it("keeps schedule history visible when current-result retrieval fails", async () => {
    apiMocks.getCurrentJurySchedule.mockRejectedValue(new Error("503: Current result unavailable."));
    apiMocks.getJuryScheduleHistory.mockResolvedValue([historyItem()]);

    renderSchedule();

    expect(await screen.findByRole("button", { name: /Result v1/ })).toBeTruthy();
    expect(await screen.findByRole("alert")).toBeTruthy();
    expect(screen.getByRole("alert").textContent).toContain("503: Current result unavailable.");
    expect(screen.getByText("No Jury schedule generated yet")).toBeTruthy();
  });

  it("explains a current result with no scheduled or unscheduled lessons", async () => {
    apiMocks.getCurrentJurySchedule.mockResolvedValue(scheduleResult({
      panel_timelines: [],
      unscheduled_lessons: [],
      warnings: [],
    }));

    renderSchedule();

    expect(await screen.findByText("No Jury-required lessons were included in this result.")).toBeTruthy();
    expect(screen.queryByText("All Jury-required lessons have a scheduled time.")).toBeNull();
  });

  it("shows a clean readiness empty state when there are no warnings", async () => {
    renderSchedule(readyReadiness());

    expect(await screen.findByText("No readiness blockers or warnings.")).toBeTruthy();
    expect(screen.queryByLabelText("Blocking readiness issues")).toBeNull();
    expect(screen.queryByLabelText("Readiness warnings")).toBeNull();
  });

  it("shows generation API failures without hiding the Schedule view", async () => {
    apiMocks.getCurrentJurySchedule.mockResolvedValue(scheduleResult());
    apiMocks.getJuryScheduleHistory.mockResolvedValue([historyItem()]);
    const historical = scheduleResult({ result_uuid: "jury-result-1", result_version: 1, stale: true });
    apiMocks.getJuryScheduleResult.mockResolvedValue(historical);
    apiMocks.generateJurySchedule.mockRejectedValue(new Error("409: Readiness changed; refresh and retry."));

    renderSchedule();
    fireEvent.click(await screen.findByRole("button", { name: /Result v1/ }));
    expect(await screen.findByRole("region", { name: "Saved schedule v1" })).toBeTruthy();
    fireEvent.click(await screen.findByRole("button", { name: "Generate Schedule" }));

    expect((await screen.findByRole("alert")).textContent).toContain("409: Readiness changed; refresh and retry.");
    expect(screen.getByRole("region", { name: "Saved schedule v1" })).toBeTruthy();
    expect(screen.getByRole("region", { name: "Current generated schedule" })).toBeTruthy();
  });
});
