import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../lib/api";
import { SCHEDULING_MODULES } from "../moduleRegistry";
import { LessonModule } from "./LessonModule";

vi.mock("../lib/api", () => ({
  api: { listLessonInstructors: vi.fn(), listLessonTags: vi.fn(), listLessonStudents: vi.fn() },
}));
vi.mock("../lib/platform", () => ({ chooseLocalFile: vi.fn() }));

beforeEach(() => {
  vi.mocked(api.listLessonInstructors).mockResolvedValue([]);
  vi.mocked(api.listLessonTags).mockResolvedValue([]);
  vi.mocked(api.listLessonStudents).mockResolvedValue([]);
});

describe("Lesson Scheduling module", () => {
  it("is registered as an available module independent of Clinical", () => {
    const module = SCHEDULING_MODULES.find((item) => item.key === "lesson");
    expect(module).toMatchObject({ name: "Lesson Scheduling", availability: "available" });
    expect(SCHEDULING_MODULES.filter((item) => item.key === "lesson")).toHaveLength(1);
    expect(SCHEDULING_MODULES.map((item) => item.key)).toContain("clinical");
  });

  it("shows the four tabs in order and starts on the Instructor Roster", () => {
    render(<LessonModule />);
    const labels = screen
      .getAllByRole("button")
      .map((button) => button.textContent)
      .slice(0, 4);
    expect(labels).toEqual(["Instructor Roster", "Student Roster", "Schedule", "Reports"]);
    expect(screen.getByRole("heading", { name: "Instructor Roster" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Instructor Roster" }).getAttribute("aria-current")).toBe("page");
  });

  it("renders the empty Instructor Roster with its actions", async () => {
    render(<LessonModule />);
    expect(await screen.findByText("No instructors yet.")).toBeTruthy();
    expect(screen.getByRole("button", { name: "Add Instructor" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Import Instructor Roster" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Manage Tags" })).toBeTruthy();
  });

  it("renders the Student Roster, Schedule, and Reports placeholders", async () => {
    const user = userEvent.setup();
    render(<LessonModule />);

    await user.click(screen.getByRole("button", { name: "Student Roster" }));
    expect(screen.getByRole("heading", { name: "Student Roster" })).toBeTruthy();
    expect(screen.getByText("No students yet.")).toBeTruthy();
    expect(screen.getByRole("button", { name: "Add Student" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "Import Student Roster" })).toBeTruthy();

    await user.click(screen.getByRole("button", { name: "Schedule" }));
    expect(screen.getByText("Scheduling not yet implemented.")).toBeTruthy();

    await user.click(screen.getByRole("button", { name: "Reports" }));
    expect(screen.getByText("Reports not yet implemented.")).toBeTruthy();
  });
});
