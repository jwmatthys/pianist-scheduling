import { fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../../lib/api";
import { chooseLocalFile } from "../../lib/platform";
import type {
  LessonInstructor,
  LessonInstructorImportInspection,
  LessonInstructorImportPreview,
  LessonTag,
} from "../../lib/types";
import { LessonInstructorRosterPage } from "./LessonInstructorRosterPage";

vi.mock("../../lib/api", () => ({
  api: {
    listLessonInstructors: vi.fn(),
    listLessonTags: vi.fn(),
    createLessonInstructor: vi.fn(),
    updateLessonInstructor: vi.fn(),
    deleteLessonInstructor: vi.fn(),
    getLessonInstructorAvailability: vi.fn(),
    setLessonInstructorAvailability: vi.fn(),
    createLessonTag: vi.fn(),
    renameLessonTag: vi.fn(),
    deleteLessonTag: vi.fn(),
    inspectLessonInstructorFile: vi.fn(),
    previewLessonInstructorImport: vi.fn(),
    applyLessonInstructorImport: vi.fn(),
  },
}));
vi.mock("../../lib/platform", () => ({ chooseLocalFile: vi.fn() }));

const tag = (id: string, name: string, instructors = 0): LessonTag => ({
  id,
  name,
  instructor_count: instructors,
  student_count: 0,
  is_unused: instructors === 0,
});

const instructor = (overrides: Partial<LessonInstructor> = {}): LessonInstructor => ({
  id: "i1",
  name: "Ada Example",
  email: "ada@example.edu",
  room: "R1",
  lesson_length_minutes: 60,
  capacity: 1,
  tags: [],
  availability_window_count: 0,
  created_at: "2026-01-01T00:00:00",
  updated_at: "2026-01-01T00:00:00",
  ...overrides,
});

const inspection: LessonInstructorImportInspection = {
  sheets: [],
  selected_sheet: null,
  columns: ["Name", "Mon", "Specialty"],
  sample_rows: [],
  suggested_mapping: {
    instructor_name: "Name",
    instructor_email: null,
    room: null,
    lesson_length: null,
    capacity: null,
    days: { Monday: "Mon" },
  },
};

const preview: LessonInstructorImportPreview = {
  selected_sheet: null,
  tag_columns: ["Specialty"],
  total_rows: 1,
  valid_rows: 1,
  invalid_rows: 0,
  rows: [
    {
      row_number: 2,
      instructor_name: "Imported Person",
      instructor_email: "",
      room: "",
      lesson_length_minutes: 60,
      capacity: 1,
      tags: ["VOICE"],
      availability_window_count: 1,
      availability: ["Mon 9:00 AM-12:00 PM"],
      warnings: [],
      errors: [],
    },
  ],
  new_tags: ["VOICE"],
  reused_tags: [],
  existing_instructor_count: 4,
  existing_placement_count: 2,
  can_apply: true,
};

const cell = (dialog: HTMLElement, name: string) => within(dialog).getByRole("button", { name, hidden: true });

describe("Lesson Instructor Roster", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(api.listLessonInstructors).mockResolvedValue([]);
    vi.mocked(api.listLessonTags).mockResolvedValue([]);
    vi.mocked(api.getLessonInstructorAvailability).mockResolvedValue({ windows: [] });
  });

  it("validates then creates an instructor with defaults and selected tags", async () => {
    vi.mocked(api.listLessonTags).mockResolvedValue([tag("t1", "VOICE")]);
    vi.mocked(api.createLessonInstructor).mockResolvedValue(instructor());
    const user = userEvent.setup();
    render(<LessonInstructorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Add Instructor" }));
    const dialog = await screen.findByRole("dialog", { name: "Add Instructor" });

    expect((within(dialog).getByLabelText("Lesson Length (minutes)") as HTMLInputElement).value).toBe("60");
    expect((within(dialog).getByLabelText("Capacity") as HTMLInputElement).value).toBe("1");

    await user.click(within(dialog).getByRole("button", { name: "Add Instructor" }));
    expect(within(dialog).getByText("Instructor Name is required.")).toBeTruthy();
    expect(api.createLessonInstructor).not.toHaveBeenCalled();

    await user.type(within(dialog).getByLabelText("Instructor Name"), "New Person");
    await user.clear(within(dialog).getByLabelText("Capacity"));
    await user.type(within(dialog).getByLabelText("Capacity"), "0");
    await user.click(within(dialog).getByRole("button", { name: "Add Instructor" }));
    expect(within(dialog).getByText("Capacity must be greater than 0.")).toBeTruthy();

    await user.clear(within(dialog).getByLabelText("Capacity"));
    await user.type(within(dialog).getByLabelText("Capacity"), "2");
    await user.click(within(dialog).getByRole("checkbox"));
    await user.click(within(dialog).getByRole("button", { name: "Add Instructor" }));

    await waitFor(() =>
      expect(api.createLessonInstructor).toHaveBeenCalledWith({
        name: "New Person",
        email: "",
        room: "",
        lesson_length_minutes: 60,
        capacity: 2,
        tag_ids: ["t1"],
      }),
    );
  });

  it("edits an instructor and updates tags in the same save", async () => {
    vi.mocked(api.listLessonInstructors).mockResolvedValue([instructor({ tags: [{ id: "t1", name: "VOICE" }] })]);
    vi.mocked(api.listLessonTags).mockResolvedValue([tag("t1", "VOICE", 1), tag("t2", "PIANO")]);
    vi.mocked(api.updateLessonInstructor).mockResolvedValue(instructor());
    const user = userEvent.setup();
    render(<LessonInstructorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Edit Ada Example" }));
    const dialog = await screen.findByRole("dialog", { name: "Edit Instructor" });
    const name = within(dialog).getByLabelText("Instructor Name");
    await user.clear(name);
    await user.type(name, "Ada Changed");
    const boxes = within(dialog).getAllByRole("checkbox");
    await user.click(boxes[0]);
    await user.click(boxes[1]);
    await user.click(within(dialog).getByRole("button", { name: "Save Changes" }));
    await waitFor(() =>
      expect(api.updateLessonInstructor).toHaveBeenCalledWith("i1", {
        name: "Ada Changed",
        email: "ada@example.edu",
        room: "R1",
        lesson_length_minutes: 60,
        capacity: 1,
        tag_ids: ["t2"],
      }),
    );
  });

  it("deletes an instructor only after confirmation", async () => {
    vi.mocked(api.listLessonInstructors).mockResolvedValue([instructor()]);
    vi.mocked(api.deleteLessonInstructor).mockResolvedValue({ deleted: true, placements_preserved: 0 } as never);
    const user = userEvent.setup();
    render(<LessonInstructorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Delete Ada Example" }));
    const dialog = await screen.findByRole("dialog", { name: "Delete Ada Example?" });
    await user.click(within(dialog).getByRole("button", { name: "Cancel" }));
    expect(api.deleteLessonInstructor).not.toHaveBeenCalled();

    await user.click(screen.getByRole("button", { name: "Delete Ada Example" }));
    await user.click(within(await screen.findByRole("dialog", { name: "Delete Ada Example?" })).getByRole("button", { name: "Delete Instructor" }));
    await waitFor(() => expect(api.deleteLessonInstructor).toHaveBeenCalledWith("i1"));
  });

  it("manages tags: greys unused, creates, renames, and warns before deleting a tag in use", async () => {
    vi.mocked(api.listLessonTags).mockResolvedValue([tag("t1", "VOICE", 2), tag("t2", "UNUSED")]);
    vi.mocked(api.createLessonTag).mockResolvedValue(tag("t3", "NEW"));
    vi.mocked(api.renameLessonTag).mockResolvedValue(tag("t2", "RENAMED"));
    vi.mocked(api.deleteLessonTag).mockResolvedValue({} as never);
    const user = userEvent.setup();
    render(<LessonInstructorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Manage Tags" }));
    const dialog = await screen.findByRole("dialog", { name: "Manage Tags" });

    expect(within(dialog).getByText("UNUSED").className).toContain("clinical-tag-unused");
    expect(within(dialog).getByText("VOICE").className).not.toContain("clinical-tag-unused");

    await user.type(within(dialog).getByLabelText("New tag"), "new");
    await user.click(within(dialog).getByRole("button", { name: "Add Tag" }));
    await waitFor(() => expect(api.createLessonTag).toHaveBeenCalledWith("new"));

    await user.click(within(dialog).getByRole("button", { name: "Rename UNUSED" }));
    const input = within(dialog).getByLabelText("New name for UNUSED");
    await user.clear(input);
    await user.type(input, "renamed");
    await user.click(within(dialog).getByRole("button", { name: "Save" }));
    await waitFor(() => expect(api.renameLessonTag).toHaveBeenCalledWith("t2", "renamed"));

    await user.click(within(dialog).getByRole("button", { name: "Delete VOICE" }));
    expect(await screen.findByText(/currently used by 2 instructors/)).toBeTruthy();
    await user.click(screen.getByRole("button", { name: "Delete Tag" }));
    await waitFor(() => expect(api.deleteLessonTag).toHaveBeenCalledWith("t1"));
  });

  it("edits availability from the row and from the edit dialog, saving recurring windows", async () => {
    vi.mocked(api.listLessonInstructors).mockResolvedValue([instructor()]);
    vi.mocked(api.getLessonInstructorAvailability).mockResolvedValue({
      windows: [{ day: "Monday", start_minute: 540, end_minute: 600, status: "Available" }],
    });
    vi.mocked(api.setLessonInstructorAvailability).mockResolvedValue({ windows: [] });
    const user = userEvent.setup();
    render(<LessonInstructorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Availability for Ada Example" }));
    const dialog = await screen.findByRole("dialog", { name: "Availability: Ada Example", hidden: true });
    await waitFor(() => expect(cell(dialog, "Monday 9:00 AM: Available")).toBeTruthy());
    fireEvent.mouseDown(cell(dialog, "Monday 10:00 AM: Unmarked"), { button: 0 });
    fireEvent.mouseUp(window);
    await user.click(within(dialog).getByRole("button", { name: "Save Availability", hidden: true }));
    await waitFor(() => expect(api.setLessonInstructorAvailability).toHaveBeenCalled());
    const [id, windows] = vi.mocked(api.setLessonInstructorAvailability).mock.calls[0];
    expect(id).toBe("i1");
    expect(windows.filter((item) => item.day === "Monday")).toEqual([
      { day: "Monday", start_minute: 420, end_minute: 540, status: "Unavailable" },
      { day: "Monday", start_minute: 540, end_minute: 630, status: "Available" },
      { day: "Monday", start_minute: 630, end_minute: 1260, status: "Unavailable" },
    ]);

    await user.click(screen.getByRole("button", { name: "Edit Ada Example" }));
    const edit = await screen.findByRole("dialog", { name: "Edit Instructor" });
    await user.click(within(edit).getByRole("button", { name: "Edit Availability" }));
    expect(await screen.findByRole("dialog", { name: "Availability: Ada Example", hidden: true })).toBeTruthy();
  });

  it("imports: choose file, map, preview, confirm replacement text, then apply", async () => {
    vi.mocked(chooseLocalFile).mockResolvedValue(new File(["x"], "i.csv"));
    vi.mocked(api.inspectLessonInstructorFile).mockResolvedValue(inspection);
    vi.mocked(api.previewLessonInstructorImport).mockResolvedValue(preview);
    vi.mocked(api.applyLessonInstructorImport).mockResolvedValue({
      instructors_removed: 4,
      placements_removed: 0,
      placements_preserved: 2,
      instructors_created: 1,
      tags_created: 1,
    });
    const user = userEvent.setup();
    render(<LessonInstructorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Import Instructor Roster" }));
    const dialog = await screen.findByRole("dialog", { name: "Import Instructor Roster" });
    await user.click(within(dialog).getByRole("button", { name: "Choose spreadsheet" }));
    expect(await within(dialog).findByLabelText("Instructor Name")).toBeTruthy();
    await user.click(within(dialog).getByRole("button", { name: "Preview Import" }));
    expect(await within(dialog).findByText("Imported Person")).toBeTruthy();
    expect(within(dialog).getByText("Mon 9:00 AM-12:00 PM")).toBeTruthy();
    expect(api.applyLessonInstructorImport).not.toHaveBeenCalled();

    await user.click(within(dialog).getByRole("button", { name: "Import Instructors" }));
    const confirm = await screen.findByRole("dialog", { name: "Replace Instructors?" });
    expect(within(confirm).getByText(/Delete all current instructors \(4 currently\)/)).toBeTruthy();
    expect(within(confirm).getByText("Preserve the Lesson Tag registry")).toBeTruthy();
    expect(within(confirm).getByText("Preserve Student records")).toBeTruthy();
    expect(within(confirm).getByText("Clear lesson schedules")).toBeTruthy();
    expect(api.applyLessonInstructorImport).not.toHaveBeenCalled();

    await user.click(within(confirm).getByRole("button", { name: "Replace Instructors" }));
    await waitFor(() => expect(api.applyLessonInstructorImport).toHaveBeenCalled());
    expect(await screen.findByRole("status")).toBeTruthy();
  });

  it("blocks apply when the preview has errors", async () => {
    vi.mocked(chooseLocalFile).mockResolvedValue(new File(["x"], "i.csv"));
    vi.mocked(api.inspectLessonInstructorFile).mockResolvedValue(inspection);
    vi.mocked(api.previewLessonInstructorImport).mockResolvedValue({
      ...preview,
      valid_rows: 0,
      invalid_rows: 1,
      can_apply: false,
      rows: [{ ...preview.rows[0], errors: ["Instructor Name is required."] }],
    });
    const user = userEvent.setup();
    render(<LessonInstructorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Import Instructor Roster" }));
    const dialog = await screen.findByRole("dialog", { name: "Import Instructor Roster" });
    await user.click(within(dialog).getByRole("button", { name: "Choose spreadsheet" }));
    await user.click(await within(dialog).findByRole("button", { name: "Preview Import" }));
    const apply = await within(dialog).findByRole("button", { name: "Import Instructors" });
    expect((apply as HTMLButtonElement).disabled).toBe(true);
  });
});
