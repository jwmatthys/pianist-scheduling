import { fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../../lib/api";
import { chooseLocalFile } from "../../lib/platform";
import type {
  TutorTutor,
  TutorTutorImportInspection,
  TutorTutorImportPreview,
  TutorTag,
} from "../../lib/types";
import { TutorTutorRosterPage } from "./TutorTutorRosterPage";

vi.mock("../../lib/api", () => ({
  api: {
    listTutorTutors: vi.fn(),
    listTutorTags: vi.fn(),
    createTutorTutor: vi.fn(),
    updateTutorTutor: vi.fn(),
    deleteTutorTutor: vi.fn(),
    getTutorTutorAvailability: vi.fn(),
    setTutorTutorAvailability: vi.fn(),
    createTutorTag: vi.fn(),
    renameTutorTag: vi.fn(),
    deleteTutorTag: vi.fn(),
    inspectTutorTutorFile: vi.fn(),
    previewTutorTutorImport: vi.fn(),
    applyTutorTutorImport: vi.fn(),
  },
}));
vi.mock("../../lib/platform", () => ({ chooseLocalFile: vi.fn() }));

const tag = (id: string, name: string, tutors = 0): TutorTag => ({
  id,
  name,
  tutor_count: tutors,
  student_count: 0,
  is_unused: tutors === 0,
});

const tutor = (overrides: Partial<TutorTutor> = {}): TutorTutor => ({
  id: "i1",
  name: "Ada Example",
  email: "ada@example.edu",
  room: "R1",
  lesson_length_minutes: 60,
  capacity: 1,
  tags: [],
  availability_window_count: 0,
  available_minutes: 0,
  tentative_minutes: 0,
  created_at: "2026-01-01T00:00:00",
  updated_at: "2026-01-01T00:00:00",
  ...overrides,
});

const inspection: TutorTutorImportInspection = {
  sheets: [],
  selected_sheet: null,
  columns: ["Name", "Mon", "Specialty"],
  sample_rows: [],
  suggested_mapping: {
    tutor_name: "Name",
    tutor_email: null,
    room: null,
    lesson_length: null,
    capacity: null,
    days: { Monday: "Mon" },
    tags: "Specialty",
  },
};

const preview: TutorTutorImportPreview = {
  selected_sheet: null,
  tag_columns: ["Specialty"],
  total_rows: 1,
  valid_rows: 1,
  invalid_rows: 0,
  rows: [
    {
      row_number: 2,
      tutor_name: "Imported Person",
      tutor_email: "",
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
  existing_tutor_count: 4,
  existing_placement_count: 2,
  can_apply: true,
};

const cell = (dialog: HTMLElement, name: string) => within(dialog).getByRole("button", { name, hidden: true });

describe("Tutor Roster", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(api.listTutorTutors).mockResolvedValue([]);
    vi.mocked(api.listTutorTags).mockResolvedValue([]);
    vi.mocked(api.getTutorTutorAvailability).mockResolvedValue({ windows: [] });
  });

  it("validates then creates a tutor with defaults and selected tags", async () => {
    vi.mocked(api.listTutorTags).mockResolvedValue([tag("t1", "VOICE")]);
    vi.mocked(api.createTutorTutor).mockResolvedValue(tutor());
    const user = userEvent.setup();
    render(<TutorTutorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Add Tutor" }));
    const dialog = await screen.findByRole("dialog", { name: "Add Tutor" });

    expect((within(dialog).getByLabelText("Lesson Length (minutes)") as HTMLInputElement).value).toBe("60");
    expect((within(dialog).getByLabelText("Capacity") as HTMLInputElement).value).toBe("1");

    await user.click(within(dialog).getByRole("button", { name: "Add Tutor" }));
    expect(within(dialog).getByText("Tutor Name is required.")).toBeTruthy();
    expect(api.createTutorTutor).not.toHaveBeenCalled();

    await user.type(within(dialog).getByLabelText("Tutor Name"), "New Person");
    await user.clear(within(dialog).getByLabelText("Capacity"));
    await user.type(within(dialog).getByLabelText("Capacity"), "0");
    await user.click(within(dialog).getByRole("button", { name: "Add Tutor" }));
    expect(within(dialog).getByText("Capacity must be greater than 0.")).toBeTruthy();

    await user.clear(within(dialog).getByLabelText("Capacity"));
    await user.type(within(dialog).getByLabelText("Capacity"), "2");
    await user.click(within(dialog).getByRole("checkbox"));
    await user.click(within(dialog).getByRole("button", { name: "Add Tutor" }));

    await waitFor(() =>
      expect(api.createTutorTutor).toHaveBeenCalledWith({
        name: "New Person",
        email: "",
        room: "",
        lesson_length_minutes: 60,
        capacity: 2,
        tag_ids: ["t1"],
      }),
    );
  });

  it("summarizes available and tentative hours on separate lines instead of a window count", async () => {
    vi.mocked(api.listTutorTutors).mockResolvedValue([
      tutor({ id: "a", name: "Ada Example", availability_window_count: 3, available_minutes: 450, tentative_minutes: 60 }),
      tutor({ id: "b", name: "Ben Example", availability_window_count: 0 }),
    ]);
    render(<TutorTutorRosterPage />);
    await screen.findByText("Ada Example");
    expect(screen.getByText("Available: 7.5 hrs")).toBeTruthy();
    expect(screen.getByText("Tentative: 1 hr")).toBeTruthy();
    expect(screen.queryByText(/windows?$/)).toBeNull();
    expect(screen.getByText("Not set")).toBeTruthy();
  });

  it("edits a tutor and updates tags in the same save", async () => {
    vi.mocked(api.listTutorTutors).mockResolvedValue([tutor({ tags: [{ id: "t1", name: "VOICE" }] })]);
    vi.mocked(api.listTutorTags).mockResolvedValue([tag("t1", "VOICE", 1), tag("t2", "PIANO")]);
    vi.mocked(api.updateTutorTutor).mockResolvedValue(tutor());
    const user = userEvent.setup();
    render(<TutorTutorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Edit Ada Example" }));
    const dialog = await screen.findByRole("dialog", { name: "Edit Tutor" });
    const name = within(dialog).getByLabelText("Tutor Name");
    await user.clear(name);
    await user.type(name, "Ada Changed");
    const boxes = within(dialog).getAllByRole("checkbox");
    await user.click(boxes[0]);
    await user.click(boxes[1]);
    await user.click(within(dialog).getByRole("button", { name: "Save Changes" }));
    await waitFor(() =>
      expect(api.updateTutorTutor).toHaveBeenCalledWith("i1", {
        name: "Ada Changed",
        email: "ada@example.edu",
        room: "R1",
        lesson_length_minutes: 60,
        capacity: 1,
        tag_ids: ["t2"],
      }),
    );
  });

  it("deletes a tutor only after confirmation", async () => {
    vi.mocked(api.listTutorTutors).mockResolvedValue([tutor()]);
    vi.mocked(api.deleteTutorTutor).mockResolvedValue({ deleted: true, placements_preserved: 0 } as never);
    const user = userEvent.setup();
    render(<TutorTutorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Delete Ada Example" }));
    const dialog = await screen.findByRole("dialog", { name: "Delete Ada Example?" });
    await user.click(within(dialog).getByRole("button", { name: "Cancel" }));
    expect(api.deleteTutorTutor).not.toHaveBeenCalled();

    await user.click(screen.getByRole("button", { name: "Delete Ada Example" }));
    await user.click(within(await screen.findByRole("dialog", { name: "Delete Ada Example?" })).getByRole("button", { name: "Delete Tutor" }));
    await waitFor(() => expect(api.deleteTutorTutor).toHaveBeenCalledWith("i1"));
  });

  it("manages tags: greys unused, creates, renames, and warns before deleting a tag in use", async () => {
    vi.mocked(api.listTutorTags).mockResolvedValue([tag("t1", "VOICE", 2), tag("t2", "UNUSED")]);
    vi.mocked(api.createTutorTag).mockResolvedValue(tag("t3", "NEW"));
    vi.mocked(api.renameTutorTag).mockResolvedValue(tag("t2", "RENAMED"));
    vi.mocked(api.deleteTutorTag).mockResolvedValue({} as never);
    const user = userEvent.setup();
    render(<TutorTutorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Manage Tags" }));
    const dialog = await screen.findByRole("dialog", { name: "Manage Tags" });

    expect(within(dialog).getByText("UNUSED").className).toContain("clinical-tag-unused");
    expect(within(dialog).getByText("VOICE").className).not.toContain("clinical-tag-unused");

    await user.type(within(dialog).getByLabelText("New tag"), "new");
    await user.click(within(dialog).getByRole("button", { name: "Add Tag" }));
    await waitFor(() => expect(api.createTutorTag).toHaveBeenCalledWith("new"));

    await user.click(within(dialog).getByRole("button", { name: "Rename UNUSED" }));
    const input = within(dialog).getByLabelText("New name for UNUSED");
    await user.clear(input);
    await user.type(input, "renamed");
    await user.click(within(dialog).getByRole("button", { name: "Save" }));
    await waitFor(() => expect(api.renameTutorTag).toHaveBeenCalledWith("t2", "renamed"));

    await user.click(within(dialog).getByRole("button", { name: "Delete VOICE" }));
    expect(await screen.findByText(/currently used by 2 tutors/)).toBeTruthy();
    await user.click(screen.getByRole("button", { name: "Delete Tag" }));
    await waitFor(() => expect(api.deleteTutorTag).toHaveBeenCalledWith("t1"));
  });

  it("edits availability from the row and from the edit dialog, saving recurring windows", async () => {
    vi.mocked(api.listTutorTutors).mockResolvedValue([tutor()]);
    vi.mocked(api.getTutorTutorAvailability).mockResolvedValue({
      windows: [{ day: "Monday", start_minute: 540, end_minute: 600, status: "Available" }],
    });
    vi.mocked(api.setTutorTutorAvailability).mockResolvedValue({ windows: [] });
    const user = userEvent.setup();
    render(<TutorTutorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Availability for Ada Example" }));
    const dialog = await screen.findByRole("dialog", { name: "Availability: Ada Example", hidden: true });
    await waitFor(() => expect(cell(dialog, "Monday 9:00 AM: Available")).toBeTruthy());
    fireEvent.mouseDown(cell(dialog, "Monday 10:00 AM: Unmarked"), { button: 0 });
    fireEvent.mouseUp(window);
    await user.click(within(dialog).getByRole("button", { name: "Save Availability", hidden: true }));
    await waitFor(() => expect(api.setTutorTutorAvailability).toHaveBeenCalled());
    const [id, windows] = vi.mocked(api.setTutorTutorAvailability).mock.calls[0];
    expect(id).toBe("i1");
    expect(windows.filter((item) => item.day === "Monday")).toEqual([
      { day: "Monday", start_minute: 420, end_minute: 540, status: "Unavailable" },
      { day: "Monday", start_minute: 540, end_minute: 630, status: "Available" },
      { day: "Monday", start_minute: 630, end_minute: 1260, status: "Unavailable" },
    ]);

    await user.click(screen.getByRole("button", { name: "Edit Ada Example" }));
    const edit = await screen.findByRole("dialog", { name: "Edit Tutor" });
    await user.click(within(edit).getByRole("button", { name: "Edit Availability" }));
    expect(await screen.findByRole("dialog", { name: "Availability: Ada Example", hidden: true })).toBeTruthy();
  });

  it("imports: choose file, map, preview, confirm replacement text, then apply", async () => {
    vi.mocked(chooseLocalFile).mockResolvedValue(new File(["x"], "i.csv"));
    vi.mocked(api.inspectTutorTutorFile).mockResolvedValue(inspection);
    vi.mocked(api.previewTutorTutorImport).mockResolvedValue(preview);
    vi.mocked(api.applyTutorTutorImport).mockResolvedValue({
      tutors_removed: 4,
      placements_removed: 2,
      tutors_created: 1,
      tags_created: 1,
    });
    const user = userEvent.setup();
    render(<TutorTutorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Import Tutor Roster" }));
    const dialog = await screen.findByRole("dialog", { name: "Import Tutor Roster" });
    await user.click(within(dialog).getByRole("button", { name: "Choose spreadsheet" }));
    expect(await within(dialog).findByLabelText("Tutor Name")).toBeTruthy();
    await user.click(within(dialog).getByRole("button", { name: "Preview Import" }));
    expect(await within(dialog).findByText("Imported Person")).toBeTruthy();
    expect(within(dialog).getByText("Mon 9:00 AM-12:00 PM")).toBeTruthy();
    expect(api.applyTutorTutorImport).not.toHaveBeenCalled();

    await user.click(within(dialog).getByRole("button", { name: "Import Tutors" }));
    const confirm = await screen.findByRole("dialog", { name: "Replace Tutors?" });
    expect(within(confirm).getByText(/Delete all current tutors \(4 currently\)/)).toBeTruthy();
    expect(within(confirm).getByText("Preserve the Tutor Tag registry")).toBeTruthy();
    expect(within(confirm).getByText("Preserve Student records")).toBeTruthy();
    expect(within(confirm).getByText("Clear tutor schedules")).toBeTruthy();
    expect(api.applyTutorTutorImport).not.toHaveBeenCalled();

    await user.click(within(confirm).getByRole("button", { name: "Replace Tutors" }));
    await waitFor(() => expect(api.applyTutorTutorImport).toHaveBeenCalled());
    expect(await screen.findByRole("status")).toBeTruthy();
  });

  it("blocks apply when the preview has errors", async () => {
    vi.mocked(chooseLocalFile).mockResolvedValue(new File(["x"], "i.csv"));
    vi.mocked(api.inspectTutorTutorFile).mockResolvedValue(inspection);
    vi.mocked(api.previewTutorTutorImport).mockResolvedValue({
      ...preview,
      valid_rows: 0,
      invalid_rows: 1,
      can_apply: false,
      rows: [{ ...preview.rows[0], errors: ["Tutor Name is required."] }],
    });
    const user = userEvent.setup();
    render(<TutorTutorRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Import Tutor Roster" }));
    const dialog = await screen.findByRole("dialog", { name: "Import Tutor Roster" });
    await user.click(within(dialog).getByRole("button", { name: "Choose spreadsheet" }));
    await user.click(await within(dialog).findByRole("button", { name: "Preview Import" }));
    const apply = await within(dialog).findByRole("button", { name: "Import Tutors" });
    expect((apply as HTMLButtonElement).disabled).toBe(true);
  });
});
