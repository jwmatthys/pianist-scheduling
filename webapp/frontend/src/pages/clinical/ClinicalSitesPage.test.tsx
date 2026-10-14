import { render, screen, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../../lib/api";
import type { ClinicalTag, PlacementOpportunity } from "../../lib/types";
import { ClinicalSitesPage } from "./ClinicalSitesPage";

const store = vi.hoisted(() => ({
  opportunities: [] as PlacementOpportunity[],
  tags: [] as { id: string; name: string; student_count: number }[],
  next: 1,
}));

function tagOut(tag: { id: string; name: string; student_count: number }): ClinicalTag {
  const count = store.opportunities.filter((o) => o.tags.some((t) => t.id === tag.id)).length;
  return { ...tag, opportunity_count: count, is_unused: count === 0 };
}

vi.mock("../../lib/api", () => {
  const find = (id: string) => store.opportunities.find((o) => o.id === id)!;
  return {
    api: {
      listClinicalOpportunities: vi.fn(async () => store.opportunities.map((o) => ({ ...o }))),
      listClinicalTags: vi.fn(async () => store.tags.map(tagOut)),
      createClinicalOpportunity: vi.fn(async (data: any) => {
        const { tag_ids, ...fields } = data;
        const item = {
          ...fields,
          id: `o${store.next++}`,
          placement_count: 0,
          tags: store.tags.filter((t) => tag_ids.includes(t.id)).map((t) => ({ id: t.id, name: t.name })),
        };
        store.opportunities.push(item);
        return item;
      }),
      updateClinicalOpportunity: vi.fn(async (id: string, data: any) => {
        Object.assign(find(id), data);
        return find(id);
      }),
      setClinicalOpportunityTags: vi.fn(async (id: string, tagIds: string[]) => {
        find(id).tags = store.tags.filter((t) => tagIds.includes(t.id)).map((t) => ({ id: t.id, name: t.name }));
        return find(id);
      }),
      clearClinicalOpportunities: vi.fn(async () => {
        const count = store.opportunities.length;
        store.opportunities = [];
        return { removed_count: count, removed_placement_count: 0 };
      }),
      deleteClinicalOpportunity: vi.fn(async (id: string) => {
        store.opportunities = store.opportunities.filter((o) => o.id !== id);
        return { removed_placement_count: 0 };
      }),
      createClinicalTag: vi.fn(async (name: string) => {
        const tag = { id: `new-tag-${store.next++}`, name: name.trim().toUpperCase(), student_count: 0 };
        store.tags.push(tag);
        return tagOut(tag);
      }),
      renameClinicalTag: vi.fn(async (id: string, name: string) => {
        const tag = store.tags.find((t) => t.id === id)!;
        tag.name = name.trim().toUpperCase();
        return tagOut(tag);
      }),
      deleteClinicalTag: vi.fn(async (id: string) => {
        store.tags = store.tags.filter((t) => t.id !== id);
        return { affected_student_count: 0 };
      }),
    },
  };
});

function opportunity(overrides: Partial<PlacementOpportunity> = {}): PlacementOpportunity {
  return {
    id: "o0",
    site_name: "Synthetic Clinic",
    site_address: "",
    contact_name: "",
    contact_email: "",
    contact_phone: "",
    day_of_week: "Tuesday",
    start_minute: 8 * 60,
    end_minute: 17 * 60,
    session_length_minutes: 60,
    travel_minutes: 15,
    capacity: 1,
    tags: [],
    placement_count: 0,
    ...overrides,
  };
}

async function fillOpportunity(user: ReturnType<typeof userEvent.setup>, dialog: HTMLElement) {
  await user.type(within(dialog).getByLabelText("Site Name"), "North Clinic");
  await user.selectOptions(within(dialog).getByLabelText("Day of Week"), "Monday");
  await user.type(within(dialog).getByLabelText("Start Time"), "9:30 AM");
  await user.type(within(dialog).getByLabelText("End Time"), "1:00 PM");
  await user.type(within(dialog).getByLabelText(/Session Length/), "45");
}

describe("ClinicalSitesPage", () => {
  beforeEach(() => {
    store.opportunities = [];
    store.tags = [];
    store.next = 1;
  });

  it("adds a Placement Opportunity and shows 12-hour times with default capacity", async () => {
    const user = userEvent.setup();
    render(<ClinicalSitesPage />);
    await screen.findByText(/No Placement Opportunities yet/);
    await user.click(screen.getByRole("button", { name: "Add Placement Opportunity" }));
    const dialog = screen.getByRole("dialog", { hidden: true });
    expect((within(dialog).getByLabelText("Capacity") as HTMLInputElement).value).toBe("1");
    await fillOpportunity(user, dialog);
    await user.click(within(dialog).getByRole("button", { name: "Add Opportunity" }));

    const row = (await screen.findByText("North Clinic")).closest("tr")!;
    expect(within(row).getByText("9:30 AM")).toBeTruthy();
    expect(within(row).getByText("1:00 PM")).toBeTruthy();
    expect(within(row).getByText("45 min")).toBeTruthy();
  });

  it("shows validation errors without saving", async () => {
    const user = userEvent.setup();
    render(<ClinicalSitesPage />);
    await screen.findByText(/No Placement Opportunities yet/);
    await user.click(screen.getByRole("button", { name: "Add Placement Opportunity" }));
    const dialog = screen.getByRole("dialog", { hidden: true });
    await user.type(within(dialog).getByLabelText("Site Name"), "X");
    await user.selectOptions(within(dialog).getByLabelText("Day of Week"), "Monday");
    await user.type(within(dialog).getByLabelText("Start Time"), "5:00 PM");
    await user.type(within(dialog).getByLabelText("End Time"), "9:00 AM");
    await user.type(within(dialog).getByLabelText(/Session Length/), "0");
    await user.click(within(dialog).getByRole("button", { name: "Add Opportunity" }));
    expect(within(dialog).getByText("End Time must be after Start Time.")).toBeTruthy();
    expect(within(dialog).getByText("Session Length must be greater than 0.")).toBeTruthy();
    const { api } = await import("../../lib/api");
    expect(api.createClinicalOpportunity).not.toHaveBeenCalled();
  });

  it("edits and deletes a Placement Opportunity", async () => {
    store.opportunities = [opportunity()];
    const user = userEvent.setup();
    render(<ClinicalSitesPage />);
    await user.click(await screen.findByRole("button", { name: "Edit Synthetic Clinic" }));
    const dialog = screen.getByRole("dialog", { hidden: true });
    const name = within(dialog).getByLabelText("Site Name");
    await user.clear(name);
    await user.type(name, "Renamed Clinic");
    await user.click(within(dialog).getByRole("button", { name: "Save Changes" }));
    await screen.findByText("Renamed Clinic");

    await user.click(screen.getByRole("button", { name: "Delete Renamed Clinic" }));
    const confirm = screen.getByRole("dialog", { hidden: true });
    await user.click(within(confirm).getByRole("button", { name: "Delete Placement Opportunity" }));
    await screen.findByText(/No Placement Opportunities yet/);
  });

  it("deletes all Clinical Sites only after confirmation", async () => {
    store.opportunities = [opportunity()];
    const user = userEvent.setup();
    render(<ClinicalSitesPage />);
    await user.click(await screen.findByRole("button", { name: "Delete all Clinical Sites" }));
    const cancel = screen.getByRole("dialog", { hidden: true });
    await user.click(within(cancel).getByRole("button", { name: "Cancel" }));
    expect(api.clearClinicalOpportunities).not.toHaveBeenCalled();
    await user.click(screen.getByRole("button", { name: "Delete all Clinical Sites" }));
    const confirm = screen.getByRole("dialog", { hidden: true });
    await user.click(within(confirm).getByRole("button", { name: "Delete all Clinical Sites" }));
    await screen.findByText(/No Placement Opportunities yet/);
    expect(api.clearClinicalOpportunities).toHaveBeenCalledTimes(1);
  });

  it("states the placement count when deleting an in-use opportunity", async () => {
    store.opportunities = [opportunity({ placement_count: 2 })];
    const user = userEvent.setup();
    render(<ClinicalSitesPage />);
    await user.click(await screen.findByRole("button", { name: "Delete Synthetic Clinic" }));
    expect(screen.getByText(/2 placements are assigned/)).toBeTruthy();
  });

  it("persists tag assignment immediately when editing, keeping the tag", async () => {
    store.tags = [{ id: "t1", name: "KIDS", student_count: 0 }];
    store.opportunities = [opportunity()];
    const user = userEvent.setup();
    render(<ClinicalSitesPage />);
    await user.click(await screen.findByRole("button", { name: "Edit Synthetic Clinic" }));
    const dialog = screen.getByRole("dialog", { hidden: true });
    const box = within(dialog).getByRole("checkbox") as HTMLInputElement;
    await user.click(box);
    expect(store.opportunities[0].tags.map((t) => t.name)).toEqual(["KIDS"]);
    expect(box.checked).toBe(true);
    await user.click(box);
    expect(store.opportunities[0].tags).toEqual([]);
    expect(store.tags).toHaveLength(1);
  });

  it("manages tags: add, rename, grey unused, delete", async () => {
    store.tags = [{ id: "t1", name: "USED", student_count: 0 }];
    store.opportunities = [opportunity({ tags: [{ id: "t1", name: "USED" }] })];
    const user = userEvent.setup();
    render(<ClinicalSitesPage />);
    await user.click(await screen.findByRole("button", { name: "Manage Tags" }));
    const dialog = screen.getByRole("dialog", { name: "Manage Tags", hidden: true });

    await user.type(within(dialog).getByLabelText("New tag"), "  mental health ");
    await user.click(within(dialog).getByRole("button", { name: "Add Tag" }));
    const added = await within(dialog).findByText("MENTAL HEALTH");
    expect(added.className).toContain("clinical-tag-unused");
    expect(within(dialog).getByText("USED").className).not.toContain("clinical-tag-unused");

    await user.click(within(dialog).getByRole("button", { name: "Rename MENTAL HEALTH" }));
    const input = within(dialog).getByLabelText("New name for MENTAL HEALTH");
    await user.clear(input);
    await user.type(input, "trauma");
    await user.click(within(dialog).getByRole("button", { name: "Save" }));
    await within(dialog).findByText("TRAUMA");

    await user.click(within(dialog).getByRole("button", { name: "Delete TRAUMA" }));
    const confirm = screen.getByRole("dialog", { name: 'Delete tag "TRAUMA"?', hidden: true });
    await user.click(within(confirm).getByRole("button", { name: "Delete Tag" }));
    expect(store.tags.map((t) => t.name)).toEqual(["USED"]);
  });

  it("warns with the student count before deleting a tag used by students", async () => {
    store.tags = [{ id: "t1", name: "MENTAL HEALTH", student_count: 7 }];
    const user = userEvent.setup();
    render(<ClinicalSitesPage />);
    await user.click(await screen.findByRole("button", { name: "Manage Tags" }));
    await user.click(await screen.findByRole("button", { name: "Delete MENTAL HEALTH" }));
    const confirm = screen.getByRole("dialog", { name: 'Delete tag "MENTAL HEALTH"?', hidden: true });
    expect(
      within(confirm).getByText(
        "This tag is currently used by 7 students. Deleting it will remove all associated student preference, requirement, and restriction settings.",
      ),
    ).toBeTruthy();
    await user.click(within(confirm).getByRole("button", { name: "Cancel" }));
    expect(store.tags).toHaveLength(1);
  });
});
