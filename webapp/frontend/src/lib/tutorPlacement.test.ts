import { describe, expect, it } from "vitest";
import { isTutorEntryPlaced, tutorPlacementState } from "./tutorPlacement";

const base = { placement_type: "automatic" as const, day: "Monday", start_minute: 540, end_minute: 600 };

describe("tutorPlacementState", () => {
  it("is assigned only when the entry has a day and times", () => {
    expect(tutorPlacementState(base)).toBe("assigned");
    expect(tutorPlacementState({ ...base, placement_type: "manual" })).toBe("assigned");
    expect(isTutorEntryPlaced(base)).toBe(true);
  });

  it("treats an automatic entry without a time as unassigned", () => {
    for (const missing of [{ day: null }, { start_minute: null }, { end_minute: null }]) {
      expect(tutorPlacementState({ ...base, ...missing })).toBe("unassigned");
      expect(isTutorEntryPlaced({ ...base, ...missing })).toBe(false);
    }
    expect(tutorPlacementState({ ...base, placement_type: null, day: null, start_minute: null, end_minute: null }))
      .toBe("unassigned");
  });

  it("reports manually unplaced regardless of times", () => {
    expect(tutorPlacementState({ ...base, placement_type: "manuallyUnplaced" })).toBe("manuallyUnplaced");
    expect(isTutorEntryPlaced({ ...base, placement_type: "manuallyUnplaced" })).toBe(false);
  });
});
