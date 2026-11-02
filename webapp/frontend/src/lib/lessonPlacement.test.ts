import { describe, expect, it } from "vitest";
import { isLessonEntryPlaced, lessonPlacementState } from "./lessonPlacement";

const base = { placement_type: "automatic" as const, day: "Monday", start_minute: 540, end_minute: 600 };

describe("lessonPlacementState", () => {
  it("is assigned only when the entry has a day and times", () => {
    expect(lessonPlacementState(base)).toBe("assigned");
    expect(lessonPlacementState({ ...base, placement_type: "manual" })).toBe("assigned");
    expect(isLessonEntryPlaced(base)).toBe(true);
  });

  it("treats an automatic entry without a time as unassigned", () => {
    for (const missing of [{ day: null }, { start_minute: null }, { end_minute: null }]) {
      expect(lessonPlacementState({ ...base, ...missing })).toBe("unassigned");
      expect(isLessonEntryPlaced({ ...base, ...missing })).toBe(false);
    }
    expect(lessonPlacementState({ ...base, placement_type: null, day: null, start_minute: null, end_minute: null }))
      .toBe("unassigned");
  });

  it("reports manually unplaced regardless of times", () => {
    expect(lessonPlacementState({ ...base, placement_type: "manuallyUnplaced" })).toBe("manuallyUnplaced");
    expect(isLessonEntryPlaced({ ...base, placement_type: "manuallyUnplaced" })).toBe(false);
  });
});
