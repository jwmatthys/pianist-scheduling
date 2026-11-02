import type { LessonScheduleEntry } from "./types";

type PlacementFields = Pick<LessonScheduleEntry, "placement_type" | "day" | "start_minute" | "end_minute">;

export type LessonPlacementState = "assigned" | "unassigned" | "manuallyUnplaced";

/**
 * Canonical placement state. A student is placed only when the entry carries an actual lesson time;
 * automatically unplaced students are stored with placement_type "automatic" but no day or time.
 */
export function lessonPlacementState(entry: PlacementFields): LessonPlacementState {
  if (entry.placement_type === "manuallyUnplaced") return "manuallyUnplaced";
  const placed = entry.day !== null && entry.start_minute !== null && entry.end_minute !== null;
  return placed ? "assigned" : "unassigned";
}

export function isLessonEntryPlaced(entry: PlacementFields): boolean {
  return lessonPlacementState(entry) === "assigned";
}
