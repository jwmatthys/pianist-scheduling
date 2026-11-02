import type { TutorScheduleEntry } from "./types";

type PlacementFields = Pick<TutorScheduleEntry, "placement_type" | "day" | "start_minute" | "end_minute">;

export type TutorPlacementState = "assigned" | "unassigned" | "manuallyUnplaced";

/**
 * Canonical placement state. A student is placed only when the entry carries an actual tutor time;
 * automatically unplaced students are stored with placement_type "automatic" but no day or time.
 */
export function tutorPlacementState(entry: PlacementFields): TutorPlacementState {
  if (entry.placement_type === "manuallyUnplaced") return "manuallyUnplaced";
  const placed = entry.day !== null && entry.start_minute !== null && entry.end_minute !== null;
  return placed ? "assigned" : "unassigned";
}

export function isTutorEntryPlaced(entry: PlacementFields): boolean {
  return tutorPlacementState(entry) === "assigned";
}
