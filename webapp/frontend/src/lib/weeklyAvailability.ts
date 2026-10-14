import type { AvailabilityStatus, RecurringAvailabilityWindow } from "./types";

export const WEEK_DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

export interface WeeklyGridRange {
  startMinute: number;
  endMinute: number;
  slotMinutes: number;
}

/** Cells are an editing aid only; they are never persisted. A missing key is "unmarked". */
export type WeeklyCells = Map<string, AvailabilityStatus>;

export const cellKey = (day: string, slotStart: number) => `${day}|${slotStart}`;

export function slotStarts(range: WeeklyGridRange): number[] {
  const slots: number[] = [];
  for (let minute = range.startMinute; minute + range.slotMinutes <= range.endMinute; minute += range.slotMinutes) {
    slots.push(minute);
  }
  return slots;
}

/** Left-click order: Unmarked, Available, Tentative, Unavailable, then back to Available. */
export function nextStatus(current: AvailabilityStatus | undefined): AvailabilityStatus {
  if (current === "Available") return "Tentative";
  if (current === "Tentative") return "Unavailable";
  return "Available";
}

export function windowsToCells(windows: RecurringAvailabilityWindow[], range: WeeklyGridRange): WeeklyCells {
  const cells: WeeklyCells = new Map();
  for (const window of windows) {
    for (const slot of slotStarts(range)) {
      if (slot >= window.start_minute && slot + range.slotMinutes <= window.end_minute) {
        cells.set(cellKey(window.day, slot), window.status);
      }
    }
  }
  return cells;
}

/** Merge adjacent same-status cells into recurring windows; unmarked cells become Unavailable. */
export function cellsToWindows(cells: WeeklyCells, range: WeeklyGridRange): RecurringAvailabilityWindow[] {
  const windows: RecurringAvailabilityWindow[] = [];
  for (const day of WEEK_DAYS) {
    let open: RecurringAvailabilityWindow | null = null;
    for (const slot of slotStarts(range)) {
      const status = cells.get(cellKey(day, slot)) ?? "Unavailable";
      if (open && open.status === status && open.end_minute === slot) {
        open.end_minute = slot + range.slotMinutes;
      } else {
        open = { day, start_minute: slot, end_minute: slot + range.slotMinutes, status };
        windows.push(open);
      }
    }
  }
  return windows;
}

/** Windows lying wholly outside the grid cannot be shown, so callers should carry them through a save. */
export function windowsOutsideRange(
  windows: RecurringAvailabilityWindow[],
  range: WeeklyGridRange
): RecurringAvailabilityWindow[] {
  return windows.filter((item) => item.end_minute <= range.startMinute || item.start_minute >= range.endMinute);
}
