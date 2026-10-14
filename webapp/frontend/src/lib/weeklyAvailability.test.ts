import { describe, expect, it } from "vitest";
import { cellKey, cellsToWindows, nextStatus, windowsToCells, type WeeklyGridRange } from "./weeklyAvailability";

const range: WeeklyGridRange = { startMinute: 420, endMinute: 1260, slotMinutes: 30 };

describe("weekly availability conversion", () => {
  it("cycles Unmarked, Available, Tentative, Unavailable, Available", () => {
    expect(nextStatus(undefined)).toBe("Available");
    expect(nextStatus("Available")).toBe("Tentative");
    expect(nextStatus("Tentative")).toBe("Unavailable");
    expect(nextStatus("Unavailable")).toBe("Available");
  });

  it("stores recurring windows by merging adjacent cells, with unmarked cells Unavailable", () => {
    const cells = new Map([
      [cellKey("Monday", 540), "Available" as const],
      [cellKey("Monday", 570), "Available" as const],
      [cellKey("Monday", 600), "Tentative" as const],
    ]);
    const windows = cellsToWindows(cells, range);
    const monday = windows.filter((item) => item.day === "Monday");
    expect(monday).toEqual([
      { day: "Monday", start_minute: 420, end_minute: 540, status: "Unavailable" },
      { day: "Monday", start_minute: 540, end_minute: 600, status: "Available" },
      { day: "Monday", start_minute: 600, end_minute: 630, status: "Tentative" },
      { day: "Monday", start_minute: 630, end_minute: 1260, status: "Unavailable" },
    ]);
    expect(windows.filter((item) => item.day === "Sunday")).toEqual([
      { day: "Sunday", start_minute: 420, end_minute: 1260, status: "Unavailable" },
    ]);
  });

  it("round-trips windows through cells", () => {
    const windows = cellsToWindows(windowsToCells([{ day: "Friday", start_minute: 600, end_minute: 720, status: "Available" }], range), range);
    expect(windows.filter((item) => item.day === "Friday")).toEqual([
      { day: "Friday", start_minute: 420, end_minute: 600, status: "Unavailable" },
      { day: "Friday", start_minute: 600, end_minute: 720, status: "Available" },
      { day: "Friday", start_minute: 720, end_minute: 1260, status: "Unavailable" },
    ]);
  });
});
