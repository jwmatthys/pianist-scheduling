import { describe, expect, it } from "vitest";
import { formatHours } from "./availabilityHours";

describe("formatHours", () => {
  it.each([
    [0, "0 hrs"],
    [60, "1 hr"],
    [90, "1.5 hrs"],
    [720, "12 hrs"],
    [45, "0.75 hrs"],
  ])("%i minutes -> %s", (minutes, expected) => {
    expect(formatHours(minutes)).toBe(expected);
  });
});
