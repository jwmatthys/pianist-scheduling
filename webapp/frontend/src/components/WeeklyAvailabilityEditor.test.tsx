import { fireEvent, render, screen } from "@testing-library/react";
import { useState } from "react";
import { describe, expect, it } from "vitest";
import type { WeeklyCells, WeeklyGridRange } from "../lib/weeklyAvailability";
import { WeeklyAvailabilityEditor } from "./WeeklyAvailabilityEditor";

const range: WeeklyGridRange = { startMinute: 420, endMinute: 1260, slotMinutes: 30 };

function Harness({ initial = new Map() }: { initial?: WeeklyCells }) {
  const [cells, setCells] = useState<WeeklyCells>(initial);
  return <WeeklyAvailabilityEditor cells={cells} range={range} onChange={setCells} />;
}

const cell = (day: string, time: string) => screen.getByRole("button", { name: new RegExp(`^${day} ${time}:`) });

describe("WeeklyAvailabilityEditor", () => {
  it("shows Monday through Sunday, 7:00 AM through 9:00 PM, and a legend", () => {
    render(<Harness />);
    for (const day of ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]) {
      expect(screen.getByText(day)).toBeTruthy();
    }
    expect(screen.getAllByRole("button")).toHaveLength(7 * 28);
    expect(cell("Monday", "7:00 AM")).toBeTruthy();
    expect(cell("Sunday", "8:30 PM")).toBeTruthy();
    expect(screen.queryByRole("button", { name: /9:00 PM/ })).toBeNull();
    const legend = screen.getByLabelText("Availability legend");
    for (const label of ["Available", "Tentative", "Unavailable", "Unmarked"]) {
      expect(legend.textContent).toContain(label);
    }
  });

  it("cycles Unmarked, Available, Tentative, Unavailable, Available on left click", () => {
    render(<Harness />);
    const label = (status: string) => `Monday 9:00 AM: ${status}`;
    for (const expected of ["Available", "Tentative", "Unavailable", "Available"]) {
      fireEvent.mouseDown(cell("Monday", "9:00 AM"), { button: 0 });
      fireEvent.mouseUp(window);
      expect(screen.getByRole("button", { name: label(expected) })).toBeTruthy();
    }
  });

  it("sets Unavailable on right click", () => {
    render(<Harness />);
    fireEvent.mouseDown(cell("Tuesday", "10:00 AM"), { button: 2 });
    fireEvent.mouseUp(window);
    expect(screen.getByRole("button", { name: "Tuesday 10:00 AM: Unavailable" })).toBeTruthy();
  });

  it("paints a single state across a drag instead of cycling each cell", () => {
    render(<Harness initial={new Map([["Monday|600", "Tentative" as const]])} />);
    fireEvent.mouseDown(cell("Monday", "9:00 AM"), { button: 0 });
    fireEvent.mouseEnter(cell("Monday", "9:30 AM"));
    fireEvent.mouseEnter(cell("Monday", "10:00 AM"));
    fireEvent.mouseUp(window);
    for (const time of ["9:00 AM", "9:30 AM", "10:00 AM"]) {
      expect(screen.getByRole("button", { name: `Monday ${time}: Available` })).toBeTruthy();
    }
    fireEvent.mouseEnter(cell("Monday", "10:30 AM"));
    expect(screen.getByRole("button", { name: "Monday 10:30 AM: Unmarked" })).toBeTruthy();
  });
});
