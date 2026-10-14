import { useEffect, useRef } from "react";
import { formatMinutes } from "../lib/time";
import type { AvailabilityStatus } from "../lib/types";
import {
  cellKey,
  nextStatus,
  slotStarts,
  WEEK_DAYS,
  type WeeklyCells,
  type WeeklyGridRange,
} from "../lib/weeklyAvailability";

const STATUS_CLASS: Record<AvailabilityStatus, string> = {
  Available: "slot-available",
  Tentative: "slot-tentative",
  Unavailable: "slot-unavailable",
};

/**
 * Controlled weekly Availability Window grid shared by modules that edit
 * recurring Available/Tentative/Unavailable windows. It owns only pointer
 * painting; loading, saving, and persistence belong to the caller.
 *
 * Left click cycles the clicked cell and a drag paints that resulting status
 * onto every cell entered. Right click (or right-drag) paints Unavailable.
 */
export function WeeklyAvailabilityEditor({
  cells,
  range,
  onChange,
}: {
  cells: WeeklyCells;
  range: WeeklyGridRange;
  onChange: (cells: WeeklyCells) => void;
}) {
  const cellsRef = useRef(cells);
  cellsRef.current = cells;
  const paintRef = useRef<AvailabilityStatus | null>(null);

  function paint(day: string, slot: number, status: AvailabilityStatus) {
    const key = cellKey(day, slot);
    if (cellsRef.current.get(key) === status) return;
    const updated = new Map(cellsRef.current);
    updated.set(key, status);
    cellsRef.current = updated;
    onChange(updated);
  }

  function endDrag() {
    paintRef.current = null;
  }

  useEffect(() => {
    window.addEventListener("mouseup", endDrag);
    return () => window.removeEventListener("mouseup", endDrag);
  }, []);

  return (
    <div className="availability-grid-wrap">
      <div className="availability-legend" aria-label="Availability legend">
        <span className="legend-chip slot-available">Available</span>
        <span className="legend-chip slot-tentative">Tentative</span>
        <span className="legend-chip slot-unavailable">Unavailable</span>
        <span className="legend-chip slot-unset">Unmarked</span>
        <span className="legend-hint">
          Click to cycle Available, Tentative, and Unavailable; drag to paint; right-click for Unavailable. Unmarked
          cells are saved as Unavailable.
        </span>
      </div>
      <div className="availability-grid" style={{ gridTemplateColumns: `90px repeat(${WEEK_DAYS.length}, 1fr)` }}>
        <div className="avail-header avail-corner" />
        {WEEK_DAYS.map((day) => (
          <div className="avail-header" key={day}>
            {day}
          </div>
        ))}
        {slotStarts(range).map((slot) => (
          <div key={slot} style={{ display: "contents" }}>
            <div className="avail-time-label">{formatMinutes(slot)}</div>
            {WEEK_DAYS.map((day) => {
              const status = cells.get(cellKey(day, slot));
              return (
                <button
                  key={day + slot}
                  type="button"
                  className={`avail-cell ${status ? STATUS_CLASS[status] : "slot-unset"}`}
                  aria-label={`${day} ${formatMinutes(slot)}: ${status ?? "Unmarked"}`}
                  onMouseDown={(event) => {
                    if (event.button === 0) {
                      paintRef.current = nextStatus(cellsRef.current.get(cellKey(day, slot)));
                    } else if (event.button === 2) {
                      paintRef.current = "Unavailable";
                    } else {
                      return;
                    }
                    paint(day, slot, paintRef.current);
                  }}
                  onMouseEnter={() => {
                    if (paintRef.current) paint(day, slot, paintRef.current);
                  }}
                  onContextMenu={(event) => event.preventDefault()}
                  onKeyDown={(event) => {
                    if (event.key === "Enter" || event.key === " ") {
                      event.preventDefault();
                      paint(day, slot, nextStatus(cellsRef.current.get(cellKey(day, slot))));
                    }
                  }}
                />
              );
            })}
          </div>
        ))}
      </div>
    </div>
  );
}
