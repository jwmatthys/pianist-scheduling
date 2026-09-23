import { useEffect, useMemo, useState } from "react";
import { api } from "../lib/api";
import type { AvailabilityStatus, Pianist } from "../lib/types";
import { formatMinutes } from "../lib/time";

const DAYS: string[] = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const START_MIN = 8 * 60; // 8:00 AM
const END_MIN = 20 * 60; // 8:00 PM
const SLOT_MINUTES = 30;

const STATUS_CYCLE: (AvailabilityStatus | "Unset")[] = ["Unset", "Available", "Tentative", "Unavailable"];

const STATUS_CLASS: Record<string, string> = {
  Unset: "slot-unset",
  Available: "slot-available",
  Tentative: "slot-tentative",
  Unavailable: "slot-unavailable",
};

function slotKey(day: string, slot: number) {
  return `${day}|${slot}`;
}

export function AvailabilityGrid({ pianist }: { pianist: Pianist }) {
  const [grid, setGrid] = useState<Map<string, AvailabilityStatus>>(new Map());
  const [dirty, setDirty] = useState(false);
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(true);

  const slots = useMemo(() => {
    const result: number[] = [];
    for (let m = START_MIN; m < END_MIN; m += SLOT_MINUTES) result.push(m);
    return result;
  }, []);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    api.getAvailability(pianist.id).then((data) => {
      if (cancelled) return;
      const next = new Map<string, AvailabilityStatus>();
      for (const slot of data) next.set(slotKey(slot.day, slot.slot_start_minute), slot.status);
      setGrid(next);
      setDirty(false);
      setLoading(false);
    });
    return () => {
      cancelled = true;
    };
  }, [pianist.id]);

  function cycle(day: string, slot: number) {
    const key = slotKey(day, slot);
    const current = grid.get(key) ?? "Unset";
    const idx = STATUS_CYCLE.indexOf(current);
    const next = STATUS_CYCLE[(idx + 1) % STATUS_CYCLE.length];
    const updated = new Map(grid);
    if (next === "Unset") updated.delete(key);
    else updated.set(key, next);
    setGrid(updated);
    setDirty(true);
  }

  async function save() {
    setSaving(true);
    const payload = Array.from(grid.entries()).map(([key, status]) => {
      const [day, slot] = key.split("|");
      return { day, slot_start_minute: Number(slot), status };
    });
    await api.setAvailability(pianist.id, payload);
    setDirty(false);
    setSaving(false);
  }

  if (loading) return <p>Loading availability&hellip;</p>;

  return (
    <div className="availability-grid-wrap">
      <div className="availability-legend">
        <span className="legend-chip slot-available">Available</span>
        <span className="legend-chip slot-tentative">Tentative</span>
        <span className="legend-chip slot-unavailable">Unavailable</span>
        <span className="legend-chip slot-unset">Unset</span>
        <span className="legend-hint">Click a cell to cycle through statuses.</span>
      </div>
      <div className="availability-grid" style={{ gridTemplateColumns: `90px repeat(${DAYS.length}, 1fr)` }}>
        <div className="avail-header avail-corner" />
        {DAYS.map((day) => (
          <div className="avail-header" key={day}>
            {day}
          </div>
        ))}
        {slots.map((slot) => (
          <div className="avail-row-contents" key={slot} style={{ display: "contents" }}>
            <div className="avail-time-label">{formatMinutes(slot)}</div>
            {DAYS.map((day) => {
              const status = grid.get(slotKey(day, slot)) ?? "Unset";
              return (
                <button
                  key={day + slot}
                  type="button"
                  className={`avail-cell ${STATUS_CLASS[status]}`}
                  onClick={() => cycle(day, slot)}
                  title={`${day} ${formatMinutes(slot)}: ${status}`}
                />
              );
            })}
          </div>
        ))}
      </div>
      <div className="availability-actions">
        <button onClick={save} disabled={!dirty || saving} className="primary-btn">
          {saving ? "Saving\u2026" : dirty ? "Save availability" : "Saved"}
        </button>
      </div>
    </div>
  );
}
