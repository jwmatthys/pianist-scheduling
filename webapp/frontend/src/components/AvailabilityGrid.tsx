import { forwardRef, useEffect, useImperativeHandle, useMemo, useRef, useState } from "react";
import { api } from "../lib/api";
import type { AvailabilityStatus, Pianist } from "../lib/types";
import { formatMinutes } from "../lib/time";

const DAYS: string[] = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const START_MIN = 6 * 60; // 6:00 AM
const END_MIN = 22 * 60; // 10:00 PM
const SLOT_MINUTES = 30;

const STATUS_CYCLE: AvailabilityStatus[] = ["Unavailable", "Available", "Tentative"];

const STATUS_CLASS: Record<string, string> = {
  Available: "slot-available",
  Tentative: "slot-tentative",
  Unavailable: "slot-unavailable",
};

function slotKey(day: string, slot: number) {
  return `${day}|${slot}`;
}

export type AvailabilityGridHandle = {
  saveBeforeLeaving: () => Promise<void>;
};

export const AvailabilityGrid = forwardRef<AvailabilityGridHandle, { pianist: Pianist }>(function AvailabilityGrid(
  { pianist },
  ref
) {
  const [grid, setGrid] = useState<Map<string, AvailabilityStatus>>(new Map());
  const [dirty, setDirty] = useState(false);
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(true);
  const gridRef = useRef<Map<string, AvailabilityStatus>>(new Map());
  const dirtyRef = useRef(false);
  const dragActiveRef = useRef(false);
  const changedKeysRef = useRef<Set<string>>(new Set());

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
      gridRef.current = next;
      setGrid(next);
      setDirty(false);
      dirtyRef.current = false;
      setLoading(false);
    });
    return () => {
      cancelled = true;
    };
  }, [pianist.id]);

  function cycle(day: string, slot: number) {
    const key = slotKey(day, slot);
    const current = gridRef.current.get(key) ?? "Unavailable";
    const idx = STATUS_CYCLE.indexOf(current);
    const next = STATUS_CYCLE[(idx + 1) % STATUS_CYCLE.length];
    const updated = new Map(gridRef.current);
    updated.set(key, next);
    gridRef.current = updated;
    setGrid(updated);
    setDirty(true);
    dirtyRef.current = true;
  }

  function setUnavailable(day: string, slot: number) {
    const key = slotKey(day, slot);
    if ((gridRef.current.get(key) ?? "Unavailable") === "Unavailable") return;
    const updated = new Map(gridRef.current);
    updated.set(key, "Unavailable");
    gridRef.current = updated;
    setGrid(updated);
    setDirty(true);
    dirtyRef.current = true;
  }

  async function save() {
    setSaving(true);
    const payload = Array.from(gridRef.current.entries()).map(([key, status]) => {
      const [day, slot] = key.split("|");
      return { day, slot_start_minute: Number(slot), status };
    });
    await api.setAvailability(pianist.id, payload);
    setDirty(false);
    dirtyRef.current = false;
    setSaving(false);
  }

  useImperativeHandle(ref, () => ({
    async saveBeforeLeaving() {
      if (!dirtyRef.current) return;
      if (confirm("Save availability changes before leaving this section?")) await save();
    },
  }), []);

  function beginDrag(day: string, slot: number) {
    dragActiveRef.current = true;
    changedKeysRef.current = new Set([slotKey(day, slot)]);
    cycle(day, slot);
  }

  function extendDrag(day: string, slot: number) {
    if (!dragActiveRef.current) return;
    const key = slotKey(day, slot);
    if (changedKeysRef.current.has(key)) return;
    changedKeysRef.current.add(key);
    cycle(day, slot);
  }

  function endDrag() {
    dragActiveRef.current = false;
    changedKeysRef.current.clear();
  }

  useEffect(() => {
    window.addEventListener("pointerup", endDrag);
    return () => window.removeEventListener("pointerup", endDrag);
  }, []);

  if (loading) return <p>Loading availability&hellip;</p>;

  return (
    <div className="availability-grid-wrap">
      <div className="availability-legend">
        <span className="legend-chip slot-available">Available</span>
        <span className="legend-chip slot-tentative">Tentative</span>
        <span className="legend-chip slot-unavailable">Unavailable</span>
        <span className="legend-hint">Click or drag to cycle statuses. Right-click to set Unavailable.</span>
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
              const status = grid.get(slotKey(day, slot)) ?? "Unavailable";
              return (
                <button
                  key={day + slot}
                  type="button"
                  className={`avail-cell ${STATUS_CLASS[status]}`}
                  onPointerDown={(event) => {
                    if (event.button === 0) beginDrag(day, slot);
                  }}
                  onPointerEnter={() => extendDrag(day, slot)}
                  onPointerUp={endDrag}
                  onPointerCancel={endDrag}
                  onContextMenu={(event) => {
                    event.preventDefault();
                    setUnavailable(day, slot);
                  }}
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
});
