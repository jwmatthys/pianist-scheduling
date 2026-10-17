import { useEffect, useMemo, useRef, useState } from "react";
import { WeeklyAvailabilityEditor } from "../../components/WeeklyAvailabilityEditor";
import type { RecurringAvailabilityWindow } from "../../lib/types";
import {
  cellsToWindows,
  windowsOutsideRange,
  windowsToCells,
  type WeeklyCells,
  type WeeklyGridRange,
} from "../../lib/weeklyAvailability";
import { Modal } from "../clinical/ClinicalDialogs";

// 7:00 AM through 9:00 PM in 30-minute cells, matching the Clinical availability editor.
const LESSON_RANGE: WeeklyGridRange = { startMinute: 7 * 60, endMinute: 21 * 60, slotMinutes: 30 };

/**
 * Recurring weekly availability dialog shared by the Lesson Instructor and Student rosters.
 * Persistence is supplied by the caller so each roster keeps its own module-owned endpoints.
 */
export function LessonAvailabilityDialog({
  title,
  ownerId,
  load,
  save,
  onSaved,
  onClose,
}: {
  title: string;
  ownerId: string;
  load: () => Promise<{ windows: RecurringAvailabilityWindow[] }>;
  save: (windows: RecurringAvailabilityWindow[]) => Promise<unknown>;
  onSaved: () => Promise<void>;
  onClose: () => void;
}) {
  const [cells, setCells] = useState<WeeklyCells | null>(null);
  const [dirty, setDirty] = useState(false);
  const [cleared, setCleared] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const preserved = useRef<RecurringAvailabilityWindow[]>([]);
  const range = useMemo(() => LESSON_RANGE, []);
  const loadRef = useRef(load);
  loadRef.current = load;

  useEffect(() => {
    let cancelled = false;
    loadRef
      .current()
      .then(({ windows }) => {
        if (cancelled) return;
        preserved.current = windowsOutsideRange(windows, range);
        setCells(windowsToCells(windows, range));
      })
      .catch((caught) => {
        if (!cancelled) setError(caught instanceof Error ? caught.message : String(caught));
      });
    return () => {
      cancelled = true;
    };
  }, [ownerId, range]);

  function requestClose() {
    if (dirty && !confirm("Discard unsaved availability changes?")) return;
    onClose();
  }

  async function submit() {
    if (!cells) return;
    setSaving(true);
    try {
      // A cleared grid removes all availability rather than saving a week of Unavailable.
      await save(cleared ? [] : [...cellsToWindows(cells, range), ...preserved.current]);
      await onSaved();
      onClose();
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : String(caught));
      setSaving(false);
    }
  }

  return (
    <Modal title={title} onClose={requestClose} className="clinical-availability-dialog">
      {error && (
        <div className="error-banner" role="alert">
          {error}
        </div>
      )}
      {cells ? (
        <WeeklyAvailabilityEditor
          cells={cells}
          range={range}
          onChange={(next) => {
            setCells(next);
            setDirty(true);
            setCleared(false);
          }}
        />
      ) : (
        !error && <p>Loading availability&hellip;</p>
      )}
      <div className="dialog-actions clinical-availability-actions">
        <button
          type="button"
          className="danger-btn"
          disabled={!cells}
          onClick={() => {
            setCells(new Map());
            setDirty(true);
            setCleared(true);
          }}
        >
          Clear Availability
        </button>
        <button type="button" onClick={requestClose}>
          Cancel
        </button>
        <button type="button" className="primary-btn" disabled={!cells || saving} onClick={() => void submit()}>
          Save Availability
        </button>
      </div>
    </Modal>
  );
}
