import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from "react";
import { api } from "../lib/api";
import type { Pianist } from "../lib/types";
import { AvailabilityGrid, type AvailabilityGridHandle } from "../components/AvailabilityGrid";
import { AvailabilityImportDialog } from "../components/AvailabilityImportDialog";

export type PianistsPageHandle = {
  saveAvailabilityBeforeLeaving: () => Promise<void>;
};

export const PianistsPage = forwardRef<PianistsPageHandle>(function PianistsPage(_, ref) {
  const [pianists, setPianists] = useState<Pianist[]>([]);
  const [selectedId, setSelectedId] = useState<number | null>(null);
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [maxHours, setMaxHours] = useState("");
  const [profileMessage, setProfileMessage] = useState<string | null>(null);
  const [profileError, setProfileError] = useState<string | null>(null);
  const availabilityGridRef = useRef<AvailabilityGridHandle>(null);

  useImperativeHandle(ref, () => ({
    saveAvailabilityBeforeLeaving: () => availabilityGridRef.current?.saveBeforeLeaving() ?? Promise.resolve(),
  }));

  async function refresh() {
    const list = await api.listPianists();
    setPianists(list);
    if (list.length && selectedId === null) setSelectedId(list[0].id);
  }

  useEffect(() => { void refresh(); }, []);

  async function addPianist(e: React.FormEvent) {
    e.preventDefault();
    if (!name.trim()) return;
    setProfileMessage(null);
    setProfileError(null);
    try {
      const created = await api.createPianist({
        name: name.trim(),
        email: email.trim(),
        max_hours_per_week: maxHours ? Number(maxHours) : null,
      });
      setName("");
      setEmail("");
      setMaxHours("");
      await refresh();
      setSelectedId(created.id);
    } catch (error) {
      setProfileError(error instanceof Error ? error.message : "Could not add Pianist.");
    }
  }

  async function updateSelectedPianist(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!selected) return;
    const form = new FormData(event.currentTarget);
    setProfileMessage(null);
    setProfileError(null);
    try {
      await api.updatePianist(selected.id, {
        name: String(form.get("name") ?? "").trim(),
        email: String(form.get("email") ?? "").trim(),
        max_hours_per_week: String(form.get("max_hours_per_week") ?? "").trim()
          ? Number(form.get("max_hours_per_week"))
          : null,
      });
      await refresh();
      setProfileMessage("Pianist details saved.");
    } catch (error) {
      setProfileError(error instanceof Error ? error.message : "Could not update Pianist.");
    }
  }

  async function removePianist(p: Pianist) {
    if (!confirm(`Remove ${p.name}? This also clears their lesson assignments.`)) return;
    await api.deletePianist(p.id);
    if (selectedId === p.id) setSelectedId(null);
    refresh();
  }

  async function selectPianist(id: number) {
    if (id === selectedId) return;
    await availabilityGridRef.current?.saveBeforeLeaving();
    setSelectedId(id);
  }

  const selected = pianists.find((p) => p.id === selectedId) ?? null;

  async function prepareAvailabilityImport() {
    return await availabilityGridRef.current?.prepareForImport() ?? true;
  }

  async function refreshAvailabilityAfterImport() {
    await availabilityGridRef.current?.reloadFromServer();
    refresh();
  }

  return (
    <div className="page pianists-page">
      <div className="pianists-sidebar">
        <h2>Pianists</h2>
        <form className="add-pianist-form" onSubmit={addPianist}>
          <input placeholder="Name" value={name} onChange={(e) => setName(e.target.value)} required />
          <input placeholder="Email" value={email} onChange={(e) => setEmail(e.target.value)} />
          <input
            placeholder="Max hrs/wk"
            type="number"
            min="0"
            step="0.5"
            value={maxHours}
            onChange={(e) => setMaxHours(e.target.value)}
          />
          <button type="submit" className="primary-btn">
            Add pianist
          </button>
        </form>
        <ul className="pianist-list">
          {pianists.map((p) => (
            <li key={p.id} className={p.id === selectedId ? "selected" : ""}>
              <button className="pianist-list-item" onClick={() => selectPianist(p.id)}>
                <strong>{p.name}</strong>
              </button>
              <button className="danger-btn small" onClick={() => removePianist(p)}>
                &times;
              </button>
            </li>
          ))}
          {pianists.length === 0 && <p className="muted">No pianists yet -- add one above.</p>}
        </ul>
      </div>
      <div className="pianists-detail">
        <div className="availability-heading-row">
        {selected ? (
          <>
            <h3>
              {selected.name}'s weekly availability
              {selected.max_hours_per_week ? ` \u2014 ${selected.max_hours_per_week}h cap` : ""}
            </h3>
          </>
        ) : (
          <h3>Pianist availability</h3>
        )}
          <AvailabilityImportDialog
            onBeforeApply={prepareAvailabilityImport}
            onApplied={refreshAvailabilityAfterImport}
          />
        </div>
        {profileMessage && <div className="availability-success" role="status">{profileMessage}</div>}
        {profileError && <div className="error-banner" role="alert">{profileError}</div>}
        {selected && (
          <form key={selected.id} className="pianist-profile-form" onSubmit={(event) => void updateSelectedPianist(event)}>
            <label>Name<input name="name" defaultValue={selected.name} required /></label>
            <label>Email<input name="email" type="email" defaultValue={selected.email} /></label>
            <label>Max Hours Per Week<input name="max_hours_per_week" type="number" min="0" step="0.5" defaultValue={selected.max_hours_per_week ?? ""} /></label>
            <button type="submit" className="secondary-btn">Save Pianist details</button>
          </form>
        )}
        {selected
          ? <AvailabilityGrid ref={availabilityGridRef} pianist={selected} />
          : <p className="muted">Select a pianist to edit availability manually.</p>}
      </div>
    </div>
  );
});
