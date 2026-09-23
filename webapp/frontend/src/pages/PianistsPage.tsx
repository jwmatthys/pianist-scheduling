import { useEffect, useState } from "react";
import { api } from "../lib/api";
import type { Pianist } from "../lib/types";
import { AvailabilityGrid } from "../components/AvailabilityGrid";

export function PianistsPage() {
  const [pianists, setPianists] = useState<Pianist[]>([]);
  const [selectedId, setSelectedId] = useState<number | null>(null);
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [maxHours, setMaxHours] = useState("");

  function refresh() {
    api.listPianists().then((list) => {
      setPianists(list);
      if (list.length && selectedId === null) setSelectedId(list[0].id);
    });
  }

  useEffect(refresh, []);

  async function addPianist(e: React.FormEvent) {
    e.preventDefault();
    if (!name.trim()) return;
    const created = await api.createPianist({
      name: name.trim(),
      email: email.trim(),
      max_hours_per_week: maxHours ? Number(maxHours) : null,
    });
    setName("");
    setEmail("");
    setMaxHours("");
    refresh();
    setSelectedId(created.id);
  }

  async function updateCap(p: Pianist, value: string) {
    await api.updatePianist(p.id, { max_hours_per_week: value ? Number(value) : null });
    refresh();
  }

  async function removePianist(p: Pianist) {
    if (!confirm(`Remove ${p.name}? This also clears their lesson assignments.`)) return;
    await api.deletePianist(p.id);
    if (selectedId === p.id) setSelectedId(null);
    refresh();
  }

  const selected = pianists.find((p) => p.id === selectedId) ?? null;

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
              <button className="pianist-list-item" onClick={() => setSelectedId(p.id)}>
                <strong>{p.name}</strong>
                <span className="muted">{p.email}</span>
              </button>
              <input
                className="cap-input"
                type="number"
                min="0"
                step="0.5"
                defaultValue={p.max_hours_per_week ?? ""}
                placeholder="cap (h)"
                onBlur={(e) => updateCap(p, e.target.value)}
              />
              <button className="danger-btn small" onClick={() => removePianist(p)}>
                &times;
              </button>
            </li>
          ))}
          {pianists.length === 0 && <p className="muted">No pianists yet -- add one above.</p>}
        </ul>
      </div>
      <div className="pianists-detail">
        {selected ? (
          <>
            <h3>
              {selected.name}'s weekly availability
              {selected.max_hours_per_week ? ` \u2014 ${selected.max_hours_per_week}h cap` : ""}
            </h3>
            <AvailabilityGrid pianist={selected} />
          </>
        ) : (
          <p className="muted">Select a pianist to edit their availability.</p>
        )}
      </div>
    </div>
  );
}
