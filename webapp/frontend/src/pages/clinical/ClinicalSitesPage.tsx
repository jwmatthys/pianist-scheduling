import { useCallback, useEffect, useState, type FormEvent } from "react";
import { api } from "../../lib/api";
import { ConfirmDialog, Modal } from "./ClinicalDialogs";
import { ClinicalSitesImportDialog } from "./ClinicalSitesImportDialog";
import { formatMinutes, parseTwelveHourTime } from "../../lib/time";
import type { ClinicalTag, PlacementOpportunity, PlacementOpportunityFields } from "../../lib/types";

const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function TagChip({ tag }: { tag: { name: string; is_unused?: boolean } }) {
  return <span className={`clinical-tag${tag.is_unused ? " clinical-tag-unused" : ""}`}>{tag.name}</span>;
}

interface FormState {
  site_name: string;
  site_address: string;
  contact_name: string;
  contact_email: string;
  contact_phone: string;
  day_of_week: string;
  start: string;
  end: string;
  session_length: string;
  travel: string;
  capacity: string;
}

function toForm(opportunity: PlacementOpportunity | null): FormState {
  return {
    site_name: opportunity?.site_name ?? "",
    site_address: opportunity?.site_address ?? "",
    contact_name: opportunity?.contact_name ?? "",
    contact_email: opportunity?.contact_email ?? "",
    contact_phone: opportunity?.contact_phone ?? "",
    day_of_week: opportunity?.day_of_week ?? "",
    start: opportunity ? formatMinutes(opportunity.start_minute) : "",
    end: opportunity ? formatMinutes(opportunity.end_minute) : "",
    session_length: opportunity ? String(opportunity.session_length_minutes) : "",
    travel: opportunity ? String(opportunity.travel_minutes) : "0",
    capacity: opportunity ? String(opportunity.capacity) : "1",
  };
}

function wholeNumber(value: string): number | null {
  return /^\d+$/.test(value.trim()) ? Number(value.trim()) : null;
}

function validate(form: FormState): { fields?: PlacementOpportunityFields; errors: string[] } {
  const errors: string[] = [];
  if (!form.site_name.trim()) errors.push("Site Name is required.");
  if (!form.day_of_week) errors.push("Day is required.");
  const start = parseTwelveHourTime(form.start);
  const end = parseTwelveHourTime(form.end);
  if (start === null) errors.push("Start Time is required, for example 8:00 AM.");
  if (end === null) errors.push("End Time is required, for example 5:00 PM.");
  if (start !== null && end !== null && end <= start) errors.push("End Time must be after Start Time.");
  const length = wholeNumber(form.session_length);
  if (length === null || length <= 0) errors.push("Session Length must be greater than 0.");
  const travel = wholeNumber(form.travel || "0");
  if (travel === null) errors.push("Travel Time must be 0 or more minutes.");
  const capacity = wholeNumber(form.capacity);
  if (capacity === null || capacity <= 0) errors.push("Capacity must be greater than 0.");
  if (errors.length) return { errors };
  return {
    errors,
    fields: {
      site_name: form.site_name.trim(),
      site_address: form.site_address.trim(),
      contact_name: form.contact_name.trim(),
      contact_email: form.contact_email.trim(),
      contact_phone: form.contact_phone.trim(),
      day_of_week: form.day_of_week,
      start_minute: start!,
      end_minute: end!,
      session_length_minutes: length!,
      travel_minutes: travel!,
      capacity: capacity!,
    },
  };
}

function OpportunityDialog({
  opportunity,
  tags,
  onChanged,
  onClose,
}: {
  opportunity: PlacementOpportunity | null;
  tags: ClinicalTag[];
  onChanged: () => Promise<void>;
  onClose: () => void;
}) {
  const [form, setForm] = useState(() => toForm(opportunity));
  const [selected, setSelected] = useState<string[]>(() => opportunity?.tags.map((tag) => tag.id) ?? []);
  const [errors, setErrors] = useState<string[]>([]);
  const [saving, setSaving] = useState(false);

  const set = (key: keyof FormState) => (event: { target: { value: string } }) =>
    setForm((current) => ({ ...current, [key]: event.target.value }));

  async function toggleTag(tagId: string) {
    const next = selected.includes(tagId) ? selected.filter((id) => id !== tagId) : [...selected, tagId];
    if (!opportunity) {
      setSelected(next);
      return;
    }
    try {
      await api.setClinicalOpportunityTags(opportunity.id, next);
      setSelected(next);
      setErrors([]);
      await onChanged();
    } catch (error) {
      setErrors([errorText(error)]);
    }
  }

  async function submit(event: FormEvent) {
    event.preventDefault();
    const result = validate(form);
    setErrors(result.errors);
    if (!result.fields) return;
    setSaving(true);
    try {
      if (opportunity) await api.updateClinicalOpportunity(opportunity.id, result.fields);
      else await api.createClinicalOpportunity({ ...result.fields, tag_ids: selected });
      await onChanged();
      onClose();
    } catch (error) {
      setErrors([errorText(error)]);
      setSaving(false);
    }
  }

  return (
    <Modal title={opportunity ? "Edit Placement Opportunity" : "Add Placement Opportunity"} onClose={onClose}>
      <form onSubmit={submit} noValidate>
        {errors.length > 0 && (
          <div className="error-banner" role="alert">
            {errors.map((message) => (
              <div key={message}>{message}</div>
            ))}
          </div>
        )}
        <label>
          Site Name
          <input value={form.site_name} onChange={set("site_name")} />
        </label>
        <label>
          Site Address
          <input value={form.site_address} onChange={set("site_address")} />
        </label>
        <label>
          Contact Name
          <input value={form.contact_name} onChange={set("contact_name")} />
        </label>
        <label>
          Contact Email
          <input value={form.contact_email} onChange={set("contact_email")} />
        </label>
        <label>
          Contact Phone
          <input value={form.contact_phone} onChange={set("contact_phone")} />
        </label>
        <label>
          Day of Week
          <select value={form.day_of_week} onChange={set("day_of_week")}>
            <option value="">Select a day</option>
            {DAYS.map((day) => (
              <option key={day} value={day}>
                {day}
              </option>
            ))}
          </select>
        </label>
        <label>
          Start Time
          <input value={form.start} onChange={set("start")} placeholder="8:00 AM" />
        </label>
        <label>
          End Time
          <input value={form.end} onChange={set("end")} placeholder="5:00 PM" />
        </label>
        <label>
          Session Length (minutes)
          <input value={form.session_length} onChange={set("session_length")} inputMode="numeric" />
        </label>
        <label>
          Travel Time (minutes)
          <input value={form.travel} onChange={set("travel")} inputMode="numeric" />
        </label>
        <label>
          Capacity
          <input value={form.capacity} onChange={set("capacity")} inputMode="numeric" />
        </label>
        <fieldset className="clinical-tag-picker">
          <legend>Tags</legend>
          {tags.length === 0 && <p className="muted">No tags yet. Use Manage Tags to add some.</p>}
          {tags.map((tag) => (
            <label key={tag.id} className="clinical-tag-option">
              <input
                type="checkbox"
                checked={selected.includes(tag.id)}
                onChange={() => void toggleTag(tag.id)}
              />
              <TagChip tag={tag} />
            </label>
          ))}
        </fieldset>
        <div className="dialog-actions">
          <button type="button" onClick={onClose}>
            {opportunity ? "Close" : "Cancel"}
          </button>
          <button type="submit" className="primary-btn" disabled={saving}>
            {opportunity ? "Save Changes" : "Add Opportunity"}
          </button>
        </div>
      </form>
    </Modal>
  );
}

function TagManagerDialog({
  tags,
  onChanged,
  onClose,
}: {
  tags: ClinicalTag[];
  onChanged: () => Promise<void>;
  onClose: () => void;
}) {
  const [newName, setNewName] = useState("");
  const [editing, setEditing] = useState<{ id: string; name: string } | null>(null);
  const [deleting, setDeleting] = useState<ClinicalTag | null>(null);
  const [error, setError] = useState("");

  async function run(action: () => Promise<unknown>) {
    try {
      await action();
      setError("");
      await onChanged();
      return true;
    } catch (caught) {
      setError(errorText(caught));
      return false;
    }
  }

  async function add(event: FormEvent) {
    event.preventDefault();
    if (!newName.trim()) {
      setError("Tag name is required.");
      return;
    }
    if (await run(() => api.createClinicalTag(newName))) setNewName("");
  }

  async function rename(event: FormEvent) {
    event.preventDefault();
    if (!editing) return;
    if (await run(() => api.renameClinicalTag(editing.id, editing.name))) setEditing(null);
  }

  async function confirmDelete() {
    if (!deleting) return;
    const target = deleting;
    setDeleting(null);
    await run(() => api.deleteClinicalTag(target.id));
  }

  return (
    <Modal title="Manage Tags" onClose={onClose} className="clinical-wide-dialog">
      <div className="clinical-tag-manager">
      {error && (
        <div className="error-banner" role="alert">
          {error}
        </div>
      )}
      <form onSubmit={add} className="clinical-tag-add">
        <label>
          New tag
          <input value={newName} onChange={(event) => setNewName(event.target.value)} />
        </label>
        <button type="submit" className="primary-btn">
          Add Tag
        </button>
      </form>
      <p className="dialog-note">Greyed-out tags are not used by any Placement Opportunity. They remain available.</p>
      <ul className="clinical-tag-list">
        {tags.length === 0 && <li className="muted">No tags yet.</li>}
        {tags.map((tag) => (
          <li key={tag.id}>
            {editing?.id === tag.id ? (
              <form onSubmit={rename} className="clinical-tag-add">
                <input
                  aria-label={`New name for ${tag.name}`}
                  value={editing.name}
                  onChange={(event) => setEditing({ id: tag.id, name: event.target.value })}
                />
                <button type="submit" className="primary-btn">
                  Save
                </button>
                <button type="button" onClick={() => setEditing(null)}>
                  Cancel
                </button>
              </form>
            ) : (
              <>
                <TagChip tag={tag} />
                <span className="muted">
                  {tag.opportunity_count} {tag.opportunity_count === 1 ? "opportunity" : "opportunities"}
                </span>
                <button
                  type="button"
                  className="clinical-tag-action"
                  aria-label={`Rename ${tag.name}`}
                  onClick={() => setEditing({ id: tag.id, name: tag.name })}
                >
                  Rename
                </button>
                <button
                  type="button"
                  className="danger-btn small clinical-tag-delete"
                  title="Delete tag"
                  aria-label={`Delete ${tag.name}`}
                  onClick={() => setDeleting(tag)}
                >
                  &times;
                </button>
              </>
            )}
          </li>
        ))}
      </ul>
      <div className="dialog-actions">
        <button type="button" onClick={onClose}>
          Done
        </button>
      </div>
      </div>
      {deleting && (
        <ConfirmDialog
          title={`Delete tag "${deleting.name}"?`}
          message={
            deleting.student_count > 0
              ? `This tag is currently used by ${deleting.student_count} ${
                  deleting.student_count === 1 ? "student" : "students"
                }. Deleting it will remove all associated student preference, requirement, and restriction settings.`
              : "Deleting it will remove it from all Placement Opportunities."
          }
          confirmLabel="Delete Tag"
          onConfirm={() => void confirmDelete()}
          onCancel={() => setDeleting(null)}
        />
      )}
    </Modal>
  );
}

export function ClinicalSitesPage() {
  const [opportunities, setOpportunities] = useState<PlacementOpportunity[]>([]);
  const [tags, setTags] = useState<ClinicalTag[]>([]);
  const [loaded, setLoaded] = useState(false);
  const [error, setError] = useState("");
  const [editing, setEditing] = useState<PlacementOpportunity | "new" | null>(null);
  const [deleting, setDeleting] = useState<PlacementOpportunity | null>(null);
  const [managingTags, setManagingTags] = useState(false);
  const [importing, setImporting] = useState(false);
  const [notice, setNotice] = useState("");

  const refresh = useCallback(async () => {
    try {
      const [nextOpportunities, nextTags] = await Promise.all([
        api.listClinicalOpportunities(),
        api.listClinicalTags(),
      ]);
      setOpportunities(nextOpportunities);
      setTags(nextTags);
      setError("");
    } catch (caught) {
      setError(errorText(caught));
    } finally {
      setLoaded(true);
    }
  }, []);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  async function confirmDeleteOpportunity() {
    if (!deleting) return;
    const target = deleting;
    setDeleting(null);
    try {
      await api.deleteClinicalOpportunity(target.id);
      await refresh();
    } catch (caught) {
      setError(errorText(caught));
    }
  }

  // Keep the open edit dialog bound to the latest persisted data.
  const editingOpportunity =
    editing && editing !== "new" ? (opportunities.find((item) => item.id === editing.id) ?? editing) : null;

  return (
    <section className="clinical-sites" aria-labelledby="clinical-sites-heading">
      <div className="dialog-heading">
        <h2 id="clinical-sites-heading">Clinical Sites</h2>
        <div className="dialog-actions">
          <button type="button" onClick={() => setImporting(true)}>
            Import Clinical Sites
          </button>
          <button type="button" onClick={() => setManagingTags(true)}>
            Manage Tags
          </button>
          <button type="button" className="primary-btn" onClick={() => setEditing("new")}>
            Add Placement Opportunity
          </button>
        </div>
      </div>
      {error && (
        <div className="error-banner" role="alert">
          {error}
        </div>
      )}
      {notice && (
        <div className="availability-success" role="status">
          {notice}
        </div>
      )}
      {loaded && !error && opportunities.length === 0 ? (
        <p className="muted">No Placement Opportunities yet.</p>
      ) : (
        <table className="clinical-table">
          <thead>
            <tr>
              <th>Site Name</th>
              <th>Day</th>
              <th>Start Time</th>
              <th>End Time</th>
              <th>Session Length</th>
              <th>Travel Time</th>
              <th>Capacity</th>
              <th>Tags</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {opportunities.map((item) => (
              <tr key={item.id}>
                <td>{item.site_name}</td>
                <td>{item.day_of_week}</td>
                <td>{formatMinutes(item.start_minute)}</td>
                <td>{formatMinutes(item.end_minute)}</td>
                <td>{item.session_length_minutes} min</td>
                <td>{item.travel_minutes} min</td>
                <td>{item.capacity}</td>
                <td>
                  {item.tags.map((tag) => (
                    <TagChip key={tag.id} tag={tag} />
                  ))}
                </td>
                <td className="clinical-row-actions">
                  <button
                    type="button"
                    aria-label={`Edit ${item.site_name}`}
                    onClick={() => setEditing(item)}
                  >
                    Edit
                  </button>
                  <button
                    type="button"
                    className="danger-btn small"
                    title="Delete Placement Opportunity"
                    aria-label={`Delete ${item.site_name}`}
                    onClick={() => setDeleting(item)}
                  >
                    &times;
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
      {editing && (
        <OpportunityDialog
          key={editingOpportunity?.id ?? "new"}
          opportunity={editingOpportunity}
          tags={tags}
          onChanged={refresh}
          onClose={() => setEditing(null)}
        />
      )}
      {importing && (
        <ClinicalSitesImportDialog
          onClose={() => setImporting(false)}
          onImported={async (message) => {
            await refresh();
            setNotice(message);
          }}
        />
      )}
      {managingTags && (
        <TagManagerDialog tags={tags} onChanged={refresh} onClose={() => setManagingTags(false)} />
      )}
      {deleting && (
        <ConfirmDialog
          title={`Delete ${deleting.site_name} on ${deleting.day_of_week}?`}
          message={
            deleting.placement_count > 0
              ? `${deleting.placement_count} ${
                  deleting.placement_count === 1 ? "placement is" : "placements are"
                } assigned to this Placement Opportunity and will be removed.`
              : "This Placement Opportunity will be removed."
          }
          confirmLabel="Delete Placement Opportunity"
          onConfirm={() => void confirmDeleteOpportunity()}
          onCancel={() => setDeleting(null)}
        />
      )}
    </section>
  );
}
