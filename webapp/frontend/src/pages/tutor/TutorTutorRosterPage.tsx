import { useCallback, useEffect, useState, type FormEvent } from "react";
import { api } from "../../lib/api";
import type { TutorTutor, TutorTutorFields, TutorTag } from "../../lib/types";
import { ConfirmDialog, Modal } from "../clinical/ClinicalDialogs";
import { TutorTutorAvailabilityDialog } from "./TutorTutorAvailabilityDialog";
import { TutorTutorImportDialog } from "./TutorTutorImportDialog";
import { TutorTagChip, TutorTagManagerDialog } from "./TutorTagManagerDialog";
import { formatHours } from "../../lib/availabilityHours";

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

interface FormState {
  name: string;
  email: string;
  room: string;
  lesson_length: string;
  capacity: string;
}

function toForm(tutor: TutorTutor | null): FormState {
  return {
    name: tutor?.name ?? "",
    email: tutor?.email ?? "",
    room: tutor?.room ?? "",
    lesson_length: String(tutor?.lesson_length_minutes ?? 60),
    capacity: String(tutor?.capacity ?? 1),
  };
}

function wholeNumber(value: string): number | null {
  return /^\d+$/.test(value.trim()) ? Number(value.trim()) : null;
}

function validate(form: FormState): { fields?: TutorTutorFields; errors: string[] } {
  const errors: string[] = [];
  if (!form.name.trim()) errors.push("Tutor Name is required.");
  const length = wholeNumber(form.lesson_length);
  if (length === null || length <= 0) errors.push("Lesson Length must be greater than 0.");
  const capacity = wholeNumber(form.capacity);
  if (capacity === null || capacity <= 0) errors.push("Capacity must be greater than 0.");
  if (errors.length) return { errors };
  return {
    errors,
    fields: {
      name: form.name.trim(),
      email: form.email.trim(),
      room: form.room.trim(),
      lesson_length_minutes: length!,
      capacity: capacity!,
    },
  };
}

function TutorDialog({
  tutor,
  tags,
  onChanged,
  onEditAvailability,
  onClose,
}: {
  tutor: TutorTutor | null;
  tags: TutorTag[];
  onChanged: () => Promise<void>;
  onEditAvailability: (tutor: TutorTutor) => void;
  onClose: () => void;
}) {
  const [form, setForm] = useState(() => toForm(tutor));
  const [selected, setSelected] = useState<string[]>(() => tutor?.tags.map((tag) => tag.id) ?? []);
  const [errors, setErrors] = useState<string[]>([]);
  const [saving, setSaving] = useState(false);

  const set = (key: keyof FormState) => (event: { target: { value: string } }) =>
    setForm((current) => ({ ...current, [key]: event.target.value }));

  function toggleTag(tagId: string) {
    setSelected((current) => (current.includes(tagId) ? current.filter((id) => id !== tagId) : [...current, tagId]));
  }

  async function submit(event: FormEvent) {
    event.preventDefault();
    const result = validate(form);
    setErrors(result.errors);
    if (!result.fields) return;
    setSaving(true);
    try {
      const data = { ...result.fields, tag_ids: selected };
      if (tutor) await api.updateTutorTutor(tutor.id, data);
      else await api.createTutorTutor(data);
      await onChanged();
      onClose();
    } catch (error) {
      setErrors([errorText(error)]);
      setSaving(false);
    }
  }

  return (
    <Modal title={tutor ? "Edit Tutor" : "Add Tutor"} onClose={onClose}>
      <form onSubmit={submit} noValidate>
        {errors.length > 0 && (
          <div className="error-banner" role="alert">
            {errors.map((message) => (
              <div key={message}>{message}</div>
            ))}
          </div>
        )}
        <label>
          Tutor Name
          <input value={form.name} onChange={set("name")} />
        </label>
        <label>
          Tutor Email
          <input value={form.email} onChange={set("email")} />
        </label>
        <label>
          Room
          <input value={form.room} onChange={set("room")} />
        </label>
        <label>
          Lesson Length (minutes)
          <input value={form.lesson_length} onChange={set("lesson_length")} inputMode="numeric" />
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
              <input type="checkbox" checked={selected.includes(tag.id)} onChange={() => toggleTag(tag.id)} />
              <TutorTagChip tag={tag} />
            </label>
          ))}
        </fieldset>
        <div className="dialog-actions">
          {tutor && (
            <button type="button" onClick={() => onEditAvailability(tutor)}>
              Edit Availability
            </button>
          )}
          <button type="button" onClick={onClose}>
            Cancel
          </button>
          <button type="submit" className={tutor ? "" : "primary-btn"} disabled={saving}>
            {tutor ? "Save Changes" : "Add Tutor"}
          </button>
        </div>
      </form>
    </Modal>
  );
}

export function TutorTutorRosterPage() {
  const [tutors, setTutors] = useState<TutorTutor[]>([]);
  const [tags, setTags] = useState<TutorTag[]>([]);
  const [loaded, setLoaded] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [editing, setEditing] = useState<TutorTutor | "new" | null>(null);
  const [deleting, setDeleting] = useState<TutorTutor | null>(null);
  const [availabilityFor, setAvailabilityFor] = useState<TutorTutor | null>(null);
  const [managingTags, setManagingTags] = useState(false);
  const [importing, setImporting] = useState(false);

  const refresh = useCallback(async () => {
    try {
      const [nextTutors, nextTags] = await Promise.all([api.listTutorTutors(), api.listTutorTags()]);
      setTutors(nextTutors);
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

  async function confirmDelete() {
    if (!deleting) return;
    const target = deleting;
    setDeleting(null);
    try {
      await api.deleteTutorTutor(target.id);
      await refresh();
    } catch (caught) {
      setError(errorText(caught));
    }
  }

  const editingTutor =
    editing && editing !== "new" ? (tutors.find((item) => item.id === editing.id) ?? editing) : null;

  return (
    <section className="clinical-sites" aria-labelledby="tutor-tutors-heading">
      <div className="dialog-heading">
        <h2 id="tutor-tutors-heading">Tutor Roster</h2>
        <div className="dialog-actions">
          <button type="button" onClick={() => setImporting(true)}>
            Import Tutor Roster
          </button>
          <button type="button" onClick={() => setManagingTags(true)}>
            Manage Tags
          </button>
          <button type="button" className="primary-btn" onClick={() => setEditing("new")}>
            Add Tutor
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
      {loaded && !error && tutors.length === 0 ? (
        <p className="muted">No tutors yet.</p>
      ) : (
        <table className="clinical-table">
          <thead>
            <tr>
              <th>Tutor Name</th>
              <th>Email</th>
              <th>Room</th>
              <th>Lesson Length</th>
              <th>Capacity</th>
              <th>Availability</th>
              <th>Tags</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {tutors.map((item) => (
              <tr key={item.id}>
                <td>{item.name}</td>
                <td>{item.email}</td>
                <td>{item.room}</td>
                <td>{item.lesson_length_minutes} min</td>
                <td>{item.capacity}</td>
                <td>
                  {item.availability_window_count === 0 ? (
                    "Not set"
                  ) : (
                    <>
                      <div>Available: {formatHours(item.available_minutes)}</div>
                      <div>Tentative: {formatHours(item.tentative_minutes)}</div>
                    </>
                  )}
                </td>
                <td>
                  {item.tags.map((tag) => (
                    <TutorTagChip key={tag.id} tag={tag} />
                  ))}
                </td>
                <td className="clinical-row-actions">
                  <button type="button" aria-label={`Edit ${item.name}`} onClick={() => setEditing(item)}>
                    Edit
                  </button>
                  <button
                    type="button"
                    aria-label={`Availability for ${item.name}`}
                    onClick={() => setAvailabilityFor(item)}
                  >
                    Availability
                  </button>
                  <button
                    type="button"
                    className="danger-btn small"
                    title="Delete tutor"
                    aria-label={`Delete ${item.name}`}
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
        <TutorDialog
          key={editingTutor?.id ?? "new"}
          tutor={editingTutor}
          tags={tags}
          onChanged={refresh}
          onEditAvailability={setAvailabilityFor}
          onClose={() => setEditing(null)}
        />
      )}
      {availabilityFor && (
        <TutorTutorAvailabilityDialog
          tutor={availabilityFor}
          onSaved={refresh}
          onClose={() => setAvailabilityFor(null)}
        />
      )}
      {importing && (
        <TutorTutorImportDialog
          onClose={() => setImporting(false)}
          onImported={async (message) => {
            await refresh();
            setNotice(message);
          }}
        />
      )}
      {managingTags && (
        <TutorTagManagerDialog tags={tags} onChanged={refresh} onClose={() => setManagingTags(false)} />
      )}
      {deleting && (
        <ConfirmDialog
          title={`Delete ${deleting.name}?`}
          message="This tutor and their availability will be removed. Existing tutor placements will be retained in the stale schedule with verification warnings."
          confirmLabel="Delete Tutor"
          onConfirm={() => void confirmDelete()}
          onCancel={() => setDeleting(null)}
        />
      )}
    </section>
  );
}
