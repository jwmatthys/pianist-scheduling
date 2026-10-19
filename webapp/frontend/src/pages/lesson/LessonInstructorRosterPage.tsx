import { useCallback, useEffect, useState, type FormEvent } from "react";
import { api } from "../../lib/api";
import type { LessonInstructor, LessonInstructorFields, LessonTag } from "../../lib/types";
import { ConfirmDialog, Modal } from "../clinical/ClinicalDialogs";
import { LessonInstructorAvailabilityDialog } from "./LessonInstructorAvailabilityDialog";
import { LessonInstructorImportDialog } from "./LessonInstructorImportDialog";
import { LessonTagChip, LessonTagManagerDialog } from "./LessonTagManagerDialog";

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

function toForm(instructor: LessonInstructor | null): FormState {
  return {
    name: instructor?.name ?? "",
    email: instructor?.email ?? "",
    room: instructor?.room ?? "",
    lesson_length: String(instructor?.lesson_length_minutes ?? 60),
    capacity: String(instructor?.capacity ?? 1),
  };
}

function wholeNumber(value: string): number | null {
  return /^\d+$/.test(value.trim()) ? Number(value.trim()) : null;
}

function validate(form: FormState): { fields?: LessonInstructorFields; errors: string[] } {
  const errors: string[] = [];
  if (!form.name.trim()) errors.push("Instructor Name is required.");
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

function InstructorDialog({
  instructor,
  tags,
  onChanged,
  onEditAvailability,
  onClose,
}: {
  instructor: LessonInstructor | null;
  tags: LessonTag[];
  onChanged: () => Promise<void>;
  onEditAvailability: (instructor: LessonInstructor) => void;
  onClose: () => void;
}) {
  const [form, setForm] = useState(() => toForm(instructor));
  const [selected, setSelected] = useState<string[]>(() => instructor?.tags.map((tag) => tag.id) ?? []);
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
      if (instructor) await api.updateLessonInstructor(instructor.id, data);
      else await api.createLessonInstructor(data);
      await onChanged();
      onClose();
    } catch (error) {
      setErrors([errorText(error)]);
      setSaving(false);
    }
  }

  return (
    <Modal title={instructor ? "Edit Instructor" : "Add Instructor"} onClose={onClose}>
      <form onSubmit={submit} noValidate>
        {errors.length > 0 && (
          <div className="error-banner" role="alert">
            {errors.map((message) => (
              <div key={message}>{message}</div>
            ))}
          </div>
        )}
        <label>
          Instructor Name
          <input value={form.name} onChange={set("name")} />
        </label>
        <label>
          Instructor Email
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
              <LessonTagChip tag={tag} />
            </label>
          ))}
        </fieldset>
        <div className="dialog-actions">
          {instructor && (
            <button type="button" onClick={() => onEditAvailability(instructor)}>
              Edit Availability
            </button>
          )}
          <button type="button" onClick={onClose}>
            Cancel
          </button>
          <button type="submit" className="primary-btn" disabled={saving}>
            {instructor ? "Save Changes" : "Add Instructor"}
          </button>
        </div>
      </form>
    </Modal>
  );
}

export function LessonInstructorRosterPage() {
  const [instructors, setInstructors] = useState<LessonInstructor[]>([]);
  const [tags, setTags] = useState<LessonTag[]>([]);
  const [loaded, setLoaded] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [editing, setEditing] = useState<LessonInstructor | "new" | null>(null);
  const [deleting, setDeleting] = useState<LessonInstructor | null>(null);
  const [availabilityFor, setAvailabilityFor] = useState<LessonInstructor | null>(null);
  const [managingTags, setManagingTags] = useState(false);
  const [importing, setImporting] = useState(false);

  const refresh = useCallback(async () => {
    try {
      const [nextInstructors, nextTags] = await Promise.all([api.listLessonInstructors(), api.listLessonTags()]);
      setInstructors(nextInstructors);
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
      await api.deleteLessonInstructor(target.id);
      await refresh();
    } catch (caught) {
      setError(errorText(caught));
    }
  }

  const editingInstructor =
    editing && editing !== "new" ? (instructors.find((item) => item.id === editing.id) ?? editing) : null;

  return (
    <section className="clinical-sites" aria-labelledby="lesson-instructors-heading">
      <div className="dialog-heading">
        <h2 id="lesson-instructors-heading">Instructor Roster</h2>
        <div className="dialog-actions">
          <button type="button" onClick={() => setImporting(true)}>
            Import Instructor Roster
          </button>
          <button type="button" onClick={() => setManagingTags(true)}>
            Manage Tags
          </button>
          <button type="button" className="primary-btn" onClick={() => setEditing("new")}>
            Add Instructor
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
      {loaded && !error && instructors.length === 0 ? (
        <p className="muted">No instructors yet.</p>
      ) : (
        <table className="clinical-table">
          <thead>
            <tr>
              <th>Instructor Name</th>
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
            {instructors.map((item) => (
              <tr key={item.id}>
                <td>{item.name}</td>
                <td>{item.email}</td>
                <td>{item.room}</td>
                <td>{item.lesson_length_minutes} min</td>
                <td>{item.capacity}</td>
                <td>
                  {item.availability_window_count === 0
                    ? "Not set"
                    : `${item.availability_window_count} ${item.availability_window_count === 1 ? "window" : "windows"}`}
                </td>
                <td>
                  {item.tags.map((tag) => (
                    <LessonTagChip key={tag.id} tag={tag} />
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
                    title="Delete instructor"
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
        <InstructorDialog
          key={editingInstructor?.id ?? "new"}
          instructor={editingInstructor}
          tags={tags}
          onChanged={refresh}
          onEditAvailability={setAvailabilityFor}
          onClose={() => setEditing(null)}
        />
      )}
      {availabilityFor && (
        <LessonInstructorAvailabilityDialog
          instructor={availabilityFor}
          onSaved={refresh}
          onClose={() => setAvailabilityFor(null)}
        />
      )}
      {importing && (
        <LessonInstructorImportDialog
          onClose={() => setImporting(false)}
          onImported={async (message) => {
            await refresh();
            setNotice(message);
          }}
        />
      )}
      {managingTags && (
        <LessonTagManagerDialog tags={tags} onChanged={refresh} onClose={() => setManagingTags(false)} />
      )}
      {deleting && (
        <ConfirmDialog
          title={`Delete ${deleting.name}?`}
          message="This instructor and their availability will be removed. Existing lesson placements will be retained in the stale schedule with verification warnings."
          confirmLabel="Delete Instructor"
          onConfirm={() => void confirmDelete()}
          onCancel={() => setDeleting(null)}
        />
      )}
    </section>
  );
}
