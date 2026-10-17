import { useState, type FormEvent } from "react";
import { api } from "../../lib/api";
import type { LessonTag } from "../../lib/types";
import { ConfirmDialog, Modal } from "../clinical/ClinicalDialogs";

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export function LessonTagChip({ tag }: { tag: { name: string; is_unused?: boolean } }) {
  return <span className={`clinical-tag${tag.is_unused ? " clinical-tag-unused" : ""}`}>{tag.name}</span>;
}

function usageText(tag: LessonTag): string {
  const parts = [`${tag.instructor_count} ${tag.instructor_count === 1 ? "instructor" : "instructors"}`];
  if (tag.student_count > 0) parts.push(`${tag.student_count} ${tag.student_count === 1 ? "student" : "students"}`);
  return parts.join(", ");
}

export function LessonTagManagerDialog({
  tags,
  onChanged,
  onClose,
}: {
  tags: LessonTag[];
  onChanged: () => Promise<void>;
  onClose: () => void;
}) {
  const [newName, setNewName] = useState("");
  const [editing, setEditing] = useState<{ id: string; name: string } | null>(null);
  const [deleting, setDeleting] = useState<LessonTag | null>(null);
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
    if (await run(() => api.createLessonTag(newName))) setNewName("");
  }

  async function rename(event: FormEvent) {
    event.preventDefault();
    if (!editing) return;
    if (await run(() => api.renameLessonTag(editing.id, editing.name))) setEditing(null);
  }

  async function confirmDelete() {
    if (!deleting) return;
    const target = deleting;
    setDeleting(null);
    await run(() => api.deleteLessonTag(target.id));
  }

  const inUse = deleting ? deleting.instructor_count + deleting.student_count > 0 : false;

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
        <p className="dialog-note">Greyed-out tags are not used by any instructor or student. They remain available.</p>
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
                  <LessonTagChip tag={tag} />
                  <span className="muted">{usageText(tag)}</span>
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
            inUse
              ? `This tag is currently used by ${usageText(deleting)}. Deleting it will remove it from them and discard any related tag settings.`
              : "This tag is not in use."
          }
          confirmLabel="Delete Tag"
          onConfirm={() => void confirmDelete()}
          onCancel={() => setDeleting(null)}
        />
      )}
    </Modal>
  );
}
