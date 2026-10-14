import { useCallback, useEffect, useState, type FormEvent } from "react";
import { api } from "../../lib/api";
import type { ClinicalStudent, ClinicalTag, ClinicalTagState } from "../../lib/types";
import { ConfirmDialog, Modal } from "./ClinicalDialogs";
import { ClinicalStudentAvailabilityDialog } from "./ClinicalStudentAvailabilityDialog";
import { ClinicalStudentRosterImportDialog } from "./ClinicalStudentRosterImportDialog";

const STATE_ORDER: ClinicalTagState[] = ["neutral", "required", "preferred", "restricted"];
const STATE_LABELS: Record<ClinicalTagState, string> = {
  neutral: "Neutral",
  required: "Required",
  preferred: "Preferred",
  restricted: "Restricted",
};

function nextState(state: ClinicalTagState): ClinicalTagState {
  return STATE_ORDER[(STATE_ORDER.indexOf(state) + 1) % STATE_ORDER.length];
}

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function availabilitySummary(count: number): string {
  return count === 0 ? "No availability configured" : "Availability entered";
}

function TagStateLegend() {
  return (
    <ul className="clinical-state-legend" aria-label="Tag state legend">
      {STATE_ORDER.map((state) => (
        <li key={state}>
          <span className={`clinical-state-swatch clinical-state-${state}`} aria-hidden="true" />
          {STATE_LABELS[state]}
        </li>
      ))}
    </ul>
  );
}

function StudentDialog({
  student,
  tags,
  onChanged,
  onClose,
}: {
  student: ClinicalStudent | null;
  tags: ClinicalTag[];
  onChanged: () => Promise<void>;
  onClose: () => void;
}) {
  const [studentId, setStudentId] = useState(student?.student_id ?? "");
  const [name, setName] = useState(student?.name ?? "");
  const [email, setEmail] = useState(student?.email ?? "");
  const [states, setStates] = useState<Record<string, ClinicalTagState>>(() => ({ ...(student?.tag_states ?? {}) }));
  const [errors, setErrors] = useState<string[]>([]);
  const [saving, setSaving] = useState(false);
  const [editingAvailability, setEditingAvailability] = useState(false);

  async function submit(event: FormEvent) {
    event.preventDefault();
    const problems: string[] = [];
    if (!studentId.trim()) problems.push("Student ID is required.");
    if (!name.trim()) problems.push("Name is required.");
    setErrors(problems);
    if (problems.length) return;
    setSaving(true);
    try {
      const fields = { student_id: studentId, name, email };
      const saved = student
        ? await api.updateClinicalStudent(student.id, fields)
        : await api.createClinicalStudent(fields);
      const changed = tags.some((tag) => (states[tag.id] ?? "neutral") !== (saved.tag_states[tag.id] ?? "neutral"));
      if (changed) await api.setClinicalStudentTagStates(saved.id, states);
      await onChanged();
      onClose();
    } catch (error) {
      setErrors([errorText(error)]);
      setSaving(false);
    }
  }

  function cycle(tagId: string) {
    setStates((current) => ({ ...current, [tagId]: nextState(current[tagId] ?? "neutral") }));
  }

  return (
    <Modal title={student ? "Edit Student" : "Add Student"} onClose={onClose}>
      <form onSubmit={submit} noValidate>
        {errors.length > 0 && (
          <div className="error-banner" role="alert" style={{ whiteSpace: "pre-line" }}>
            {errors.map((message) => (
              <div key={message}>{message}</div>
            ))}
          </div>
        )}
        <label>
          Student ID
          <input value={studentId} onChange={(event) => setStudentId(event.target.value)} />
        </label>
        <label>
          Name
          <input value={name} onChange={(event) => setName(event.target.value)} />
        </label>
        <label>
          Email
          <input value={email} onChange={(event) => setEmail(event.target.value)} />
        </label>
        <fieldset className="clinical-tag-picker">
          <legend>Tag selections</legend>
          {tags.length === 0 ? (
            <p className="muted">No Clinical Tags yet. Add tags with Manage Tags on the Clinical Sites tab.</p>
          ) : (
            <>
              <p className="dialog-note">Select a tag to cycle Neutral, Required, Preferred, and Restricted.</p>
              <TagStateLegend />
              <div className="clinical-state-buttons">
                {tags.map((tag) => {
                  const state = states[tag.id] ?? "neutral";
                  return (
                    <button
                      key={tag.id}
                      type="button"
                      className={`clinical-state-button clinical-state-${state}`}
                      aria-label={`${tag.name}: ${STATE_LABELS[state]}`}
                      onClick={() => cycle(tag.id)}
                    >
                      {tag.name}
                      {state !== "neutral" && <span className="clinical-state-name">{STATE_LABELS[state]}</span>}
                    </button>
                  );
                })}
              </div>
            </>
          )}
        </fieldset>
        <div className="clinical-availability-row">
          <button
            type="button"
            disabled={!student}
            onClick={() => setEditingAvailability(true)}
          >
            Edit Availability
          </button>
          <span className="muted">
            {student ? availabilitySummary(student.availability_window_count) : "Save the student first to enter availability."}
          </span>
        </div>
        <div className="dialog-actions">
          <button type="button" onClick={onClose}>
            Cancel
          </button>
          <button type="submit" className="primary-btn" disabled={saving}>
            {student ? "Save Student" : "Add Student"}
          </button>
        </div>
      </form>
      {student && editingAvailability && (
        <ClinicalStudentAvailabilityDialog
          student={student}
          onSaved={onChanged}
          onClose={() => setEditingAvailability(false)}
        />
      )}
    </Modal>
  );
}

function TagSummary({ student, tags }: { student: ClinicalStudent; tags: ClinicalTag[] }) {
  const selected = tags.filter((tag) => student.tag_states[tag.id]);
  if (selected.length === 0) return <span className="muted">None</span>;
  return (
    <>
      {selected.map((tag) => (
        <span key={tag.id} className={`clinical-tag clinical-state-${student.tag_states[tag.id]}`}>
          {tag.name}: {STATE_LABELS[student.tag_states[tag.id]]}
        </span>
      ))}
    </>
  );
}

export function ClinicalStudentRosterPage() {
  const [students, setStudents] = useState<ClinicalStudent[]>([]);
  const [tags, setTags] = useState<ClinicalTag[]>([]);
  const [loaded, setLoaded] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [editing, setEditing] = useState<ClinicalStudent | "new" | null>(null);
  const [deleting, setDeleting] = useState<ClinicalStudent | null>(null);
  const [importing, setImporting] = useState(false);
  const [clearing, setClearing] = useState(false);
  const [availabilityFor, setAvailabilityFor] = useState<ClinicalStudent | null>(null);

  const refresh = useCallback(async () => {
    try {
      const [nextStudents, nextTags] = await Promise.all([api.listClinicalStudents(), api.listClinicalTags()]);
      setStudents(nextStudents);
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

  async function confirmClearAll() {
    setClearing(false);
    try {
      const result = await api.clearClinicalStudents();
      await refresh();
      setNotice(`Deleted ${result.removed_count} ${result.removed_count === 1 ? "student" : "students"}.`);
    } catch (caught) {
      setError(errorText(caught));
    }
  }

  async function confirmDelete() {
    if (!deleting) return;
    const target = deleting;
    setDeleting(null);
    try {
      await api.deleteClinicalStudent(target.id);
      await refresh();
    } catch (caught) {
      setError(errorText(caught));
    }
  }

  const editingStudent =
    editing && editing !== "new" ? (students.find((item) => item.id === editing.id) ?? editing) : null;

  return (
    <section className="clinical-sites" aria-labelledby="clinical-roster-heading">
      <div className="dialog-heading">
        <h2 id="clinical-roster-heading">Student Roster</h2>
        <div className="dialog-actions">
          <button type="button" onClick={() => setImporting(true)}>
            Import Student Roster
          </button>
          <button
            type="button"
            className="danger-btn"
            disabled={students.length === 0}
            onClick={() => setClearing(true)}
          >
            Delete all students
          </button>
          <button type="button" className="primary-btn" onClick={() => setEditing("new")}>
            Add Student
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
      {loaded && !error && students.length === 0 ? (
        <p className="muted">No students yet.</p>
      ) : (
        <table className="clinical-table">
          <thead>
            <tr>
              <th>Student ID</th>
              <th>Name</th>
              <th>Email</th>
              <th>Availability</th>
              <th>Tags</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {students.map((item) => (
              <tr key={item.id}>
                <td>{item.student_id}</td>
                <td>{item.name}</td>
                <td>{item.email}</td>
                <td className={item.availability_window_count === 0 ? "clinical-availability-missing" : undefined}>
                  <div className="clinical-availability-cell">
                    <span className={item.availability_window_count === 0 ? undefined : "muted"}>
                      {availabilitySummary(item.availability_window_count)}
                    </span>
                    <button
                      type="button"
                      aria-label={`Edit availability for ${item.name}`}
                      onClick={() => setAvailabilityFor(item)}
                    >
                      Edit Availability
                    </button>
                  </div>
                </td>
                <td>
                  <TagSummary student={item} tags={tags} />
                </td>
                <td className="clinical-row-actions">
                  <button type="button" aria-label={`Edit ${item.name}`} onClick={() => setEditing(item)}>
                    Edit
                  </button>
                  <button
                    type="button"
                    className="danger-btn small"
                    title="Delete Student"
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
        <StudentDialog
          key={editingStudent?.id ?? "new"}
          student={editingStudent}
          tags={tags}
          onChanged={refresh}
          onClose={() => setEditing(null)}
        />
      )}
      {availabilityFor && (
        <ClinicalStudentAvailabilityDialog
          student={availabilityFor}
          onSaved={refresh}
          onClose={() => setAvailabilityFor(null)}
        />
      )}
      {importing && (
        <ClinicalStudentRosterImportDialog
          onClose={() => setImporting(false)}
          onImported={async (message) => {
            await refresh();
            setNotice(message);
          }}
        />
      )}
      {clearing && (
        <ConfirmDialog
          title="Delete all students?"
          message="Every student with their tag selections, availability, and all placements will be removed. Clinical Sites and the Clinical Tag registry are kept. This cannot be undone."
          confirmLabel="Delete all students"
          onConfirm={() => void confirmClearAll()}
          onCancel={() => setClearing(false)}
        />
      )}
      {deleting && (
        <ConfirmDialog
          title={`Delete ${deleting.name}?`}
          message="This student, their tag selections, availability, and any placement will be removed."
          confirmLabel="Delete Student"
          onConfirm={() => void confirmDelete()}
          onCancel={() => setDeleting(null)}
        />
      )}
    </section>
  );
}
