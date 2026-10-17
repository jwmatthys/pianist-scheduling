import { useCallback, useEffect, useState, type FormEvent } from "react";
import { api } from "../../lib/api";
import type { LessonStudent, LessonTag, LessonTagState } from "../../lib/types";
import { ConfirmDialog, Modal } from "../clinical/ClinicalDialogs";
import { LessonStudentAvailabilityDialog } from "./LessonStudentAvailabilityDialog";
import { LessonStudentImportDialog } from "./LessonStudentImportDialog";

const STATE_ORDER: LessonTagState[] = ["neutral", "required", "preferred", "restricted"];
const STATE_LABELS: Record<LessonTagState, string> = {
  neutral: "Neutral",
  required: "Required",
  preferred: "Preferred",
  restricted: "Restricted",
};

function nextState(state: LessonTagState): LessonTagState {
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
  onEditAvailability,
  onClose,
}: {
  student: LessonStudent | null;
  tags: LessonTag[];
  onChanged: () => Promise<void>;
  onEditAvailability: (student: LessonStudent) => void;
  onClose: () => void;
}) {
  const [studentId, setStudentId] = useState(student?.student_id ?? "");
  const [name, setName] = useState(student?.name ?? "");
  const [email, setEmail] = useState(student?.email ?? "");
  const [states, setStates] = useState<Record<string, LessonTagState>>(() => ({ ...(student?.tag_states ?? {}) }));
  const [errors, setErrors] = useState<string[]>([]);
  const [saving, setSaving] = useState(false);

  async function submit(event: FormEvent) {
    event.preventDefault();
    const problems: string[] = [];
    if (!studentId.trim()) problems.push("Student ID is required.");
    if (!name.trim()) problems.push("Name is required.");
    setErrors(problems);
    if (problems.length) return;
    setSaving(true);
    try {
      // Fields and tag states travel together so one save is one source-data revision.
      const data = { student_id: studentId, name, email, tag_states: states };
      if (student) await api.updateLessonStudent(student.id, data);
      else await api.createLessonStudent(data);
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
            <p className="muted">No Lesson Tags yet. Add tags with Manage Tags on the Instructor Roster tab.</p>
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
          <button type="button" disabled={!student} onClick={() => student && onEditAvailability(student)}>
            Edit Availability
          </button>
          <span className="muted">
            {student
              ? availabilitySummary(student.availability_window_count)
              : "Save the student first to enter availability."}
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
    </Modal>
  );
}

function TagSummary({ student, tags }: { student: LessonStudent; tags: LessonTag[] }) {
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

export function LessonStudentRosterPage() {
  const [students, setStudents] = useState<LessonStudent[]>([]);
  const [tags, setTags] = useState<LessonTag[]>([]);
  const [loaded, setLoaded] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [editing, setEditing] = useState<LessonStudent | "new" | null>(null);
  const [deleting, setDeleting] = useState<LessonStudent | null>(null);
  const [importing, setImporting] = useState(false);
  const [availabilityFor, setAvailabilityFor] = useState<LessonStudent | null>(null);

  const refresh = useCallback(async () => {
    try {
      const [nextStudents, nextTags] = await Promise.all([api.listLessonStudents(), api.listLessonTags()]);
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

  async function confirmDelete() {
    if (!deleting) return;
    const target = deleting;
    setDeleting(null);
    try {
      await api.deleteLessonStudent(target.id);
      await refresh();
    } catch (caught) {
      setError(errorText(caught));
    }
  }

  const editingStudent =
    editing && editing !== "new" ? (students.find((item) => item.id === editing.id) ?? editing) : null;
  const availabilityStudent = availabilityFor
    ? (students.find((item) => item.id === availabilityFor.id) ?? availabilityFor)
    : null;

  return (
    <section className="clinical-sites" aria-labelledby="lesson-students-heading">
      <div className="dialog-heading">
        <h2 id="lesson-students-heading">Student Roster</h2>
        <div className="dialog-actions">
          <button type="button" onClick={() => setImporting(true)}>
            Import Student Roster
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
          onEditAvailability={setAvailabilityFor}
          onClose={() => setEditing(null)}
        />
      )}
      {availabilityStudent && (
        <LessonStudentAvailabilityDialog
          student={availabilityStudent}
          onSaved={refresh}
          onClose={() => setAvailabilityFor(null)}
        />
      )}
      {importing && (
        <LessonStudentImportDialog
          onClose={() => setImporting(false)}
          onImported={async (message) => {
            await refresh();
            setNotice(message);
          }}
        />
      )}
      {deleting && (
        <ConfirmDialog
          title={`Delete ${deleting.name}?`}
          message="This student, their tag selections, availability, and any lesson placement will be removed."
          confirmLabel="Delete Student"
          onConfirm={() => void confirmDelete()}
          onCancel={() => setDeleting(null)}
        />
      )}
    </section>
  );
}
