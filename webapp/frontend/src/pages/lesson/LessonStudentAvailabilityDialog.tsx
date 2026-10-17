import { api } from "../../lib/api";
import type { LessonStudent } from "../../lib/types";
import { LessonAvailabilityDialog } from "./LessonAvailabilityDialog";

export function LessonStudentAvailabilityDialog({
  student,
  onSaved,
  onClose,
}: {
  student: LessonStudent;
  onSaved: () => Promise<void>;
  onClose: () => void;
}) {
  return (
    <LessonAvailabilityDialog
      title={`Availability: ${student.name}`}
      ownerId={student.id}
      load={() => api.getLessonStudentAvailability(student.id)}
      save={(windows) => api.setLessonStudentAvailability(student.id, windows)}
      onSaved={onSaved}
      onClose={onClose}
    />
  );
}
