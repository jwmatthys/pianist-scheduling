import { api } from "../../lib/api";
import type { TutorStudent } from "../../lib/types";
import { TutorAvailabilityDialog } from "./TutorAvailabilityDialog";

export function TutorStudentAvailabilityDialog({
  student,
  onSaved,
  onClose,
}: {
  student: TutorStudent;
  onSaved: () => Promise<void>;
  onClose: () => void;
}) {
  return (
    <TutorAvailabilityDialog
      title={`Availability: ${student.name}`}
      ownerId={student.id}
      load={() => api.getTutorStudentAvailability(student.id)}
      save={(windows) => api.setTutorStudentAvailability(student.id, windows)}
      onSaved={onSaved}
      onClose={onClose}
    />
  );
}
