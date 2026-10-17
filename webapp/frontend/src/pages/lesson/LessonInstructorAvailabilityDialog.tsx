import { api } from "../../lib/api";
import type { LessonInstructor } from "../../lib/types";
import { LessonAvailabilityDialog } from "./LessonAvailabilityDialog";

export function LessonInstructorAvailabilityDialog({
  instructor,
  onSaved,
  onClose,
}: {
  instructor: LessonInstructor;
  onSaved: () => Promise<void>;
  onClose: () => void;
}) {
  return (
    <LessonAvailabilityDialog
      title={`Availability: ${instructor.name}`}
      ownerId={instructor.id}
      load={() => api.getLessonInstructorAvailability(instructor.id)}
      save={(windows) => api.setLessonInstructorAvailability(instructor.id, windows)}
      onSaved={onSaved}
      onClose={onClose}
    />
  );
}
