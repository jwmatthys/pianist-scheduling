import { api } from "../../lib/api";
import type { TutorTutor } from "../../lib/types";
import { TutorAvailabilityDialog } from "./TutorAvailabilityDialog";

export function TutorTutorAvailabilityDialog({
  tutor,
  onSaved,
  onClose,
}: {
  tutor: TutorTutor;
  onSaved: () => Promise<void>;
  onClose: () => void;
}) {
  return (
    <TutorAvailabilityDialog
      title={`Availability: ${tutor.name}`}
      ownerId={tutor.id}
      load={() => api.getTutorTutorAvailability(tutor.id)}
      save={(windows) => api.setTutorTutorAvailability(tutor.id, windows)}
      onSaved={onSaved}
      onClose={onClose}
    />
  );
}
