import type { SchedulingSession } from "./lib/types";

export type ModuleKey = "accompanist" | "clinical" | "juries";
export type ModuleAvailability = "available" | "shell";

export interface SchedulingModule {
  key: ModuleKey;
  name: string;
  description: string;
  availability: ModuleAvailability;
}

export const SCHEDULING_MODULES: SchedulingModule[] = [
  {
    key: "accompanist",
    name: "Accompanist Scheduling",
    description: "Assign collaborative pianists to student lessons.",
    availability: "available",
  },
  {
    key: "juries",
    name: "Performance Jury Scheduling",
    description: "Configure lesson-based jury participation, panels, and pianist availability.",
    availability: "available",
  },
  {
    key: "clinical",
    name: "Clinical Placements",
    description: "Plan music therapy and music education student placements with clinical sites and field experiences.",
    availability: "shell",
  },
];

export function formatSessionIdentity(session: SchedulingSession | null): string {
  if (!session) return "Loading active session";
  return `${session.institution_name} · ${session.program_name} · ${session.term_label}`;
}