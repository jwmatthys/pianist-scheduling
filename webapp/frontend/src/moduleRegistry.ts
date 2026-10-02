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
    description: "Assign pianists to student lessons using availability and workload.",
    availability: "available",
  },
  {
    key: "juries",
    name: "Performance Juries",
    description: "Configure lesson-based jury participation, panels, and pianist availability.",
    availability: "available",
  },
  {
    key: "clinical",
    name: "Clinical Placements",
    description: "Plan Music Therapy student placements with clinical sites.",
    availability: "shell",
  },
];

export function formatSessionIdentity(session: SchedulingSession | null): string {
  if (!session) return "Loading active session";
  const term = session.year && !session.term_label.includes(String(session.year))
    ? `${session.term_label} ${session.year}`
    : session.term_label;
  return `${session.institution_name} · ${session.program_name} · ${term}`;
}