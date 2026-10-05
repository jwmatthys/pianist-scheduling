import type {
  AvailabilityImportApplyResult,
  AvailabilityImportInspection,
  AvailabilityImportMapping,
  AvailabilityImportPreview,
  AvailabilitySlot,
  AccompanistFinalizationState,
  ImportCommitResult,
  ImportPreview,
  JuryAvailability,
  JuryAvailabilityClearResult,
  JuryAvailabilityInput,
  JurySynchronizationSummary,
  JuryConfiguration,
  JuryLessonEntry,
  JuryPanel,
  JuryPanelInput,
  JuryPanelImportInspection,
  JuryPanelImportMapping,
  JuryPanelImportResult,
  JuryReadiness,
  JuryScheduleHistoryItem,
  JuryScheduleResult,
  Lesson,
  Pianist,
  RunAssignmentResult,
  SchedulingSession,
  SchedulingSessionInput,
} from "./types";
import { getApiBase } from "./platform";

async function request<T>(path: string, options: RequestInit = {}): Promise<T> {
  const res = await fetch(`${await getApiBase()}${path}`, {
    headers: options.body instanceof FormData ? undefined : { "Content-Type": "application/json" },
    ...options,
  });
  if (!res.ok) {
    let detail = res.statusText;
    try {
      const body = await res.json();
      detail = typeof body.detail === "string"
        ? body.detail
        : body.detail?.message
          ? `${body.detail.code ?? "Request failed"}: ${body.detail.message}${
              Array.isArray(body.detail.issues) && body.detail.issues.length ? `\n${body.detail.issues.join("\n")}` : ""
            }`
          : JSON.stringify(body.detail ?? body);
    } catch {
      /* ignore */
    }
    throw new Error(`${res.status}: ${detail}`);
  }
  if (res.status === 204) return undefined as T;
  const contentType = res.headers.get("content-type") ?? "";
  if (contentType.includes("application/json")) return res.json();
  return res.text() as unknown as T;
}

export const api = {
  // Active Scheduling Session
  getSession: () => request<SchedulingSession>("/api/session"),
  updateSession: (data: SchedulingSessionInput) =>
    request<SchedulingSession>("/api/session", { method: "PUT", body: JSON.stringify(data) }),
  createSession: (data: SchedulingSessionInput) =>
    request<SchedulingSession>("/api/session/new", {
      method: "POST",
      body: JSON.stringify({ ...data, confirmed: true }),
    }),
  exportSession: async () => {
    const response = await fetch(`${await getApiBase()}/api/session/export`);
    if (!response.ok) throw new Error(`Session export failed (${response.status}).`);
    return new Uint8Array(await response.arrayBuffer());
  },
  restoreSession: (archive: Uint8Array) =>
    request<SchedulingSession>("/api/session/restore?confirmed=true", {
      method: "POST",
      headers: { "Content-Type": "application/vnd.music-program-scheduler.session+zip" },
      body: archive,
    }),

  // Pianists
  listPianists: () => request<Pianist[]>("/api/pianists"),
  createPianist: (data: { name: string; email?: string; max_hours_per_week?: number | null }) =>
    request<Pianist>("/api/pianists", { method: "POST", body: JSON.stringify(data) }),
  updatePianist: (id: number, data: Partial<Pick<Pianist, "name" | "email" | "max_hours_per_week">>) =>
    request<Pianist>(`/api/pianists/${id}`, { method: "PATCH", body: JSON.stringify(data) }),
  deletePianist: (id: number) => request<{ ok: boolean }>(`/api/pianists/${id}`, { method: "DELETE" }),

  getAvailability: (pianistId: number) =>
    request<AvailabilitySlot[]>(`/api/pianists/${pianistId}/availability`),
  setAvailability: (pianistId: number, slots: AvailabilitySlot[]) =>
    request<AvailabilitySlot[]>(`/api/pianists/${pianistId}/availability`, {
      method: "PUT",
      body: JSON.stringify({ slots }),
    }),

  inspectAvailabilityFile: (file: File) => {
    const form = new FormData();
    form.append("file", file);
    return request<AvailabilityImportInspection>("/api/availability-import/inspect", {
      method: "POST",
      body: form,
    });
  },
  inspectAvailabilitySheet: (uploadToken: string, sheetName: string) =>
    request<AvailabilityImportInspection>(
      `/api/availability-import/inspect/${encodeURIComponent(uploadToken)}?sheet_name=${encodeURIComponent(sheetName)}`
    ),
  previewAvailabilityImport: (
    uploadToken: string,
    sheetName: string | null,
    mapping: AvailabilityImportMapping,
  ) => request<AvailabilityImportPreview>("/api/availability-import/preview", {
    method: "POST",
    body: JSON.stringify({
      upload_token: uploadToken,
      sheet_name: sheetName,
      mapping,
    }),
  }),
  applyAvailabilityImport: (previewToken: string) =>
    request<AvailabilityImportApplyResult>("/api/availability-import/apply", {
      method: "POST",
      body: JSON.stringify({ preview_token: previewToken, confirmed: true }),
    }),

  // Lessons
  listLessons: () => request<Lesson[]>("/api/lessons"),
  createLesson: (data: Partial<Lesson>) =>
    request<Lesson>("/api/lessons", { method: "POST", body: JSON.stringify(data) }),
  updateLesson: (id: number, data: Partial<Lesson> & { clear_assigned_pianist?: boolean }) =>
    request<Lesson>(`/api/lessons/${id}`, { method: "PATCH", body: JSON.stringify(data) }),
  deleteLesson: (id: number) => request<{ ok: boolean }>(`/api/lessons/${id}`, { method: "DELETE" }),
  deleteAllLessons: () => request<{ ok: boolean }>("/api/lessons", { method: "DELETE" }),

  // Import
  previewImport: (file: File) => {
    const form = new FormData();
    form.append("file", file);
    return request<ImportPreview>("/api/import/preview", { method: "POST", body: form });
  },
  targetFields: () => request<{ fields: string[] }>("/api/import/target-fields"),
  commitImport: (uploadToken: string, mapping: Record<string, string | null>, profileName?: string) =>
    request<ImportCommitResult>("/api/import/commit", {
      method: "POST",
      body: JSON.stringify({ upload_token: uploadToken, mapping, save_profile_name: profileName ?? null }),
    }),
  listProfiles: () => request<{ id: number; name: string; mapping: Record<string, string> }[]>(
    "/api/import/profiles"
  ),

  // Assignments
  runAssignment: () => request<RunAssignmentResult>("/api/assignments/run", { method: "POST" }),
  clearAssignments: () => request<{ ok: boolean }>("/api/assignments", { method: "DELETE" }),
  validate: () => request<RunAssignmentResult>("/api/assignments/validate"),

  // Reports
  markdownReport: () => request<string>("/api/reports/markdown"),

  // Accompanist -> Jury source boundary
  getAccompanistFinalizationState: () =>
    request<AccompanistFinalizationState>("/api/accompanist/finalization-state"),
  finalizeAccompanist: (expectedSourceRevision: number) =>
    request<unknown>("/api/accompanist/finalize", {
      method: "POST",
      body: JSON.stringify({ expected_source_revision: expectedSourceRevision }),
    }),

  // Jury setup and readiness
  getJuryConfiguration: () => request<JuryConfiguration>("/api/jury/configuration"),
  getJuryPanels: () => request<JuryPanel[]>("/api/jury/panels"),
  createJuryPanel: (panel: JuryPanelInput) =>
    request<JuryPanel>("/api/jury/panels", { method: "POST", body: JSON.stringify(panel) }),
  updateJuryPanel: (panelUuid: string, panel: JuryPanelInput) =>
    request<JuryPanel>(`/api/jury/panels/${encodeURIComponent(panelUuid)}`, {
      method: "PUT",
      body: JSON.stringify(panel),
    }),
  deleteJuryPanel: (panelUuid: string) =>
    request<void>(`/api/jury/panels/${encodeURIComponent(panelUuid)}`, { method: "DELETE" }),
  inspectJuryPanelFile: (file: File, sheetName?: string | null) => {
    const form = new FormData();
    form.append("file", file);
    if (sheetName) form.append("sheet_name", sheetName);
    return request<JuryPanelImportInspection>("/api/jury/panels/import/inspect", { method: "POST", body: form });
  },
  applyJuryPanelImport: (file: File, sheetName: string | null, mapping: JuryPanelImportMapping) => {
    const form = new FormData();
    form.append("file", file);
    if (sheetName) form.append("sheet_name", sheetName);
    form.append("mapping", JSON.stringify(mapping));
    return request<JuryPanelImportResult>("/api/jury/panels/import/apply", { method: "POST", body: form });
  },
  synchronizeJuryRoster: () =>
    request<JuryLessonEntry[]>("/api/jury/roster/synchronize", { method: "POST" }),
  synchronizeJuryWithAccompanist: () =>
    request<JurySynchronizationSummary>("/api/jury/synchronize", { method: "POST" }),
  getJuryEntries: () => request<JuryLessonEntry[]>("/api/jury/entries"),
  assignJuryPanelsByInstrument: () =>
    request<{ assigned_count: number; entries: JuryLessonEntry[] }>("/api/jury/entries/assign-panels-by-instrument", { method: "POST" }),
  updateJuryEntry: (lessonUuid: string, panelUuid: string | null) =>
    request<JuryLessonEntry>(`/api/jury/entries/${encodeURIComponent(lessonUuid)}`, {
      method: "PATCH",
      body: JSON.stringify({ panel_uuid: panelUuid }),
    }),
  updateJuryRequired: (lessonUuid: string, juryRequired: boolean) =>
    request<JuryLessonEntry>(`/api/jury/entries/${encodeURIComponent(lessonUuid)}/jury-required`, {
      method: "PATCH",
      body: JSON.stringify({ jury_required: juryRequired }),
    }),
  getJuryAvailability: (pianistUuid: string, juryDate: string) =>
    request<JuryAvailability>(
      `/api/jury/pianists/${encodeURIComponent(pianistUuid)}/availability/${encodeURIComponent(juryDate)}`
    ),
  saveJuryAvailability: (
    pianistUuid: string,
    juryDate: string,
    availability: JuryAvailabilityInput,
  ) => request<JuryAvailability>(
    `/api/jury/pianists/${encodeURIComponent(pianistUuid)}/availability/${encodeURIComponent(juryDate)}`,
    { method: "PUT", body: JSON.stringify(availability) },
  ),
  clearJuryAvailability: (pianistUuid: string, juryDate: string) =>
    request<JuryAvailabilityClearResult>(
      `/api/jury/pianists/${encodeURIComponent(pianistUuid)}/availability/${encodeURIComponent(juryDate)}`,
      { method: "DELETE" },
  ),
  getJuryReadiness: () => request<JuryReadiness>("/api/jury/readiness"),
  generateJurySchedule: (expectedJuryInputRevision: number) =>
    request<JuryScheduleResult>("/api/jury/generate", {
      method: "POST",
      body: JSON.stringify({ expected_jury_input_revision: expectedJuryInputRevision }),
    }),
  getCurrentJurySchedule: () => request<JuryScheduleResult>("/api/jury/results/current"),
  getJuryScheduleHistory: () => request<JuryScheduleHistoryItem[]>("/api/jury/results/history"),
  getJuryScheduleResult: (resultUuid: string) =>
    request<JuryScheduleResult>(`/api/jury/results/${encodeURIComponent(resultUuid)}`),
};
