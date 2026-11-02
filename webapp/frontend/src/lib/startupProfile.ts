// Startup milestones. Console logging is opt-in (dev builds or ?startupProfile=1), but
// milestones are always recorded and forwarded to the native shell, which drives the
// startup splash screen from them. Only fixed phase labels and elapsed milliseconds are
// used; never pass user or session data. To remove the timing logs, delete this module's
// logging; the splash depends on mark(...) and setStartupSink(...).

export type StartupSink = (label: string) => void;

const marks = new Map<string, number>();
let sink: StartupSink | undefined;

export function startupProfilingEnabled(): boolean {
  if (typeof window === "undefined") return false;
  return (import.meta.env.DEV && import.meta.env.MODE !== "test")
    || new URLSearchParams(window.location.search).get("startupProfile") === "1";
}

const enabled = startupProfilingEnabled();

export function mark(label: string): void {
  if (marks.has(label)) return;
  const elapsed = performance.now();
  marks.set(label, elapsed);
  if (enabled) console.info(`[STARTUP frontend +${elapsed.toFixed(0)}ms] ${label}`);
  sink?.(label);
}

// The platform layer forwards milestones to the native shell, which owns the
// end-to-end clock. Marks recorded before the sink exists are replayed.
export function setStartupSink(next: StartupSink): void {
  sink = next;
  for (const label of marks.keys()) next(label);
}

export function span(begin: string, end: string): number | undefined {
  const start = marks.get(begin);
  const finish = marks.get(end);
  return start === undefined || finish === undefined ? undefined : finish - start;
}

// Prints measured spans only; a span is omitted when either end was not recorded.
export function logSummary(): void {
  if (!enabled) return;
  const total = marks.get("First usable UI rendered");
  if (total === undefined) return;
  console.info(`[STARTUP frontend] Time from page start to first usable UI: ${total.toFixed(0)} ms`);
  const spans: [string, string, string][] = [
    ["Frontend initialization (React begin to session request)", "React bootstrap begin", "Session load begin"],
    ["API base resolution", "API base resolution begin", "API base resolution complete"],
    ["Session load", "Session load begin", "Session load complete"],
    ["Dashboard render (session loaded to first usable UI)", "Session load complete", "First usable UI rendered"],
  ];
  for (const [name, begin, end] of spans) {
    const value = span(begin, end);
    if (value !== undefined) console.info(`[STARTUP frontend]   ${name}: ${value.toFixed(0)} ms`);
  }
}

mark("React bootstrap begin");

export function resetStartupProfileForTests(): void {
  marks.clear();
  sink = undefined;
}
