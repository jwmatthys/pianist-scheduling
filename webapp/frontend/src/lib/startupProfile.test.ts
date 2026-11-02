import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

async function loadProfiler(search: string) {
  window.history.replaceState(null, "", `/${search}`);
  vi.resetModules();
  return import("./startupProfile");
}

describe("startupProfile", () => {
  let info: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    info = vi.spyOn(console, "info").mockImplementation(() => undefined);
  });

  afterEach(() => {
    info.mockRestore();
    window.history.replaceState(null, "", "/");
  });

  it("is silent by default and logs nothing to the console", async () => {
    const profiler = await loadProfiler("");
    profiler.mark("Session load begin");
    profiler.logSummary();
    expect(info).not.toHaveBeenCalled();
  });

  it("still forwards each milestone once to the sink when logging is disabled", async () => {
    const profiler = await loadProfiler("");
    const sink = vi.fn();
    profiler.mark("React bootstrap begin");
    profiler.setStartupSink(sink);
    profiler.mark("Session load begin");
    profiler.mark("Session load begin");
    expect(sink.mock.calls.map((call) => call[0])).toEqual(["React bootstrap begin", "Session load begin"]);
    expect(info).not.toHaveBeenCalled();
  });

  it("logs each label once with the [STARTUP] tag when enabled", async () => {
    const profiler = await loadProfiler("?startupProfile=1");
    profiler.mark("Session load begin");
    profiler.mark("Session load begin");
    const lines = info.mock.calls.map((call: unknown[]) => String(call[0]));
    expect(lines.filter((line: string) => line.includes("Session load begin"))).toHaveLength(1);
    expect(lines.every((line: string) => line.startsWith("[STARTUP"))).toBe(true);
  });

  it("replays earlier marks to a late sink and forwards later ones", async () => {
    const profiler = await loadProfiler("?startupProfile=1");
    const sink = vi.fn();
    profiler.mark("Platform initialized");
    profiler.setStartupSink(sink);
    profiler.mark("Session load begin");
    expect(sink.mock.calls.map((call) => call[0])).toEqual([
      "React bootstrap begin",
      "Platform initialized",
      "Session load begin",
    ]);
  });

  it("summarizes only measured spans", async () => {
    const profiler = await loadProfiler("?startupProfile=1");
    profiler.mark("Session load begin");
    profiler.mark("Session load complete");
    profiler.mark("First usable UI rendered");
    info.mockClear();
    profiler.logSummary();
    const text = info.mock.calls.map((call: unknown[]) => String(call[0])).join("\n");
    expect(text).toContain("first usable UI");
    expect(text).toContain("Session load:");
    expect(text).not.toContain("API base resolution:");
  });
});
