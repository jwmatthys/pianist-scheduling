import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import { useRef, useState } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { FindOnPage } from "./FindOnPage";

type FakeHighlight = { ranges: Range[]; priority: number };

// jsdom has no CSS Custom Highlight API; a minimal stand-in lets tests inspect highlights.
function installHighlightApi() {
  const registry = new Map<string, FakeHighlight>();
  class FakeHighlightImpl {
    ranges: Range[];
    priority = 0;
    constructor(...ranges: Range[]) {
      this.ranges = ranges;
    }
  }
  vi.stubGlobal("CSS", { highlights: registry });
  vi.stubGlobal("Highlight", FakeHighlightImpl);
  return registry;
}

const text = (range: Range) => range.toString();

function Harness() {
  const scopeRef = useRef<HTMLElement>(null);
  const [tab, setTab] = useState<"roster" | "schedule">("roster");
  return (
    <div>
      <p>Outside content Smith</p>
      <main ref={scopeRef}>
        <nav>
          <button type="button" aria-current={tab === "roster" ? "page" : undefined} onClick={() => setTab("roster")}>Roster</button>
          <button type="button" aria-current={tab === "schedule" ? "page" : undefined} onClick={() => setTab("schedule")}>Schedule</button>
        </nav>
        {tab === "roster" ? (
          <ul>
            <li>John Smith</li>
            <li>SMITHSON, Alex</li>
            <li>Taylor Jones</li>
            <li hidden>Hidden Smith</li>
            <li>smith again</li>
          </ul>
        ) : (
          <p>Schedule for Smith</p>
        )}
      </main>
      <FindOnPage scopeRef={scopeRef} />
    </div>
  );
}

function openFind() {
  fireEvent.keyDown(document, { key: "f", ctrlKey: true });
  return screen.getByRole("textbox", { name: "Find on page" }) as HTMLInputElement;
}

function ControlsHarness() {
  const scopeRef = useRef<HTMLElement>(null);
  return (
    <div>
      <main ref={scopeRef}>
        <label>Name <input defaultValue="Alex Smith" /></label>
        <label>Notes <textarea defaultValue="Likes smith chords" /></label>
        <input type="password" defaultValue="smith" aria-label="secret" />
        <input hidden defaultValue="smith" aria-label="hidden field" />
        <input defaultValue="Taylor" aria-label="other" />
        <p>Smith in text</p>
      </main>
      <FindOnPage scopeRef={scopeRef} />
    </div>
  );
}

describe("FindOnPage", () => {
  let registry: Map<string, FakeHighlight>;
  const scrollIntoView = vi.fn();

  beforeEach(() => {
    registry = installHighlightApi();
    Element.prototype.scrollIntoView = scrollIntoView;
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("opens with Ctrl+F, focuses the field, and closes with Escape", () => {
    render(<Harness />);
    expect(screen.queryByRole("search")).toBeNull();

    const input = openFind();
    expect(document.activeElement).toBe(input);

    fireEvent.keyDown(input, { key: "Escape" });
    expect(screen.queryByRole("search")).toBeNull();
  });

  it("matches case-insensitive substrings and highlights them without hiding any content", () => {
    render(<Harness />);
    const input = openFind();
    fireEvent.change(input, { target: { value: "smith" } });

    // John Smith, SMITHSON, smith again; hidden and out-of-scope text is ignored.
    expect(screen.getByRole("status").textContent).toBe("1 of 3");
    const current = registry.get("find-current")!.ranges.map(text);
    const others = registry.get("find-match")!.ranges.map(text);
    expect(current).toEqual(["Smith"]);
    expect(others).toEqual(["SMITH", "smith"]);
    expect(registry.get("find-current")!.priority).toBe(1);

    expect(screen.getByText("Taylor Jones")).toBeTruthy();
    expect(screen.getByText("John Smith")).toBeTruthy();
  });

  it("shows no matches and no highlights for text that is absent", () => {
    render(<Harness />);
    fireEvent.change(openFind(), { target: { value: "zzz" } });
    expect(screen.getByRole("status").textContent).toBe("No matches");
    expect(registry.has("find-match")).toBe(false);
    expect(registry.has("find-current")).toBe(false);
  });

  it("moves between matches with arrow keys, wrapping around, and scrolls the current match into view", () => {
    render(<Harness />);
    const input = openFind();
    fireEvent.change(input, { target: { value: "smith" } });
    scrollIntoView.mockClear();

    fireEvent.keyDown(input, { key: "ArrowDown" });
    expect(screen.getByRole("status").textContent).toBe("2 of 3");
    expect(registry.get("find-current")!.ranges.map(text)).toEqual(["SMITH"]);
    expect(scrollIntoView).toHaveBeenCalledTimes(1);

    fireEvent.keyDown(input, { key: "ArrowRight" });
    expect(screen.getByRole("status").textContent).toBe("3 of 3");

    fireEvent.keyDown(input, { key: "ArrowRight" });
    expect(screen.getByRole("status").textContent).toBe("1 of 3");

    fireEvent.keyDown(input, { key: "ArrowUp" });
    expect(screen.getByRole("status").textContent).toBe("3 of 3");

    fireEvent.keyDown(input, { key: "ArrowLeft" });
    expect(screen.getByRole("status").textContent).toBe("2 of 3");
  });

  it("supports the previous and next buttons", () => {
    render(<Harness />);
    fireEvent.change(openFind(), { target: { value: "smith" } });
    fireEvent.click(screen.getByRole("button", { name: "Next match" }));
    expect(screen.getByRole("status").textContent).toBe("2 of 3");
    fireEvent.click(screen.getByRole("button", { name: "Previous match" }));
    expect(screen.getByRole("status").textContent).toBe("1 of 3");
  });

  it("clears the search and highlights when closed, and starts empty when reopened", () => {
    render(<Harness />);
    const input = openFind();
    fireEvent.change(input, { target: { value: "smith" } });
    expect(registry.has("find-match")).toBe(true);

    fireEvent.keyDown(input, { key: "Escape" });
    expect(registry.has("find-match")).toBe(false);
    expect(registry.has("find-current")).toBe(false);

    expect(openFind().value).toBe("");
  });

  it("closes the search when the visible tab changes", async () => {
    render(<Harness />);
    fireEvent.change(openFind(), { target: { value: "smith" } });
    expect(screen.getByRole("search")).toBeTruthy();

    fireEvent.click(screen.getByRole("button", { name: "Schedule" }));

    await waitFor(() => expect(screen.queryByRole("search")).toBeNull());
    expect(registry.has("find-match")).toBe(false);
    expect(screen.getByText("Schedule for Smith")).toBeTruthy();
  });

  it("returns focus to the previously focused element on Escape", () => {
    render(<Harness />);
    const rosterButton = screen.getByRole("button", { name: "Roster" });
    rosterButton.focus();
    const input = openFind();
    fireEvent.keyDown(input, { key: "Escape" });
    expect(document.activeElement).toBe(rosterButton);
  });

  it("also matches text boxes and marks them, skipping passwords and hidden fields", () => {
    render(<ControlsHarness />);
    const input = openFind();
    fireEvent.change(input, { target: { value: "smith" } });
    expect(screen.getByRole("status").textContent).toBe("1 of 3");

    const marked = Array.from(document.querySelectorAll("[data-find-match]"));
    expect(marked.map((element) => element.tagName)).toEqual(["INPUT", "TEXTAREA"]);

    fireEvent.change(input, { target: { value: "" } });
    expect(document.querySelectorAll("[data-find-match]").length).toBe(0);
  });
});
