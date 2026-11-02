// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { DashboardPage } from "./DashboardPage";
import { ThemeProvider } from "../theme/ThemeContext";

afterEach(cleanup);

function renderDashboard(onOpenModule = vi.fn()) {
  const view = render(
    <ThemeProvider>
      <DashboardPage
        session={null}
        busy={false}
        onEditSession={vi.fn()}
        onNewSession={vi.fn()}
        onOpenSession={vi.fn()}
        onExportSession={vi.fn()}
        onOpenModule={onOpenModule}
      />
    </ThemeProvider>,
  );
  return { ...view, onOpenModule };
}

describe("DashboardPage module layout", () => {
  it("separates independent modules from the Accompanist to Jury chain", () => {
    renderDashboard();
    const independent = screen.getByRole("navigation", { name: "Independent scheduling modules" });
    expect(Array.from(independent.querySelectorAll(".module-card-title")).map((n) => n.textContent)).toEqual([
      "Lesson Scheduling",
      "Tutor Scheduling",
      "Clinical Placements",
    ]);
    const chain = screen.getByRole("navigation", { name: "Accompanist to jury scheduling workflow" });
    expect(Array.from(chain.querySelectorAll(".module-card-title")).map((n) => n.textContent)).toEqual([
      "Accompanist Scheduling",
      "Performance Jury Scheduling",
    ]);
  });

  it("renders a non-focusable, informational dependency connector", () => {
    const { container } = renderDashboard();
    const connector = container.querySelector(".module-dependency") as HTMLElement;
    expect(connector.textContent).toBe("Uses finalized accompanist assignments");
    expect(connector.querySelector("svg")?.getAttribute("aria-hidden")).toBe("true");
    expect(connector.querySelectorAll("button, a, [tabindex]").length).toBe(0);
  });

  it("keeps module navigation working", () => {
    const { onOpenModule } = renderDashboard();
    for (const [name, key] of [
      ["Lesson Scheduling", "lesson"],
      ["Accompanist Scheduling", "accompanist"],
      ["Performance Jury Scheduling", "juries"],
    ] as const) {
      fireEvent.click(screen.getByText(name).closest("button")!);
      expect(onOpenModule).toHaveBeenLastCalledWith(key);
    }
  });

  it("still offers the theme chooser", () => {
    renderDashboard();
    expect(screen.getByRole("button", { name: "Choose theme" })).not.toBeNull();
  });
});
