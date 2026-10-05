import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "./lib/api";
import type { SchedulingSession } from "./lib/types";
import App from "./App";

vi.mock("./lib/api", () => ({
  api: {
    getSession: vi.fn(),
    createSession: vi.fn(),
    updateSession: vi.fn(),
  },
}));

vi.mock("./lib/platform", () => ({
  chooseSessionArchive: vi.fn(),
  saveSessionArchive: vi.fn(),
}));

const session: SchedulingSession = {
  session_uuid: "session-uuid",
  institution_name: "Example University",
  program_name: "Music",
  term_label: "Fall 2027",
  year: 2026,
  start_date: null,
  end_date: null,
  created_at: "2026-01-01T00:00:00",
  modified_at: "2026-01-01T00:00:00",
};

const dialogDescriptors = {
  showModal: Object.getOwnPropertyDescriptor(HTMLDialogElement.prototype, "showModal"),
  close: Object.getOwnPropertyDescriptor(HTMLDialogElement.prototype, "close"),
};

beforeEach(() => {
  Object.defineProperty(HTMLDialogElement.prototype, "showModal", {
    configurable: true,
    value(this: HTMLDialogElement) {
      this.setAttribute("open", "");
    },
  });
  Object.defineProperty(HTMLDialogElement.prototype, "close", {
    configurable: true,
    value(this: HTMLDialogElement) {
      this.removeAttribute("open");
    },
  });
  vi.mocked(api.getSession).mockResolvedValue(session);
  vi.mocked(api.createSession).mockResolvedValue(session);
  vi.mocked(api.updateSession).mockResolvedValue(session);
  vi.spyOn(window, "confirm").mockReturnValue(true);
});

afterEach(() => {
  vi.restoreAllMocks();
  for (const [method, descriptor] of Object.entries(dialogDescriptors)) {
    if (descriptor) Object.defineProperty(HTMLDialogElement.prototype, method, descriptor);
    else Reflect.deleteProperty(HTMLDialogElement.prototype, method);
  }
});

describe("Scheduling Session form", () => {
  it("uses the term label as the only term detail when creating a session", async () => {
    render(<App />);
    await screen.findByText("Example University");
    expect(screen.getByText("Music · Fall 2027")).toBeTruthy();
    fireEvent.click(screen.getByRole("button", { name: "New Session" }));

    const dialog = screen.getByRole("dialog");
    expect(dialog.querySelector('[name="year"]')).toBeNull();
    expect(dialog.querySelector('[name="start_date"]')).toBeNull();
    expect(dialog.querySelector('[name="end_date"]')).toBeNull();

    fireEvent.change(screen.getByLabelText("Institution"), { target: { value: "New University" } });
    fireEvent.change(screen.getByLabelText("Program"), { target: { value: "Strings" } });
    fireEvent.change(screen.getByLabelText("Term label"), { target: { value: "Spring 2028" } });
    fireEvent.click(screen.getByRole("button", { name: "Create Session" }));

    await waitFor(() => {
      expect(api.createSession).toHaveBeenCalledWith({
        institution_name: "New University",
        program_name: "Strings",
        term_label: "Spring 2028",
      });
    });
  });

  it("omits the separate year and date fields when editing a session", async () => {
    render(<App />);
    await screen.findByText("Example University");
    fireEvent.click(screen.getByRole("button", { name: "Edit Session" }));

    const dialog = screen.getByRole("dialog");
    expect(dialog.querySelector('[name="year"]')).toBeNull();
    expect(dialog.querySelector('[name="start_date"]')).toBeNull();
    expect(dialog.querySelector('[name="end_date"]')).toBeNull();

    fireEvent.click(screen.getByRole("button", { name: "Update Session" }));

    await waitFor(() => {
      expect(api.updateSession).toHaveBeenCalledWith({
        institution_name: "Example University",
        program_name: "Music",
        term_label: "Fall 2027",
      });
    });
  });
});
