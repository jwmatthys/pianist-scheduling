import { act, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "./lib/api";
import type { SchedulingSession } from "./lib/types";
import { chooseLocalFile, chooseSessionArchive } from "./lib/platform";
import App from "./App";

vi.mock("./lib/api", () => ({
  api: {
    getSession: vi.fn(),
    trialHeartbeat: vi.fn(),
    getLicenseStatus: vi.fn(),
    activateLicenseFile: vi.fn(),
    activateLicenseKey: vi.fn(),
    createSession: vi.fn(),
    updateSession: vi.fn(),
    inspectSessionArchive: vi.fn(),
    restoreSession: vi.fn(),
  },
}));

vi.mock("./pages/TutorModule", () => ({
  TutorModule: () => <div>Tutor module content<input aria-label="Unsaved draft" /></div>,
}));

vi.mock("./lib/platform", () => ({
  chooseLocalFile: vi.fn(),
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
  vi.mocked(api.getSession).mockClear();
  vi.mocked(api.trialHeartbeat).mockReset();
  vi.mocked(api.activateLicenseFile).mockReset();
  vi.mocked(api.activateLicenseKey).mockReset();
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
  vi.mocked(api.trialHeartbeat).mockResolvedValue({
    status: "unlicensed", payload: null, activated_at: null, error: null, trial_seconds_remaining: 1800, signing_key_id: null,
  });
  vi.mocked(api.createSession).mockResolvedValue(session);
  vi.mocked(api.updateSession).mockResolvedValue(session);
  vi.spyOn(window, "confirm").mockReturnValue(true);
});

const activeLicense = {
  status: "licensed" as const,
  payload: {
    product: "Scheduling Maestro", licensed_to: "Synthetic University",
    license_type: "PERPETUAL", issued: "2026-11-23", license_version: 1 as const, key_id: "v1",
  },
  activated_at: "2026-11-23T12:00:00Z",
  error: null,
  trial_seconds_remaining: null,
  signing_key_id: "v1",
};

it("locks an open module at zero and retains its draft until key activation", async () => {
  vi.useFakeTimers();
  const clock = vi.spyOn(performance, "now").mockReturnValue(0);
  vi.mocked(api.trialHeartbeat).mockResolvedValue({
    status: "unlicensed", payload: null, activated_at: null, error: null, trial_seconds_remaining: 1, signing_key_id: null,
  });
  vi.mocked(api.activateLicenseKey).mockRejectedValueOnce(new Error("Invalid key")).mockResolvedValue(activeLicense);
  const { container } = render(<App />);
  await act(async () => {});
  fireEvent.click(screen.getByText("Tutor Scheduling").closest("button")!);
  fireEvent.change(screen.getByLabelText("Unsaved draft"), { target: { value: "draft" } });
  clock.mockReturnValue(1000);
  await act(async () => { await vi.advanceTimersByTimeAsync(1000); });

  const dialog = screen.getByRole("dialog", { name: "Scheduling Maestro trial expired. Upload or enter your license key to activate." });
  expect(container.querySelector(".app-shell")?.hasAttribute("inert")).toBe(true);
  expect(dialog.querySelector(".dialog-close")).toBeNull();
  const cancel = new Event("cancel", { bubbles: true, cancelable: true });
  fireEvent(dialog, cancel);
  expect(cancel.defaultPrevented).toBe(true);
  clock.mockReturnValue(5000);
  await act(async () => { await vi.advanceTimersByTimeAsync(4000); });
  expect(screen.getByRole("dialog", { name: "Scheduling Maestro trial expired. Upload or enter your license key to activate." })).toBeTruthy();
  fireEvent.change(within(dialog).getByLabelText("Enter License Key"), { target: { value: "bad" } });
  await act(async () => { fireEvent.click(within(dialog).getByRole("button", { name: "Activate License Key" })); });
  expect(within(dialog).getByRole("alert").textContent).toBe("Invalid key");
  fireEvent.change(within(dialog).getByLabelText("Enter License Key"), { target: { value: "valid" } });
  await act(async () => { fireEvent.click(within(dialog).getByRole("button", { name: "Activate License Key" })); });
  expect(screen.queryByRole("dialog")).toBeNull();
  expect(container.querySelector(".app-shell")?.hasAttribute("inert")).toBe(false);
  expect((screen.getByLabelText("Unsaved draft") as HTMLInputElement).value).toBe("draft");
});

it("starts locked without loading session data and unlocks after file activation", async () => {
  vi.mocked(api.trialHeartbeat).mockResolvedValue({
    status: "unlicensed", payload: null, activated_at: null, error: null, trial_seconds_remaining: 0, signing_key_id: null,
  });
  const file = new File(["synthetic"], "scheduling-maestro.license");
  vi.mocked(chooseSessionArchive).mockReset();
  vi.mocked(chooseLocalFile).mockResolvedValue(file);
  vi.mocked(api.activateLicenseFile).mockResolvedValue(activeLicense);
  render(<App />);
  const dialog = await screen.findByRole("dialog", { name: "Scheduling Maestro trial expired. Upload or enter your license key to activate." });
  expect(api.getSession).not.toHaveBeenCalled();
  fireEvent.click(within(dialog).getByRole("button", { name: "Import License File" }));
  expect(await screen.findByText("Example University")).toBeTruthy();
  expect(api.activateLicenseFile).toHaveBeenCalledWith(file);
});

it("does not load the application when the trial status cannot be checked", async () => {
  vi.mocked(api.trialHeartbeat).mockResolvedValue({
    status: "unlicensed", payload: null, activated_at: null, error: null, trial_seconds_remaining: null, signing_key_id: null,
  });
  render(<App />);
  expect(await screen.findByText(/Could not check license status/)).toBeTruthy();
  expect(api.getSession).not.toHaveBeenCalled();
  expect(screen.queryByRole("button", { name: "New Session" })).toBeNull();
});

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
  for (const [method, descriptor] of Object.entries(dialogDescriptors)) {
    if (descriptor) Object.defineProperty(HTMLDialogElement.prototype, method, descriptor);
    else Reflect.deleteProperty(HTMLDialogElement.prototype, method);
  }
});

it("keeps tracking trial time while a module is open", async () => {
  vi.useFakeTimers();
  vi.mocked(api.trialHeartbeat).mockClear();
  await act(async () => { render(<App />); });
  fireEvent.click(screen.getByText("Tutor Scheduling").closest("button")!);
  expect(screen.getByText("Tutor module content")).toBeTruthy();
  await act(async () => { await vi.advanceTimersByTimeAsync(5000); });
  expect(api.trialHeartbeat).toHaveBeenCalledTimes(2);
  await act(async () => { fireEvent.click(screen.getByRole("button", { name: "Home" })); });
  expect(screen.getByText("Trial: 30:00 remaining")).toBeTruthy();
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

describe("Opening an older session archive", () => {
  const archive = new Uint8Array([1, 2, 3]);

  beforeEach(() => {
    vi.mocked(chooseSessionArchive).mockResolvedValue(archive);
    vi.mocked(api.restoreSession).mockResolvedValue(session);
  });

  async function openArchive() {
    render(<App />);
    await screen.findByText("Example University");
    fireEvent.click(screen.getByRole("button", { name: "Open Session" }));
  }

  it("warns without restoring, and Cancel changes nothing", async () => {
    vi.mocked(api.inspectSessionArchive).mockResolvedValue({
      archive_exported_utc: "2026-11-01T10:00:00Z",
      current_activity_utc: "2026-11-02T21:14:33Z",
      is_older: true,
    });
    await openArchive();

    const dialog = await screen.findByRole("dialog", { name: "Open Older Session Archive?" });
    expect(dialog.textContent).toContain("saved earlier than the session currently open");
    expect(dialog.textContent).toContain("Current session:");
    expect(dialog.textContent).toContain("Archive being opened:");
    fireEvent.click(within(dialog).getByRole("button", { name: "Cancel" }));

    expect(screen.queryByRole("dialog", { name: "Open Older Session Archive?" })).toBeNull();
    expect(api.restoreSession).not.toHaveBeenCalled();
  });

  it("Continue Opening proceeds with the existing restore workflow", async () => {
    vi.mocked(api.inspectSessionArchive).mockResolvedValue({
      archive_exported_utc: "2026-11-01T10:00:00Z",
      current_activity_utc: "2026-11-02T21:14:33Z",
      is_older: true,
    });
    await openArchive();

    const dialog = await screen.findByRole("dialog", { name: "Open Older Session Archive?" });
    fireEvent.click(within(dialog).getByRole("button", { name: "Continue Opening" }));

    await waitFor(() => expect(api.restoreSession).toHaveBeenCalledWith(archive));
    expect(window.confirm).toHaveBeenCalled();
  });

  it("opens a non-older archive without the warning", async () => {
    vi.mocked(api.inspectSessionArchive).mockResolvedValue({
      archive_exported_utc: null,
      current_activity_utc: null,
      is_older: false,
    });
    await openArchive();

    await waitFor(() => expect(api.restoreSession).toHaveBeenCalledWith(archive));
    expect(screen.queryByRole("dialog", { name: "Open Older Session Archive?" })).toBeNull();
  });
});
