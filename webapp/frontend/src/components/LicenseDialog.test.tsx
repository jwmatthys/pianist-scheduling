// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { LicenseDialog } from "./LicenseDialog";
import { api } from "../lib/api";
import { chooseLocalFile } from "../lib/platform";

vi.mock("../lib/api", () => ({
  api: {
    getLicenseStatus: vi.fn(),
    activateLicenseFile: vi.fn(),
    activateLicenseKey: vi.fn(),
  },
}));

vi.mock("../lib/platform", () => ({
  chooseLocalFile: vi.fn(),
}));

const unlicensed = {
  status: "unlicensed" as const,
  payload: null,
  activated_at: null,
  error: null,
  trial_seconds_remaining: 1800,
  signing_key_id: null,
};

const licensed = {
  status: "licensed" as const,
  payload: {
    product: "Scheduling Maestro",
    licensed_to: "Synthetic University",
    license_type: "PERPETUAL",
    issued: "2026-11-23",
    license_version: 1 as const,
    key_id: "v1",
  },
  activated_at: "2026-11-23T12:00:00Z",
  error: null,
  trial_seconds_remaining: null,
  signing_key_id: "v1",
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
});

afterEach(() => {
  cleanup();
  vi.useRealTimers();
  vi.mocked(api.getLicenseStatus).mockReset();
  vi.mocked(api.activateLicenseFile).mockReset();
  vi.mocked(api.activateLicenseKey).mockReset();
  vi.mocked(chooseLocalFile).mockReset();
  delete (HTMLDialogElement.prototype as HTMLDialogElement & { showModal?: () => void }).showModal;
  delete (HTMLDialogElement.prototype as HTMLDialogElement & { close?: () => void }).close;
});

describe("LicenseDialog", () => {
  it("shows status and activates a license key", async () => {
    vi.mocked(api.getLicenseStatus).mockResolvedValue(unlicensed);
    vi.mocked(api.activateLicenseKey).mockResolvedValue(licensed);
    render(<LicenseDialog licenseStatus={unlicensed} remaining={1800} trialError={null} onLicenseStatus={vi.fn()} />);

    fireEvent.click(screen.getByRole("button", { name: "License" }));
    expect(await screen.findByText("Not activated")).not.toBeNull();
    fireEvent.change(screen.getByLabelText("Enter License Key"), { target: { value: "SM1-signed-key" } });
    fireEvent.click(screen.getByRole("button", { name: "Activate License Key" }));

    expect(await screen.findByText("Synthetic University")).not.toBeNull();
    expect(screen.getByText("PERPETUAL")).not.toBeNull();
    expect(screen.getByText("v1")).not.toBeNull();
    expect(api.activateLicenseKey).toHaveBeenCalledWith("SM1-signed-key");
  });

  it("imports a selected license file", async () => {
    const file = new File(["{}"], "scheduling-maestro.license", { type: "application/json" });
    vi.mocked(api.getLicenseStatus).mockResolvedValue(unlicensed);
    vi.mocked(chooseLocalFile).mockResolvedValue(file);
    vi.mocked(api.activateLicenseFile).mockResolvedValue(licensed);
    render(<LicenseDialog licenseStatus={unlicensed} remaining={1800} trialError={null} onLicenseStatus={vi.fn()} />);

    fireEvent.click(screen.getByRole("button", { name: "License" }));
    await screen.findByText("Not activated");
    fireEvent.click(screen.getByRole("button", { name: "Import License File" }));

    await waitFor(() => expect(api.activateLicenseFile).toHaveBeenCalledWith(file));
    expect(await screen.findByText("Synthetic University")).not.toBeNull();
    expect(chooseLocalFile).toHaveBeenCalledWith("Scheduling Maestro License", ["license", "txt"]);
  });

  it.each(["file", "key"] as const)("replaces existing license metadata using a %s", async (method) => {
    const replacement = {
      ...licensed,
      payload: { ...licensed.payload, licensed_to: "Replacement Conservatory", issued: "2026-11-24" },
      activated_at: "2026-11-24T12:00:00Z",
    };
    const file = new File(["replacement"], "scheduling-maestro.license");
    const onLicenseStatus = vi.fn();
    vi.mocked(api.getLicenseStatus).mockResolvedValueOnce(licensed).mockResolvedValue(replacement);
    vi.mocked(chooseLocalFile).mockResolvedValue(file);
    vi.mocked(api.activateLicenseFile).mockResolvedValue(replacement);
    vi.mocked(api.activateLicenseKey).mockResolvedValue(replacement);
    render(<LicenseDialog licenseStatus={licensed} remaining={null} trialError={null} onLicenseStatus={onLicenseStatus} />);

    fireEvent.click(screen.getByRole("button", { name: "License" }));
    await screen.findByText("Synthetic University");
    if (method === "file") {
      fireEvent.click(screen.getByRole("button", { name: "Import License File" }));
    } else {
      fireEvent.change(screen.getByLabelText("Enter License Key"), { target: { value: "SM1-replacement" } });
      fireEvent.click(screen.getByRole("button", { name: "Activate License Key" }));
    }

    expect(await screen.findByText("Replacement Conservatory")).not.toBeNull();
    expect(screen.getByText("2026-11-24")).not.toBeNull();
    expect(screen.queryByText("Synthetic University")).toBeNull();
    expect(onLicenseStatus).toHaveBeenCalledWith(replacement);
    fireEvent.click(screen.getByRole("button", { name: "Close" }));
    fireEvent.click(screen.getByRole("button", { name: "License" }));
    expect(await screen.findByText("Replacement Conservatory")).not.toBeNull();
  });

  it("does not let an older status request overwrite a replacement", async () => {
    const replacement = {
      ...licensed,
      payload: { ...licensed.payload, licensed_to: "Replacement Conservatory" },
    };
    let finishActivation!: (status: typeof licensed) => void;
    let finishStatus!: (status: typeof licensed) => void;
    vi.mocked(api.getLicenseStatus)
      .mockResolvedValueOnce(licensed)
      .mockImplementationOnce(() => new Promise((resolve) => { finishStatus = resolve; }));
    vi.mocked(api.activateLicenseKey)
      .mockImplementationOnce(() => new Promise((resolve) => { finishActivation = resolve; }));
    render(<LicenseDialog licenseStatus={licensed} remaining={null} trialError={null} onLicenseStatus={vi.fn()} />);

    fireEvent.click(screen.getByRole("button", { name: "License" }));
    await screen.findByText("Synthetic University");
    fireEvent.change(screen.getByLabelText("Enter License Key"), { target: { value: "SM1-replacement" } });
    fireEvent.click(screen.getByRole("button", { name: "Activate License Key" }));
    fireEvent.click(screen.getByRole("button", { name: "Close" }));
    fireEvent.click(screen.getByRole("button", { name: "License" }));
    await act(async () => { finishActivation(replacement); });
    expect(screen.getByText("Replacement Conservatory")).not.toBeNull();
    await act(async () => { finishStatus(licensed); });
    expect(screen.getByText("Replacement Conservatory")).not.toBeNull();
    expect(screen.queryByText("Synthetic University")).toBeNull();
    expect(screen.getByRole("button", { name: "Import License File" }).hasAttribute("disabled")).toBe(false);
  });

  it("shows remaining trial time only when unlicensed", () => {
    const onLicenseStatus = vi.fn();
    const { rerender } = render(
      <LicenseDialog licenseStatus={unlicensed} remaining={2} trialError={null} onLicenseStatus={onLicenseStatus} />,
    );
    expect(screen.getByText("Trial: 0:02 remaining")).not.toBeNull();
    rerender(<LicenseDialog licenseStatus={unlicensed} remaining={0} trialError={null} onLicenseStatus={onLicenseStatus} />);
    expect(screen.getByText("Trial: 0:00 remaining")).not.toBeNull();
    rerender(<LicenseDialog licenseStatus={licensed} remaining={null} trialError={null} onLicenseStatus={onLicenseStatus} />);
    expect(screen.queryByText("Trial: 0:00 remaining")).toBeNull();
  });
});
