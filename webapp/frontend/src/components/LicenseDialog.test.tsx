// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
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
};

const licensed = {
  status: "licensed" as const,
  payload: {
    product: "Scheduling Maestro",
    licensed_to: "Synthetic University",
    license_type: "PERPETUAL",
    issued: "2026-11-23",
    license_version: 1 as const,
  },
  activated_at: "2026-11-23T12:00:00Z",
  error: null,
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
    render(<LicenseDialog />);

    fireEvent.click(screen.getByRole("button", { name: "License" }));
    expect(await screen.findByText("Not activated")).not.toBeNull();
    fireEvent.change(screen.getByLabelText("Enter License Key"), { target: { value: "SM1-signed-key" } });
    fireEvent.click(screen.getByRole("button", { name: "Activate License Key" }));

    expect(await screen.findByText("Synthetic University")).not.toBeNull();
    expect(screen.getByText("PERPETUAL")).not.toBeNull();
    expect(api.activateLicenseKey).toHaveBeenCalledWith("SM1-signed-key");
  });

  it("imports a selected license file", async () => {
    const file = new File(["{}"], "scheduling-maestro.license", { type: "application/json" });
    vi.mocked(api.getLicenseStatus).mockResolvedValue(unlicensed);
    vi.mocked(chooseLocalFile).mockResolvedValue(file);
    vi.mocked(api.activateLicenseFile).mockResolvedValue(licensed);
    render(<LicenseDialog />);

    fireEvent.click(screen.getByRole("button", { name: "License" }));
    await screen.findByText("Not activated");
    fireEvent.click(screen.getByRole("button", { name: "Import License File" }));

    await waitFor(() => expect(api.activateLicenseFile).toHaveBeenCalledWith(file));
    expect(await screen.findByText("Synthetic University")).not.toBeNull();
    expect(chooseLocalFile).toHaveBeenCalledWith("Scheduling Maestro License", ["license"]);
  });
});
