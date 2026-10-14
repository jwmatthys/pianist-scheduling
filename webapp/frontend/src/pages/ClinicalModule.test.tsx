import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, it } from "vitest";
import { ClinicalModule } from "./ClinicalModule";

describe("ClinicalModule", () => {
  it("shows the four tabs in order and starts on Clinical Sites", () => {
    render(<ClinicalModule />);
    const labels = screen
      .getAllByRole("button")
      .map((button) => button.textContent);
    expect(labels).toEqual(["Clinical Sites", "Student Roster", "Schedule", "Reports"]);
    expect(screen.getByRole("heading", { name: "Clinical Sites" })).toBeTruthy();
  });

  it("shows not-implemented placeholders for Schedule and Reports", async () => {
    const user = userEvent.setup();
    render(<ClinicalModule />);

    await user.click(screen.getByRole("button", { name: "Schedule" }));
    expect(screen.getByText(/scheduling functionality is not implemented/i)).toBeTruthy();

    await user.click(screen.getByRole("button", { name: "Reports" }));
    expect(screen.getByText(/reports and exports are not implemented/i)).toBeTruthy();
  });
});
