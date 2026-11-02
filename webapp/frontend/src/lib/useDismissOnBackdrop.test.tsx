// @vitest-environment jsdom
import { fireEvent, render, screen } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { useDismissOnBackdrop } from "./useDismissOnBackdrop";

function Harness({ onClose, withClose = true }: { onClose: () => void; withClose?: boolean }) {
  useDismissOnBackdrop();
  return (
    <dialog open data-testid="dialog">
      {withClose && <button type="button" className="dialog-close" onClick={onClose}>x</button>}
      <p>Inside</p>
    </dialog>
  );
}

function backdropClick(dialog: HTMLElement) {
  dialog.getBoundingClientRect = () => ({ left: 100, right: 200, top: 100, bottom: 200, width: 100, height: 100, x: 100, y: 100, toJSON: () => ({}) });
  fireEvent.mouseDown(dialog, { clientX: 10, clientY: 10 });
  fireEvent.click(dialog, { clientX: 10, clientY: 10 });
}

describe("useDismissOnBackdrop", () => {
  it("closes through the X button when the backdrop is clicked", () => {
    const onClose = vi.fn();
    render(<Harness onClose={onClose} />);
    backdropClick(screen.getByTestId("dialog"));
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it("ignores clicks inside the dialog and its content", () => {
    const onClose = vi.fn();
    render(<Harness onClose={onClose} />);
    const dialog = screen.getByTestId("dialog");
    dialog.getBoundingClientRect = () => ({ left: 100, right: 200, top: 100, bottom: 200, width: 100, height: 100, x: 100, y: 100, toJSON: () => ({}) });
    fireEvent.mouseDown(dialog, { clientX: 150, clientY: 150 });
    fireEvent.click(dialog, { clientX: 150, clientY: 150 });
    fireEvent.click(screen.getByText("Inside"));
    expect(onClose).not.toHaveBeenCalled();
  });

  it("does not close when a drag starts inside and ends on the backdrop", () => {
    const onClose = vi.fn();
    render(<Harness onClose={onClose} />);
    const dialog = screen.getByTestId("dialog");
    dialog.getBoundingClientRect = () => ({ left: 100, right: 200, top: 100, bottom: 200, width: 100, height: 100, x: 100, y: 100, toJSON: () => ({}) });
    fireEvent.mouseDown(dialog, { clientX: 150, clientY: 150 });
    fireEvent.click(dialog, { clientX: 10, clientY: 10 });
    expect(onClose).not.toHaveBeenCalled();
  });

  it("leaves dialogs without an X button alone", () => {
    render(<Harness onClose={vi.fn()} withClose={false} />);
    backdropClick(screen.getByTestId("dialog"));
    expect(screen.getByTestId("dialog").hasAttribute("open")).toBe(true);
  });
});
