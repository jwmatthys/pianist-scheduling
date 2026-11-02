import { useEffect } from "react";

// Clicking the backdrop of an open dialog behaves like its X button. Dialogs without an X
// button (confirmations that need an explicit choice) are left alone.
export function useDismissOnBackdrop(): void {
  useEffect(() => {
    let pressedOnBackdrop = false;

    function isBackdropPoint(dialog: HTMLDialogElement, event: MouseEvent): boolean {
      if (event.target !== dialog) return false;
      const rect = dialog.getBoundingClientRect();
      return event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom;
    }

    function onMouseDown(event: MouseEvent) {
      pressedOnBackdrop = event.target instanceof HTMLDialogElement && isBackdropPoint(event.target, event);
    }

    // Requiring the press to start on the backdrop keeps a text selection that is dragged out of the dialog from closing it.
    function onClick(event: MouseEvent) {
      const dialog = event.target;
      if (!(dialog instanceof HTMLDialogElement) || !dialog.open || !pressedOnBackdrop || !isBackdropPoint(dialog, event)) return;
      pressedOnBackdrop = false;
      dialog.querySelector<HTMLButtonElement>(":scope .dialog-close")?.click();
    }

    document.addEventListener("mousedown", onMouseDown, true);
    document.addEventListener("click", onClick, true);
    return () => {
      document.removeEventListener("mousedown", onMouseDown, true);
      document.removeEventListener("click", onClick, true);
    };
  }, []);
}
