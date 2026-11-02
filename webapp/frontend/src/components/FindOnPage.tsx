import { useCallback, useEffect, useRef, useState, type KeyboardEvent, type RefObject } from "react";
import "./FindOnPage.css";
import {
  applyHighlights,
  clearHighlights,
  findTextRanges,
  pageSignature,
  scrollRangeIntoView,
} from "../lib/findOnPage";

type Props = {
  scopeRef: RefObject<HTMLElement | null>;
  // Changing this value (for example, entering another module) closes the search.
  resetKey?: string;
};

// Browser-style Find-on-Page for the content inside `scopeRef`. It only highlights matches;
// it never filters, navigates, or changes page state.
export function FindOnPage({ scopeRef, resetKey }: Props) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [ranges, setRanges] = useState<Range[]>([]);
  const [index, setIndex] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);
  const returnFocusRef = useRef<HTMLElement | null>(null);
  const signatureRef = useRef("");
  const scrollPendingRef = useRef(false);

  const reset = useCallback(() => {
    setOpen(false);
    setQuery("");
    setRanges([]);
    setIndex(0);
    clearHighlights();
  }, []);

  const close = useCallback(() => {
    reset();
    returnFocusRef.current?.focus?.();
    returnFocusRef.current = null;
  }, [reset]);

  useEffect(() => {
    function onKeyDown(event: globalThis.KeyboardEvent) {
      if ((event.ctrlKey || event.metaKey) && !event.altKey && !event.shiftKey && event.key.toLowerCase() === "f") {
        event.preventDefault();
        if (!open) {
          returnFocusRef.current = document.activeElement instanceof HTMLElement ? document.activeElement : null;
          signatureRef.current = pageSignature(scopeRef.current);
          setOpen(true);
        } else {
          inputRef.current?.focus();
          inputRef.current?.select();
        }
      } else if (event.key === "Escape" && open) {
        event.preventDefault();
        close();
      }
    }
    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, [open, close, scopeRef]);

  useEffect(() => {
    if (!open) return;
    inputRef.current?.focus();
    inputRef.current?.select();
  }, [open]);

  useEffect(() => {
    reset();
  }, [resetKey, reset]);

  useEffect(() => clearHighlights, []);

  // Typing recomputes matches immediately and jumps to the first one.
  useEffect(() => {
    if (!open) return;
    const root = scopeRef.current;
    scrollPendingRef.current = true;
    setIndex(0);
    setRanges(root && query ? findTextRanges(root, query) : []);
  }, [open, query, scopeRef]);

  // Content can change while searching (data loads, rows edited); a tab change closes the search.
  useEffect(() => {
    const root = scopeRef.current;
    if (!open || !root) return;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const observer = new MutationObserver(() => {
      if (pageSignature(root) !== signatureRef.current) {
        reset();
        return;
      }
      clearTimeout(timer);
      timer = setTimeout(() => {
        setRanges(query ? findTextRanges(root, query) : []);
      }, 100);
    });
    observer.observe(root, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ["aria-current", "hidden", "style", "class"] });
    return () => {
      observer.disconnect();
      clearTimeout(timer);
    };
  }, [open, query, scopeRef, reset]);

  useEffect(() => {
    if (!open || ranges.length === 0) {
      clearHighlights();
      return;
    }
    const current = ranges[Math.min(index, ranges.length - 1)];
    applyHighlights(ranges, current);
    if (scrollPendingRef.current) {
      scrollPendingRef.current = false;
      scrollRangeIntoView(current);
    }
  }, [open, ranges, index]);

  function step(delta: number) {
    if (ranges.length === 0) return;
    scrollPendingRef.current = true;
    setIndex((current) => (Math.min(current, ranges.length - 1) + delta + ranges.length) % ranges.length);
  }

  function onBarKeyDown(event: KeyboardEvent<HTMLDivElement>) {
    if (event.key === "ArrowDown" || event.key === "ArrowRight" || (event.key === "Enter" && !event.shiftKey && event.target === inputRef.current)) {
      event.preventDefault();
      step(1);
    } else if (event.key === "ArrowUp" || event.key === "ArrowLeft" || (event.key === "Enter" && event.shiftKey)) {
      event.preventDefault();
      step(-1);
    }
  }

  if (!open) return null;

  const position = Math.min(index, Math.max(ranges.length - 1, 0));
  const status = !query ? "" : ranges.length === 0 ? "No matches" : `${position + 1} of ${ranges.length}`;

  return (
    <div className="find-on-page" role="search" aria-label="Find on page" onKeyDown={onBarKeyDown}>
      <input
        ref={inputRef}
        type="text"
        aria-label="Find on page"
        placeholder="Find on page"
        autoComplete="off"
        spellCheck={false}
        value={query}
        onChange={(event) => setQuery(event.target.value)}
      />
      <span className="find-on-page-status" role="status" aria-live="polite">{status}</span>
      <button type="button" aria-label="Previous match" disabled={ranges.length === 0} onClick={() => step(-1)}>{"\u2191"}</button>
      <button type="button" aria-label="Next match" disabled={ranges.length === 0} onClick={() => step(1)}>{"\u2193"}</button>
      <button type="button" aria-label="Close find" onClick={close}>{"\u00d7"}</button>
    </div>
  );
}
