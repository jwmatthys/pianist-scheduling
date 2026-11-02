import { useEffect, useRef } from "react";

/** Returns a ref for the preview section; scrolls it into view whenever a new preview appears. */
export function useScrollToPreview<T extends HTMLElement>(preview: unknown) {
  const ref = useRef<T>(null);
  useEffect(() => {
    if (preview) ref.current?.scrollIntoView?.({ behavior: "smooth", block: "start" });
  }, [preview]);
  return ref;
}
