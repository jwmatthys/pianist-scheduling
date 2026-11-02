// Pure helpers for browser-style Find-on-Page. Matching is case-insensitive substring
// search over rendered text nodes; highlighting uses the CSS Custom Highlight API so the
// React-managed DOM is never mutated.

export const MATCH_HIGHLIGHT = "find-match";
export const CURRENT_HIGHLIGHT = "find-current";
export const DEFAULT_MATCH_LIMIT = 2000;

const SKIPPED_TAGS = new Set(["SCRIPT", "STYLE", "NOSCRIPT", "TEXTAREA", "SELECT", "OPTION"]);

function isRendered(element: Element, root: Element, cache: Map<Element, boolean>): boolean {
  const cached = cache.get(element);
  if (cached !== undefined) return cached;
  let rendered = !SKIPPED_TAGS.has(element.tagName) && !element.hasAttribute("hidden");
  if (rendered) {
    const style = getComputedStyle(element);
    rendered = style.display !== "none" && style.visibility !== "hidden";
  }
  if (rendered && element !== root && element.parentElement) {
    rendered = isRendered(element.parentElement, root, cache);
  }
  cache.set(element, rendered);
  return rendered;
}

export function findTextRanges(root: Element, query: string, limit = DEFAULT_MATCH_LIMIT): Range[] {
  const needle = query.toLowerCase();
  if (!needle) return [];
  const ranges: Range[] = [];
  const cache = new Map<Element, boolean>();
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node && ranges.length < limit; node = walker.nextNode()) {
    const parent = node.parentElement;
    if (!parent || !isRendered(parent, root, cache)) continue;
    const text = (node as Text).data;
    const lowered = text.toLowerCase();
    // Some characters change length when lowercased; offsets would no longer line up.
    if (lowered.length !== text.length) continue;
    for (let at = lowered.indexOf(needle); at !== -1 && ranges.length < limit; at = lowered.indexOf(needle, at + needle.length)) {
      const range = document.createRange();
      range.setStart(node, at);
      range.setEnd(node, at + needle.length);
      ranges.push(range);
    }
  }
  return ranges;
}

// Identifies the visible page/tab: the text of every active navigation control in scope.
export function pageSignature(root: Element | null): string {
  if (!root) return "";
  return Array.from(root.querySelectorAll('[aria-current="page"]'))
    .map((element) => element.textContent ?? "")
    .join("|");
}

type HighlightRegistry = {
  set(name: string, highlight: unknown): void;
  delete(name: string): void;
};
type HighlightConstructor = new (...ranges: Range[]) => { priority: number };

function highlightSupport(): { registry: HighlightRegistry; Highlight: HighlightConstructor } | null {
  const scope = globalThis as unknown as {
    CSS?: { highlights?: HighlightRegistry };
    Highlight?: HighlightConstructor;
  };
  if (!scope.CSS?.highlights || !scope.Highlight) return null;
  return { registry: scope.CSS.highlights, Highlight: scope.Highlight };
}

export function applyHighlights(ranges: Range[], current: Range | null): void {
  const support = highlightSupport();
  if (!support) return;
  const others = ranges.filter((range) => range !== current);
  support.registry.set(MATCH_HIGHLIGHT, new support.Highlight(...others));
  if (current) {
    const currentHighlight = new support.Highlight(current);
    currentHighlight.priority = 1;
    support.registry.set(CURRENT_HIGHLIGHT, currentHighlight);
  } else {
    support.registry.delete(CURRENT_HIGHLIGHT);
  }
}

export function clearHighlights(): void {
  const support = highlightSupport();
  if (!support) return;
  support.registry.delete(MATCH_HIGHLIGHT);
  support.registry.delete(CURRENT_HIGHLIGHT);
}

export function scrollRangeIntoView(range: Range): void {
  const target = range.startContainer.parentElement;
  target?.scrollIntoView?.({ block: "center", inline: "nearest" });
}
