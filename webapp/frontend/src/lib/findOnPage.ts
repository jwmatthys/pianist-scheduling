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

export type FindMatch = Range | HTMLElement;

const CONTROL_ATTRIBUTE = "data-find-match";
const NON_TEXT_INPUTS = new Set(["hidden", "checkbox", "radio", "button", "submit", "reset", "file", "image", "range", "color", "password"]);

function controlText(element: Element): string | null {
  if (element instanceof HTMLTextAreaElement) return element.value;
  if (element instanceof HTMLSelectElement) return element.selectedOptions[0]?.text ?? "";
  if (element instanceof HTMLInputElement && !NON_TEXT_INPUTS.has(element.type)) return element.value;
  return null;
}

function isControlRendered(element: Element, root: Element): boolean {
  if (element.hasAttribute("hidden") || element.closest('[role="search"][aria-label="Find on page"]')) return false;
  const style = getComputedStyle(element);
  if (style.display === "none" || style.visibility === "hidden") return false;
  return !element.parentElement || isRendered(element.parentElement, root, new Map());
}

// Text boxes (inputs, textareas, and the selected label of dropdowns) in document order.
export function findControlMatches(root: Element, query: string, limit = DEFAULT_MATCH_LIMIT): HTMLElement[] {
  const needle = query.toLowerCase();
  if (!needle) return [];
  const matches: HTMLElement[] = [];
  for (const element of Array.from(root.querySelectorAll<HTMLElement>("input, textarea, select"))) {
    if (matches.length >= limit) break;
    const value = controlText(element);
    if (value && value.toLowerCase().includes(needle) && isControlRendered(element, root)) matches.push(element);
  }
  return matches;
}

// Text and control matches merged into document order.
export function findMatches(root: Element, query: string, limit = DEFAULT_MATCH_LIMIT): FindMatch[] {
  const nodeOf = (match: FindMatch): Node => (match instanceof Range ? match.startContainer : match);
  const combined: FindMatch[] = [...findTextRanges(root, query, limit), ...findControlMatches(root, query, limit)];
  combined.sort((a, b) => {
    const position = nodeOf(a).compareDocumentPosition(nodeOf(b));
    return position & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : position & Node.DOCUMENT_POSITION_PRECEDING ? 1 : 0;
  });
  return combined.slice(0, limit);
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

function clearControlMarks(): void {
  document.querySelectorAll(`[${CONTROL_ATTRIBUTE}]`).forEach((element) => element.removeAttribute(CONTROL_ATTRIBUTE));
}

export function applyHighlights(matches: FindMatch[], current: FindMatch | null): void {
  clearControlMarks();
  for (const match of matches) {
    if (!(match instanceof Range)) match.setAttribute(CONTROL_ATTRIBUTE, match === current ? "current" : "match");
  }
  const support = highlightSupport();
  if (!support) return;
  const others = matches.filter((match): match is Range => match instanceof Range && match !== current);
  support.registry.set(MATCH_HIGHLIGHT, new support.Highlight(...others));
  if (current instanceof Range) {
    const currentHighlight = new support.Highlight(current);
    currentHighlight.priority = 1;
    support.registry.set(CURRENT_HIGHLIGHT, currentHighlight);
  } else {
    support.registry.delete(CURRENT_HIGHLIGHT);
  }
}

export function clearHighlights(): void {
  clearControlMarks();
  const support = highlightSupport();
  if (!support) return;
  support.registry.delete(MATCH_HIGHLIGHT);
  support.registry.delete(CURRENT_HIGHLIGHT);
}

export function scrollRangeIntoView(match: FindMatch): void {
  const target = match instanceof Range ? match.startContainer.parentElement : match;
  target?.scrollIntoView?.({ block: "center", inline: "nearest" });
}
