/**
 * Map a browser DOM selection to the editor's document positions, using the
 * `data-wk-*` anchors the renderer emits. Browser-only; guarded so importing
 * the module in Node (tests) doesn't touch `window`.
 */

import type { CellCoord, DocPosition, Selection } from "./selection.js";

function parseCell(value: string | null): CellCoord | undefined {
  if (!value) return undefined;
  const [r, c] = value.split(",").map((n) => Number.parseInt(n, 10));
  if (r === undefined || c === undefined || Number.isNaN(r) || Number.isNaN(c)) return undefined;
  return { row: r, col: c };
}

/** Walk up from a DOM node to the nearest element carrying `data-wk-block`. */
function anchorElement(node: Node | null): HTMLElement | null {
  let el: Node | null = node;
  while (el && !(el instanceof HTMLElement && el.hasAttribute("data-wk-block"))) {
    el = el.parentNode;
  }
  return el as HTMLElement | null;
}

/** Build a {@link DocPosition} from a DOM container node and character offset. */
export function positionFromDom(node: Node | null, offset: number): DocPosition | null {
  const anchor = anchorElement(node);
  if (!anchor) return null;
  const block = Number.parseInt(anchor.getAttribute("data-wk-block") ?? "", 10);
  if (Number.isNaN(block)) return null;
  const cell = parseCell(anchor.getAttribute("data-wk-cell"));
  const inlineAttr = anchor.getAttribute("data-wk-inline");
  const inline = inlineAttr ? Number.parseInt(inlineAttr, 10) : undefined;
  return {
    block,
    ...(cell ? { cell } : {}),
    ...(inline !== undefined && !Number.isNaN(inline) ? { inline } : {}),
    offset,
  };
}

/** Read the current window selection as an editor {@link Selection}. */
export function readDomSelection(root: Document | ShadowRoot = document): Selection | null {
  const domSel =
    "getSelection" in root ? (root as Document).getSelection?.() : window.getSelection();
  if (!domSel || domSel.rangeCount === 0) return null;
  const range = domSel.getRangeAt(0);
  const anchor = positionFromDom(range.startContainer, range.startOffset);
  const focus = positionFromDom(range.endContainer, range.endOffset);
  if (!anchor || !focus) return null;
  return { anchor, focus };
}
