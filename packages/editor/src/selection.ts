/**
 * Document selection model.
 *
 * A position addresses a place in the document by block index (into
 * {@link WmlBody.blocks}), and optionally the inline (run) index within a
 * paragraph and a character offset within that run's text. Table cells are
 * addressed through {@link DocPosition.cell} when the block is a table.
 *
 * The editor keeps selection here — decoupled from the DOM — so commands can
 * reason about "what is selected" without touching rendered nodes. The
 * DOM↔selection mapping lives in {@link ./dom-selection.ts}.
 */

/** Coordinates of a table cell inside a table block. */
export interface CellCoord {
  readonly row: number;
  readonly col: number;
}

/** A single caret position in the document. */
export interface DocPosition {
  /** Index into `body.blocks`. */
  readonly block: number;
  /** When the block is a table: which cell. */
  readonly cell?: CellCoord;
  /** When inside a paragraph (top-level or in a cell): paragraph index in the cell. */
  readonly para?: number;
  /** Inline (run) index within the target paragraph. */
  readonly inline?: number;
  /** Character offset within the target run's text. */
  readonly offset?: number;
}

/** A selection range from anchor (where it started) to focus (the moving end). */
export interface Selection {
  readonly anchor: DocPosition;
  readonly focus: DocPosition;
}

/** A selection reduced to document order: start comes at or before end. */
export interface OrderedSelection {
  readonly start: DocPosition;
  readonly end: DocPosition;
  /** True when start and end are the same position (a caret, not a range). */
  readonly collapsed: boolean;
}

function comparePositions(a: DocPosition, b: DocPosition): number {
  if (a.block !== b.block) return a.block - b.block;
  const ar = a.cell?.row ?? -1;
  const br = b.cell?.row ?? -1;
  if (ar !== br) return ar - br;
  const ac = a.cell?.col ?? -1;
  const bc = b.cell?.col ?? -1;
  if (ac !== bc) return ac - bc;
  if ((a.para ?? 0) !== (b.para ?? 0)) return (a.para ?? 0) - (b.para ?? 0);
  if ((a.inline ?? 0) !== (b.inline ?? 0)) return (a.inline ?? 0) - (b.inline ?? 0);
  return (a.offset ?? 0) - (b.offset ?? 0);
}

/** Order a selection so start ≤ end, and report whether it is collapsed. */
export function orderSelection(sel: Selection): OrderedSelection {
  const cmp = comparePositions(sel.anchor, sel.focus);
  const start = cmp <= 0 ? sel.anchor : sel.focus;
  const end = cmp <= 0 ? sel.focus : sel.anchor;
  return { start, end, collapsed: cmp === 0 };
}

/** Build a collapsed selection (a caret) at a single position. */
export function caretAt(pos: DocPosition): Selection {
  return { anchor: pos, focus: pos };
}
