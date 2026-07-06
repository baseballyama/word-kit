/**
 * Selection-aware run access for character formatting.
 *
 * The naïve approach — format every run the selection *touches* — bolds the
 * whole line when you select a few characters, because runs are the atomic unit
 * of formatting. Instead, {@link applyToSelectionRuns} isolates exactly the
 * selected characters into their own runs (splitting at the boundaries) before
 * applying the change, and updates the selection so the same text stays
 * highlighted. {@link overlappingSelectionRuns} is the read-only counterpart for
 * `isActive` / `isEnabled` state (it never mutates the document).
 */

import {
  isolateParagraphRunRange,
  runTextLength,
  type WmlParagraph,
  type WmlRun,
} from "@office-kit/docx";
import { paragraphAt } from "./doc-access.js";
import type { EditorModel } from "./model.js";
import { type CellCoord, type DocPosition, orderSelection } from "./selection.js";

function paragraphRuns(para: WmlParagraph): WmlRun[] {
  return para.children.filter((c): c is WmlRun => c.kind === "run");
}

/** Absolute character offset of a position within its paragraph's run text. */
function absoluteChar(runs: WmlRun[], inline: number, offset: number): number {
  let n = 0;
  for (let i = 0; i < inline && i < runs.length; i++) n += runTextLength(runs[i]!);
  return n + offset;
}

interface ParaWindow {
  para: WmlParagraph;
  startChar: number;
  endChar: number;
  block: number;
  cell?: CellCoord;
}

/** Whether two positions denote the same collapsed caret. */
function isCollapsed(a: DocPosition, b: DocPosition): boolean {
  return (
    a.block === b.block &&
    (a.inline ?? 0) === (b.inline ?? 0) &&
    (a.offset ?? 0) === (b.offset ?? 0) &&
    a.cell?.row === b.cell?.row &&
    a.cell?.col === b.cell?.col
  );
}

/** The paragraph windows the selection covers, with per-paragraph char ranges. */
function selectionWindows(model: EditorModel): ParaWindow[] {
  const sel = model.selection;
  if (!sel) return [];
  const { start, end } = orderSelection(sel);

  // Single paragraph (including inside a table cell): a precise char window.
  if (
    start.block === end.block &&
    start.cell?.row === end.cell?.row &&
    start.cell?.col === end.cell?.col
  ) {
    const para = paragraphAt(model.doc, start);
    if (!para) return [];
    const runs = paragraphRuns(para);
    return [
      {
        para,
        startChar: absoluteChar(runs, start.inline ?? 0, start.offset ?? 0),
        endChar: absoluteChar(runs, end.inline ?? 0, end.offset ?? 0),
        block: start.block,
        ...(start.cell ? { cell: start.cell } : {}),
      },
    ];
  }

  // Multi-block selection over top-level paragraphs. Tables fall back to whole
  // cell paragraphs (a rare selection shape not worth per-char precision here).
  const out: ParaWindow[] = [];
  const blocks = model.doc.document.body.blocks;
  for (let b = start.block; b <= end.block && b < blocks.length; b++) {
    const node = blocks[b];
    if (node?.kind === "paragraph") {
      const runs = paragraphRuns(node);
      const full = runs.reduce((n, r) => n + runTextLength(r), 0);
      out.push({
        para: node,
        startChar: b === start.block ? absoluteChar(runs, start.inline ?? 0, start.offset ?? 0) : 0,
        endChar: b === end.block ? absoluteChar(runs, end.inline ?? 0, end.offset ?? 0) : full,
        block: b,
      });
    } else if (node?.kind === "table") {
      for (const row of node.rows) {
        for (const cell of row.cells) {
          for (const p of cell.paragraphs) {
            const full = paragraphRuns(p).reduce((n, r) => n + runTextLength(r), 0);
            out.push({ para: p, startChar: 0, endChar: full, block: b });
          }
        }
      }
    }
  }
  return out;
}

/** Runs overlapping the selection — read-only, for active/enabled state. */
export function overlappingSelectionRuns(model: EditorModel): WmlRun[] {
  const sel = model.selection;
  if (!sel) return [];
  const { start, end } = orderSelection(sel);
  const wins = selectionWindows(model);

  // Collapsed caret: the run at the caret (so toggles still report state).
  if (isCollapsed(start, end)) {
    const para = paragraphAt(model.doc, start);
    if (!para) return [];
    const runs = paragraphRuns(para);
    const at = runs[start.inline ?? 0];
    return at ? [at] : runs.slice(0, 1);
  }

  const out: WmlRun[] = [];
  for (const w of wins) {
    let cursor = 0;
    for (const run of paragraphRuns(w.para)) {
      const s = cursor;
      const e = cursor + runTextLength(run);
      cursor = e;
      if (e > s && s < w.endChar && e > w.startChar) out.push(run);
    }
  }
  return out;
}

/**
 * Apply `applyFn` to exactly the runs covering the selection, splitting runs at
 * the selection boundaries first so a partial selection formats only the
 * selected characters. Keeps the same text selected afterwards.
 */
export function applyToSelectionRuns(model: EditorModel, applyFn: (run: WmlRun) => void): void {
  const sel = model.selection;
  if (!sel) return;
  const { start, end } = orderSelection(sel);

  // Collapsed caret: apply to the whole caret run (there is no range to isolate).
  if (isCollapsed(start, end)) {
    for (const run of overlappingSelectionRuns(model)) applyFn(run);
    return;
  }

  const wins = selectionWindows(model);
  for (const w of wins) {
    for (const run of isolateParagraphRunRange(w.para, w.startChar, w.endChar)) applyFn(run);
  }

  // Re-anchor the selection to the isolated run boundaries (single paragraph).
  if (wins.length === 1) {
    const w = wins[0]!;
    const runs = paragraphRuns(w.para);
    let cursor = 0;
    let startInline = -1;
    let endInline = -1;
    runs.forEach((run, idx) => {
      const s = cursor;
      const e = cursor + runTextLength(run);
      cursor = e;
      if (e > s && s >= w.startChar && e <= w.endChar) {
        if (startInline < 0) startInline = idx;
        endInline = idx;
      }
    });
    if (startInline >= 0 && endInline >= 0) {
      const cell = w.cell ? { cell: w.cell } : {};
      model.setSelection({
        anchor: { block: w.block, ...cell, inline: startInline, offset: 0 },
        focus: {
          block: w.block,
          ...cell,
          inline: endInline,
          offset: runTextLength(runs[endInline]!),
        },
      });
    }
  }
}
