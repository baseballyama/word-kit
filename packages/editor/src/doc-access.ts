/**
 * Helpers for resolving selection positions to the live AST nodes the
 * `@office-kit/docx` mutation functions operate on. These read the document
 * structure only; they never serialize or mutate.
 */

import type { Docx, WmlBlock, WmlParagraph, WmlRun, WmlTable } from "@office-kit/docx";
import type { DocPosition, OrderedSelection } from "./selection.js";

/** All top-level blocks in document order. */
export function blocks(doc: Docx): WmlBlock[] {
  return doc.document.body.blocks;
}

export function blockAt(doc: Docx, index: number): WmlBlock | undefined {
  return doc.document.body.blocks[index];
}

/** Narrow a block to a paragraph, or undefined if it is a table / raw block. */
export function asParagraph(block: WmlBlock | undefined): WmlParagraph | undefined {
  return block && block.kind === "paragraph" ? block : undefined;
}

export function asTable(block: WmlBlock | undefined): WmlTable | undefined {
  return block && block.kind === "table" ? block : undefined;
}

/**
 * Resolve a position to the paragraph it addresses, following into a table cell
 * when the position carries cell coordinates. Returns undefined when the
 * position does not land on a paragraph (e.g. a raw block).
 */
export function paragraphAt(doc: Docx, pos: DocPosition): WmlParagraph | undefined {
  const block = blockAt(doc, pos.block);
  if (!block) return undefined;
  if (block.kind === "paragraph") return block;
  if (block.kind === "table" && pos.cell) {
    const row = block.rows[pos.cell.row];
    const cell = row?.cells[pos.cell.col];
    if (!cell) return undefined;
    return cell.paragraphs[pos.para ?? 0];
  }
  return undefined;
}

/** The runs of the paragraph a position addresses. */
export function runsAt(doc: Docx, pos: DocPosition): WmlRun[] {
  const para = paragraphAt(doc, pos);
  if (!para) return [];
  return para.children.filter((c): c is WmlRun => c.kind === "run");
}

/**
 * Collect every paragraph touched by an ordered selection, walking top-level
 * blocks and (for table blocks) every cell paragraph in range. Used by
 * paragraph-level commands (alignment, indent, style…) that apply to the whole
 * span, not just the caret paragraph.
 */
export function paragraphsInRange(doc: Docx, sel: OrderedSelection): WmlParagraph[] {
  const out: WmlParagraph[] = [];
  const all = blocks(doc);
  for (let i = sel.start.block; i <= sel.end.block && i < all.length; i++) {
    const block = all[i];
    if (!block) continue;
    if (block.kind === "paragraph") {
      out.push(block);
    } else if (block.kind === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          for (const p of cell.paragraphs) out.push(p);
        }
      }
    }
  }
  return out;
}

/**
 * Every run touched by an ordered selection. When the selection is collapsed
 * (a caret) this is the runs of the caret paragraph, letting toggle commands
 * (bold on an empty selection) still report/flip state.
 */
export function runsInRange(doc: Docx, sel: OrderedSelection): WmlRun[] {
  const out: WmlRun[] = [];
  for (const para of paragraphsInRange(doc, sel)) {
    for (const child of para.children) {
      if (child.kind === "run") out.push(child);
    }
  }
  return out;
}
