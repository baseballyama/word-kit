/**
 * Helpers for inserting a freshly-built block at a specific position.
 *
 * The `@office-kit/docx` builders (`appendParagraph`, `addTable`, …) construct a
 * node and push it to the end of the body. To insert at the caret instead, we
 * let the builder create the node at the end, then move that last block to the
 * target index. `body.blocks` is a public mutable array, so this stays on the
 * supported path while getting positional inserts the append API doesn't offer.
 */

import type { Docx, WmlBlock } from "@office-kit/docx";

/** Move the block the builder just appended to sit right after `afterIndex`. */
export function moveLastBlockAfter(doc: Docx, afterIndex: number): void {
  const list = doc.document.body.blocks;
  if (list.length === 0) return;
  const node = list.pop() as WmlBlock;
  const target = Math.min(Math.max(afterIndex + 1, 0), list.length);
  list.splice(target, 0, node);
}

/** The block index the caret currently sits on, or the last block. */
export function caretBlockIndex(doc: Docx, block: number | undefined): number {
  const count = doc.document.body.blocks.length;
  if (block === undefined) return Math.max(count - 1, 0);
  return Math.min(Math.max(block, 0), Math.max(count - 1, 0));
}
