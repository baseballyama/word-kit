/**
 * Block-structure commands: insert paragraphs / headings / breaks, and remove
 * blocks. Positional inserts use the append-then-move helper so new content
 * lands at the caret rather than the end of the document.
 */

import {
  appendHeading,
  appendLineBreak,
  appendPageBreak,
  appendParagraph,
  type AppendParagraphOptions,
  ensureHeadingStyles,
  mergeParagraphIntoPrevious,
  paragraphs,
  removeParagraph,
  removeTable,
  splitParagraphAt,
} from "@office-kit/docx";
import { blockAt, paragraphAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { caretAt } from "../selection.js";
import { caretBlockIndex, moveLastBlockAfter } from "./insert-util.js";
import type { Command } from "./types.js";

/**
 * Split the caret paragraph in two at the caret (the Enter key). Falls back to
 * inserting an empty paragraph after the caret block when the caret is not on a
 * top-level paragraph (e.g. inside a table cell).
 */
export const splitParagraphCommand: Command<void> = {
  id: "structure.splitParagraph",
  group: "structure",
  label: "Split paragraph",
  run(model) {
    const pos = model.selection?.focus;
    if (!pos) return;
    const newBlock = splitParagraphAt(model.doc, pos.block, pos.inline ?? 0, pos.offset ?? 0);
    if (newBlock >= 0) {
      model.setSelection(caretAt({ block: newBlock, inline: 0, offset: 0 }));
    } else {
      const at = caretBlockIndex(model.doc, pos.block);
      appendParagraph(model.doc, "", {});
      moveLastBlockAfter(model.doc, at);
      model.setSelection(caretAt({ block: at + 1 }));
    }
  },
};

/**
 * Merge the caret paragraph into the previous one (Backspace at paragraph
 * start). No-op when there is no preceding paragraph.
 */
export const mergeBackCommand: Command<void> = {
  id: "structure.mergeBack",
  group: "structure",
  label: "Merge with previous paragraph",
  run(model) {
    const pos = model.selection?.focus;
    if (!pos) return;
    const joined = mergeParagraphIntoPrevious(model.doc, pos.block);
    if (joined) model.setSelection(caretAt(joined));
  },
  isEnabled: (model) => {
    const block = model.selection?.focus.block;
    return block !== undefined && block > 0;
  },
};

export const insertParagraphCommand: Command<{ text?: string; options?: AppendParagraphOptions }> =
  {
    id: "structure.insertParagraph",
    group: "structure",
    label: "Insert paragraph",
    run(model, params) {
      const at = caretBlockIndex(model.doc, model.selection?.focus.block);
      appendParagraph(model.doc, params.text ?? "", params.options ?? {});
      moveLastBlockAfter(model.doc, at);
      model.setSelection(caretAt({ block: at + 1 }));
    },
  };

export const insertHeadingCommand: Command<{ text: string; level?: number }> = {
  id: "structure.insertHeading",
  group: "structure",
  label: "Insert heading",
  run(model, { text, level = 1 }) {
    // Define the built-in heading styles so the heading's pStyle reference is
    // not dangling (Word would otherwise fall back to Normal).
    ensureHeadingStyles(model.doc);
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    appendHeading(model.doc, text, level);
    moveLastBlockAfter(model.doc, at);
    model.setSelection(caretAt({ block: at + 1 }));
  },
};

export const insertPageBreakCommand: Command<void> = {
  id: "structure.insertPageBreak",
  group: "structure",
  label: "Page break",
  run(model) {
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    appendPageBreak(model.doc);
    moveLastBlockAfter(model.doc, at);
  },
};

export const insertLineBreakCommand: Command<{ kind?: "line" | "page" | "column" }> = {
  id: "structure.insertLineBreak",
  group: "structure",
  label: "Line break",
  run(model, { kind = "line" }) {
    const pos = model.selection?.focus;
    const para = pos ? paragraphAt(model.doc, pos) : undefined;
    if (!para) return;
    appendLineBreak(model.doc, para, kind);
  },
  isEnabled: (model) => {
    const pos = model.selection?.focus;
    return !!(pos && paragraphAt(model.doc, pos));
  },
};

export const deleteBlockCommand: Command<void> = {
  id: "structure.deleteBlock",
  group: "structure",
  label: "Delete block",
  run(model) {
    const block = model.selection?.focus.block;
    if (block === undefined) return;
    const node = blockAt(model.doc, block);
    if (!node) return;
    if (node.kind === "table") {
      const tableIndex = countKind(model, "table", block);
      removeTable(model.doc, tableIndex);
    } else if (node.kind === "paragraph") {
      const paraIndex = paragraphs(model.doc).indexOf(node);
      if (paraIndex >= 0) removeParagraph(model.doc, paraIndex);
    }
    model.setSelection(caretAt({ block: Math.max(block - 1, 0) }));
  },
  isEnabled: (model) => {
    const block = model.selection?.focus.block;
    return block !== undefined && !!blockAt(model.doc, block);
  },
};

/** Index of a block among blocks of the same kind up to `blockIndex`. */
function countKind(model: EditorModel, kind: "table" | "paragraph", blockIndex: number): number {
  const list = model.doc.document.body.blocks;
  let n = 0;
  for (let i = 0; i < blockIndex && i < list.length; i++) {
    if (list[i]?.kind === kind) n++;
  }
  return n;
}

export const structureCommands = [
  insertParagraphCommand,
  splitParagraphCommand,
  mergeBackCommand,
  insertHeadingCommand,
  insertPageBreakCommand,
  insertLineBreakCommand,
  deleteBlockCommand,
];
