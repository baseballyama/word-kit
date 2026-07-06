/**
 * List commands: turn the selected paragraphs into a bullet or numbered list,
 * or drop an entire list block. `applyListToParagraph` binds a paragraph to a
 * numbering definition; `addBulletList` / `addNumberedList` seed the numbering
 * part and append fresh list items.
 */

import {
  addBulletList,
  addNumberedList,
  applyListToParagraph,
  getParagraphNumbering,
} from "@office-kit/docx";
import { paragraphsInRange } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { orderSelection } from "../selection.js";
import { caretBlockIndex, moveLastBlockAfter } from "./insert-util.js";
import type { Command } from "./types.js";

/** Insert a new bullet/numbered list built from lines, after the caret. */
function insertList(
  model: EditorModel,
  items: readonly string[],
  build: typeof addBulletList,
): void {
  const at = caretBlockIndex(model.doc, model.selection?.focus.block);
  const paras = build(model.doc, items.length ? items : [""]);
  // Builders append each list paragraph to the end; move them as a group.
  for (let i = 0; i < paras.length; i++) moveLastBlockAfter(model.doc, at + i);
}

export const insertBulletListCommand: Command<{ items: readonly string[] }> = {
  id: "list.insertBullet",
  group: "list",
  label: "Bulleted list",
  run(model, { items }) {
    insertList(model, items, addBulletList);
  },
};

export const insertNumberedListCommand: Command<{ items: readonly string[] }> = {
  id: "list.insertNumbered",
  group: "list",
  label: "Numbered list",
  run(model, { items }) {
    insertList(model, items, addNumberedList);
  },
};

/** Bind every selected paragraph to a list (reusing or seeding a numId). */
export const applyListCommand: Command<{ kind: "bullet" | "numbered" }> = {
  id: "list.apply",
  group: "list",
  label: "Apply list",
  run(model, { kind }) {
    const sel = model.selection;
    if (!sel) return;
    const paras = paragraphsInRange(model.doc, orderSelection(sel));
    if (paras.length === 0) return;
    // Seed a list definition of the requested kind, read the numId off the
    // seeded paragraph, bind the real selection to it, then drop the seed.
    const seed =
      kind === "bullet" ? addBulletList(model.doc, [""]) : addNumberedList(model.doc, [""]);
    const seededBlock = seed[0];
    const numId = seededBlock ? getParagraphNumbering(seededBlock)?.numId : undefined;
    if (seededBlock) {
      const idx = model.doc.document.body.blocks.indexOf(seededBlock);
      if (idx >= 0) model.doc.document.body.blocks.splice(idx, 1);
    }
    if (numId === undefined) return;
    for (const p of paras) applyListToParagraph(model.doc, p, numId);
  },
  isEnabled: (model) => {
    const sel = model.selection;
    return !!sel && paragraphsInRange(model.doc, orderSelection(sel)).length > 0;
  },
};

export const listCommands = [insertBulletListCommand, insertNumberedListCommand, applyListCommand];
