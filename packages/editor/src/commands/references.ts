/**
 * Reference commands: bookmarks, hyperlinks (external + internal anchor),
 * footnotes, endnotes, fields (PAGE/DATE/…), MERGEFIELD, and a table of
 * contents. Paragraph-scoped commands attach to the caret paragraph.
 */

import {
  addBookmark,
  addEndnote,
  addFootnote,
  addHyperlink,
  addInternalHyperlink,
  addTableOfContents,
  type AddTableOfContentsOptions,
  appendField,
  appendMergeField,
} from "@office-kit/docx";
import { paragraphAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import type { Command } from "./types.js";
import { caretBlockIndex, moveLastBlockAfter } from "./insert-util.js";

/** The caret paragraph, or undefined when the caret is not in a paragraph. */
function caretParagraph(model: EditorModel) {
  const pos = model.selection?.focus;
  return pos ? paragraphAt(model.doc, pos) : undefined;
}

export const addBookmarkCommand: Command<{ name: string }> = {
  id: "references.bookmark",
  group: "references",
  label: "Bookmark",
  run(model, { name }) {
    const para = caretParagraph(model);
    if (para) addBookmark(model.doc, name, para);
  },
  isEnabled: (model) => !!caretParagraph(model),
};

export const insertHyperlinkCommand: Command<{ url: string; text: string; tooltip?: string }> = {
  id: "references.hyperlink",
  group: "references",
  label: "Hyperlink",
  run(model, { url, text, tooltip }) {
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    addHyperlink(model.doc, url, text, tooltip !== undefined ? { tooltip } : {});
    moveLastBlockAfter(model.doc, at);
  },
};

export const insertInternalLinkCommand: Command<{
  bookmark: string;
  text: string;
  tooltip?: string;
}> = {
  id: "references.internalLink",
  group: "references",
  label: "Link to bookmark",
  run(model, { bookmark, text, tooltip }) {
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    addInternalHyperlink(model.doc, bookmark, text, tooltip !== undefined ? { tooltip } : {});
    moveLastBlockAfter(model.doc, at);
  },
};

export const addFootnoteCommand: Command<{ text: string }> = {
  id: "references.footnote",
  group: "references",
  label: "Footnote",
  run(model, { text }) {
    const para = caretParagraph(model);
    if (para) addFootnote(model.doc, para, text);
  },
  isEnabled: (model) => !!caretParagraph(model),
};

export const addEndnoteCommand: Command<{ text: string }> = {
  id: "references.endnote",
  group: "references",
  label: "Endnote",
  run(model, { text }) {
    const para = caretParagraph(model);
    if (para) addEndnote(model.doc, para, text);
  },
  isEnabled: (model) => !!caretParagraph(model),
};

export const insertFieldCommand: Command<{ instruction: string }> = {
  id: "references.field",
  group: "references",
  label: "Field",
  run(model, { instruction }) {
    const para = caretParagraph(model);
    if (para) appendField(model.doc, para, instruction);
  },
  isEnabled: (model) => !!caretParagraph(model),
};

export const insertMergeFieldCommand: Command<{ fieldName: string; displayText?: string }> = {
  id: "references.mergeField",
  group: "references",
  label: "Merge field",
  run(model, { fieldName, displayText }) {
    const para = caretParagraph(model);
    if (para) appendMergeField(model.doc, para, fieldName, displayText);
  },
  isEnabled: (model) => !!caretParagraph(model),
};

export const insertTocCommand: Command<AddTableOfContentsOptions> = {
  id: "references.toc",
  group: "references",
  label: "Table of contents",
  run(model, options) {
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    addTableOfContents(model.doc, options);
    moveLastBlockAfter(model.doc, at);
  },
};

export const referencesCommands = [
  addBookmarkCommand,
  insertHyperlinkCommand,
  insertInternalLinkCommand,
  addFootnoteCommand,
  addEndnoteCommand,
  insertFieldCommand,
  insertMergeFieldCommand,
  insertTocCommand,
];
