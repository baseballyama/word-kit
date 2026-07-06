/**
 * Section / page-setup commands: page size, margins, orientation, and section
 * breaks. These operate on the document's trailing section properties via the
 * `@office-kit/docx` page functions.
 */

import {
  appendSectionBreak,
  type PageMargins,
  type PageSize,
  setPageMargins,
  setPageOrientation,
  setPageSize,
} from "@office-kit/docx";
import { caretBlockIndex, moveLastBlockAfter } from "./insert-util.js";
import type { Command } from "./types.js";

export const setPageSizeCommand: Command<{ size: PageSize }> = {
  id: "section.pageSize",
  group: "section",
  label: "Page size",
  run(model, { size }) {
    setPageSize(model.doc, size);
  },
};

export const setPageMarginsCommand: Command<{ margins: PageMargins }> = {
  id: "section.pageMargins",
  group: "section",
  label: "Margins",
  run(model, { margins }) {
    setPageMargins(model.doc, margins);
  },
};

export const setOrientationCommand: Command<{ orientation: "portrait" | "landscape" }> = {
  id: "section.orientation",
  group: "section",
  label: "Orientation",
  run(model, { orientation }) {
    setPageOrientation(model.doc, orientation);
  },
};

export const insertSectionBreakCommand: Command<{
  type?: "continuous" | "nextPage" | "evenPage" | "oddPage" | "nextColumn";
}> = {
  id: "section.break",
  group: "section",
  label: "Section break",
  run(model, { type = "nextPage" }) {
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    appendSectionBreak(model.doc, type);
    moveLastBlockAfter(model.doc, at);
  },
};

export const sectionCommands = [
  setPageSizeCommand,
  setPageMarginsCommand,
  setOrientationCommand,
  insertSectionBreakCommand,
];
