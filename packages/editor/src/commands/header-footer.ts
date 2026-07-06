/**
 * Header / footer commands. `addHeader` / `addFooter` create the part, wire the
 * relationship, and reference it from the section; `addPageNumberFooter` seeds a
 * centered PAGE field footer.
 */

import { addFooter, addHeader, addPageNumberFooter, type HeaderFooterType } from "@office-kit/docx";
import type { Command } from "./types.js";

export const addHeaderCommand: Command<{ text: string; type?: HeaderFooterType }> = {
  id: "headerFooter.addHeader",
  group: "headerFooter",
  label: "Header",
  run(model, { text, type = "default" }) {
    addHeader(model.doc, text, type);
  },
};

export const addFooterCommand: Command<{ text: string; type?: HeaderFooterType }> = {
  id: "headerFooter.addFooter",
  group: "headerFooter",
  label: "Footer",
  run(model, { text, type = "default" }) {
    addFooter(model.doc, text, type);
  },
};

export const addPageNumberFooterCommand: Command<{ type?: HeaderFooterType }> = {
  id: "headerFooter.pageNumber",
  group: "headerFooter",
  label: "Page number",
  run(model, { type = "default" }) {
    addPageNumberFooter(model.doc, type);
  },
};

export const headerFooterCommands = [
  addHeaderCommand,
  addFooterCommand,
  addPageNumberFooterCommand,
];
