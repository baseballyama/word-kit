/**
 * Style commands: define a new paragraph/character style in `styles.xml`, and
 * ensure the built-in heading styles exist. Applying a style to a paragraph is
 * `paragraph.style` (see ./paragraph.ts).
 */

import { addStyle, type BuildStyleOptions, ensureHeadingStyles } from "@office-kit/docx";
import type { Command } from "./types.js";

/** Options for a new style. */
export type AddStyleParams = BuildStyleOptions;

export const addStyleCommand: Command<AddStyleParams> = {
  id: "style.add",
  group: "style",
  label: "New style",
  run(model, options) {
    addStyle(model.doc, options);
  },
};

export const ensureHeadingStylesCommand: Command<void> = {
  id: "style.ensureHeadings",
  group: "style",
  label: "Ensure heading styles",
  run(model) {
    ensureHeadingStyles(model.doc);
  },
};

export const styleCommands = [addStyleCommand, ensureHeadingStylesCommand];
