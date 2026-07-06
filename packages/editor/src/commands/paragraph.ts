/**
 * Paragraph-level formatting commands: alignment, indent, spacing, borders,
 * shading, and paragraph style. Each applies to every paragraph the selection
 * spans, via the `@office-kit/docx` paragraph functions.
 */

import {
  getParagraphAlignment,
  type ParagraphAlignment,
  type ParagraphBordersOptions,
  type ParagraphIndent,
  type ParagraphShadingOptions,
  type ParagraphSpacing,
  setParagraphAlignment,
  setParagraphBorders,
  setParagraphIndent,
  setParagraphShading,
  setParagraphSpacing,
  setParagraphStyle,
  type WmlParagraph,
} from "@office-kit/docx";
import { paragraphsInRange } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { orderSelection } from "../selection.js";
import type { Command } from "./types.js";

function selectedParagraphs(model: EditorModel): WmlParagraph[] {
  const sel = model.selection;
  if (!sel) return [];
  return paragraphsInRange(model.doc, orderSelection(sel));
}

export const setAlignmentCommand: Command<{ alignment: ParagraphAlignment }> = {
  id: "paragraph.align",
  group: "paragraph",
  label: "Alignment",
  run(model, { alignment }) {
    for (const p of selectedParagraphs(model)) setParagraphAlignment(p, alignment);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

/** Build a fixed-alignment command (used for the four ribbon buttons). */
function alignTo(id: string, alignment: ParagraphAlignment, label: string): Command<void> {
  return {
    id,
    group: "paragraph",
    label,
    run(model) {
      for (const p of selectedParagraphs(model)) setParagraphAlignment(p, alignment);
    },
    isEnabled: (model) => selectedParagraphs(model).length > 0,
    isActive(model) {
      const ps = selectedParagraphs(model);
      return ps.length > 0 && ps.every((p) => getParagraphAlignment(p) === alignment);
    },
  };
}

export const alignLeftCommand = alignTo("paragraph.alignLeft", "left", "Align left");
export const alignCenterCommand = alignTo("paragraph.alignCenter", "center", "Center");
export const alignRightCommand = alignTo("paragraph.alignRight", "right", "Align right");
export const alignJustifyCommand = alignTo("paragraph.alignJustify", "both", "Justify");

export const setIndentCommand: Command<ParagraphIndent> = {
  id: "paragraph.indent",
  group: "paragraph",
  label: "Indent",
  run(model, indent) {
    for (const p of selectedParagraphs(model)) setParagraphIndent(p, indent);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

export const setSpacingCommand: Command<ParagraphSpacing> = {
  id: "paragraph.spacing",
  group: "paragraph",
  label: "Line & paragraph spacing",
  run(model, spacing) {
    for (const p of selectedParagraphs(model)) setParagraphSpacing(p, spacing);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

export const setParagraphBordersCommand: Command<ParagraphBordersOptions> = {
  id: "paragraph.borders",
  group: "paragraph",
  label: "Borders",
  run(model, borders) {
    for (const p of selectedParagraphs(model)) setParagraphBorders(p, borders);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

export const setParagraphShadingCommand: Command<ParagraphShadingOptions> = {
  id: "paragraph.shading",
  group: "paragraph",
  label: "Shading",
  run(model, shading) {
    for (const p of selectedParagraphs(model)) setParagraphShading(p, shading);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

export const setParagraphStyleCommand: Command<{ styleId: string | undefined }> = {
  id: "paragraph.style",
  group: "paragraph",
  label: "Paragraph style",
  run(model, { styleId }) {
    for (const p of selectedParagraphs(model)) setParagraphStyle(p, styleId);
  },
  isEnabled: (model) => selectedParagraphs(model).length > 0,
};

export const paragraphCommands = [
  setAlignmentCommand,
  alignLeftCommand,
  alignCenterCommand,
  alignRightCommand,
  alignJustifyCommand,
  setIndentCommand,
  setSpacingCommand,
  setParagraphBordersCommand,
  setParagraphShadingCommand,
  setParagraphStyleCommand,
];
