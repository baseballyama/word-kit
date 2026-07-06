/**
 * Run-level (character) formatting commands. Each routes through
 * `setRunFormat` / `getRunFormat` from `@office-kit/docx`, applied to exactly
 * the selected characters — {@link applyToSelectionRuns} splits runs at the
 * selection boundaries first, so bolding part of a line bolds only that part.
 */

import { getRunFormat, type RunFormatting, setRunFormat, type WmlRun } from "@office-kit/docx";
import type { EditorModel } from "../model.js";
import { applyToSelectionRuns, overlappingSelectionRuns } from "../selection-runs.js";
import type { Command } from "./types.js";

/** True when every run overlapping the selection has the given boolean on. */
function allHave(runs: WmlRun[], key: "bold" | "italic" | "strike"): boolean {
  return runs.length > 0 && runs.every((r) => getRunFormat(r)[key] === true);
}

function allUnderlined(runs: WmlRun[]): boolean {
  return runs.length > 0 && runs.every((r) => (getRunFormat(r).underline ?? "none") !== "none");
}

function applyToRuns(model: EditorModel, patch: RunFormatting): void {
  applyToSelectionRuns(model, (run) => setRunFormat(run, patch));
}

/** Build a boolean toggle command (bold / italic / strike). */
function toggleBool(id: string, key: "bold" | "italic" | "strike", label: string): Command<void> {
  return {
    id,
    group: "text",
    label,
    run(model) {
      const next = !allHave(overlappingSelectionRuns(model), key);
      applyToSelectionRuns(model, (run) => setRunFormat(run, { [key]: next }));
    },
    isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
    isActive: (model) => allHave(overlappingSelectionRuns(model), key),
  };
}

export const toggleBoldCommand = toggleBool("text.bold", "bold", "Bold");
export const toggleItalicCommand = toggleBool("text.italic", "italic", "Italic");
export const toggleStrikeCommand = toggleBool("text.strike", "strike", "Strikethrough");

export const toggleUnderlineCommand: Command<void> = {
  id: "text.underline",
  group: "text",
  label: "Underline",
  run(model) {
    const next = allUnderlined(overlappingSelectionRuns(model)) ? "none" : "single";
    applyToSelectionRuns(model, (run) => setRunFormat(run, { underline: next }));
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
  isActive: (model) => allUnderlined(overlappingSelectionRuns(model)),
};

/** Set a specific underline style (single/double/wave/dotted/thick). */
export const setUnderlineStyleCommand: Command<{
  style: NonNullable<RunFormatting["underline"]>;
}> = {
  id: "text.underlineStyle",
  group: "text",
  label: "Underline style",
  run(model, { style }) {
    applyToRuns(model, { underline: style });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

/** Set font family (ASCII/hAnsi). */
export const setFontCommand: Command<{ font: string }> = {
  id: "text.font",
  group: "text",
  label: "Font",
  run(model, { font }) {
    applyToRuns(model, { font });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

/** Set East Asian (CJK) font family. */
export const setFontEastAsiaCommand: Command<{ font: string }> = {
  id: "text.fontEastAsia",
  group: "text",
  label: "East Asian font",
  run(model, { font }) {
    applyToRuns(model, { fontEastAsia: font });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

/** Set font size in points; converted to the half-points OOXML stores. */
export const setFontSizeCommand: Command<{ points: number }> = {
  id: "text.fontSize",
  group: "text",
  label: "Font size",
  run(model, { points }) {
    applyToRuns(model, { fontSizeHalfPoints: Math.round(points * 2) });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

/** Set text color as a hex RGB string (no leading #). */
export const setColorCommand: Command<{ color: string }> = {
  id: "text.color",
  group: "text",
  label: "Font color",
  run(model, { color }) {
    applyToRuns(model, { color: normalizeHex(color) });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

/** Set highlight color as a hex RGB string. */
export const setHighlightCommand: Command<{ color: string }> = {
  id: "text.highlight",
  group: "text",
  label: "Highlight",
  run(model, { color }) {
    applyToRuns(model, { highlight: normalizeHex(color) });
  },
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

/** Clear all direct run formatting in the selection. */
export const clearFormattingCommand: Command<void> = {
  id: "text.clearFormat",
  group: "text",
  run(model) {
    applyToSelectionRuns(model, (run) =>
      setRunFormat(run, {
        bold: false,
        italic: false,
        strike: false,
        underline: "none",
      }),
    );
  },
  label: "Clear formatting",
  isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
};

function normalizeHex(color: string): string {
  return color.startsWith("#") ? color.slice(1) : color;
}

export const textCommands = [
  toggleBoldCommand,
  toggleItalicCommand,
  toggleStrikeCommand,
  toggleUnderlineCommand,
  setUnderlineStyleCommand,
  setFontCommand,
  setFontEastAsiaCommand,
  setFontSizeCommand,
  setColorCommand,
  setHighlightCommand,
  clearFormattingCommand,
];
