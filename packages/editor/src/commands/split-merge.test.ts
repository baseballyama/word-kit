/**
 * Split (Enter) and merge (Backspace-at-start) round-trip tests — the core of
 * the interactive editing experience.
 */

import {
  createDocx,
  openDocx,
  paragraphs,
  paragraphText,
  toUint8Array,
  validate,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { editorFor } from "../index.js";
import { caretAt } from "../selection.js";
import { mergeBackCommand, splitParagraphCommand } from "./structure.js";
import { runCommand } from "./types.js";

describe("splitParagraphCommand", () => {
  it("splits the caret paragraph at the offset and round-trips", () => {
    const model = editorFor(createDocx({ paragraphs: ["Hello world"] }));
    model.setSelection(caretAt({ block: 0, inline: 0, offset: 5 }));
    runCommand(model, splitParagraphCommand, undefined);
    expect(paragraphs(model.doc).map(paragraphText)).toEqual(["Hello", " world"]);
    // Caret lands at the start of the new paragraph.
    expect(model.selection?.focus.block).toBe(1);

    const re = openDocx(toUint8Array(model.doc));
    expect(validate(re).length).toBe(0);
    expect(paragraphs(re).map(paragraphText)).toEqual(["Hello", " world"]);
  });
});

describe("mergeBackCommand", () => {
  it("merges the caret paragraph into the previous one and round-trips", () => {
    const model = editorFor(createDocx({ paragraphs: ["Hello", " world"] }));
    model.setSelection(caretAt({ block: 1, inline: 0, offset: 0 }));
    runCommand(model, mergeBackCommand, undefined);
    expect(paragraphs(model.doc).map(paragraphText)).toEqual(["Hello world"]);
    // Caret lands at the join point.
    expect(model.selection?.focus.block).toBe(0);
    expect(model.selection?.focus.offset).toBe(5);

    const re = openDocx(toUint8Array(model.doc));
    expect(validate(re).length).toBe(0);
    expect(paragraphs(re).map(paragraphText)).toEqual(["Hello world"]);
  });
});
