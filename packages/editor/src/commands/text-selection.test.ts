/**
 * Selection-aware character formatting: the fix for "bolding a selection bolds
 * the whole line". Formatting must split runs at the selection boundaries and
 * apply only to the selected characters.
 */

import {
  createDocx,
  getRunFormat,
  openDocx,
  paragraphs,
  toUint8Array,
  validate,
  type WmlRun,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { editorFor } from "../index.js";
import type { EditorModel } from "../model.js";
import { toggleBoldCommand, toggleItalicCommand } from "./text.js";
import { runCommand } from "./types.js";

/** [text, bold] for every run of the first paragraph. */
function runProfile(model: EditorModel): Array<[string, boolean]> {
  const runs = paragraphs(model.doc)[0]!.children.filter((c): c is WmlRun => c.kind === "run");
  return runs.map((r) => [
    r.pieces
      .filter(
        (p): p is { kind: "text"; value: string; preserveSpace: boolean } => p.kind === "text",
      )
      .map((p) => p.value)
      .join(""),
    getRunFormat(r).bold === true,
  ]);
}

function select(model: EditorModel, startOffset: number, endOffset: number): void {
  model.setSelection({
    anchor: { block: 0, inline: 0, offset: startOffset },
    focus: { block: 0, inline: 0, offset: endOffset },
  });
}

describe("selection-aware bold", () => {
  it("bolds only the selected characters, not the whole line", () => {
    const model = editorFor(createDocx({ paragraphs: ["Hello world foo"] }));
    select(model, 6, 11); // "world"
    runCommand(model, toggleBoldCommand, undefined);
    expect(runProfile(model)).toEqual([
      ["Hello ", false],
      ["world", true],
      [" foo", false],
    ]);
  });

  it("bolds a selection at the very start", () => {
    const model = editorFor(createDocx({ paragraphs: ["Hello world"] }));
    select(model, 0, 5); // "Hello"
    runCommand(model, toggleBoldCommand, undefined);
    expect(runProfile(model)).toEqual([
      ["Hello", true],
      [" world", false],
    ]);
  });

  it("bolds a selection to the end", () => {
    const model = editorFor(createDocx({ paragraphs: ["Hello world"] }));
    select(model, 6, 11); // "world"
    runCommand(model, toggleBoldCommand, undefined);
    expect(runProfile(model)).toEqual([
      ["Hello ", false],
      ["world", true],
    ]);
  });

  it("keeps the same text selected after formatting", () => {
    const model = editorFor(createDocx({ paragraphs: ["Hello world foo"] }));
    select(model, 6, 11);
    runCommand(model, toggleBoldCommand, undefined);
    // Selection now targets the isolated "world" run.
    const sel = model.selection!;
    const runs = paragraphs(model.doc)[0]!.children.filter((c): c is WmlRun => c.kind === "run");
    const focusRun = runs[sel.focus.inline ?? 0]!;
    expect(focusRun.pieces.map((p) => (p.kind === "text" ? p.value : "")).join("")).toBe("world");
    expect(getRunFormat(focusRun).bold).toBe(true);
  });

  it("toggles bold off when the selection is already fully bold", () => {
    const model = editorFor(createDocx({ paragraphs: ["Hello world"] }));
    select(model, 0, 5);
    runCommand(model, toggleBoldCommand, undefined); // on
    select(model, 0, 5);
    runCommand(model, toggleBoldCommand, undefined); // off
    expect(runProfile(model).find(([t]) => t.startsWith("Hello"))?.[1]).toBe(false);
  });

  it("round-trips a partially-bolded paragraph", () => {
    const model = editorFor(createDocx({ paragraphs: ["Hello world foo"] }));
    select(model, 6, 11);
    runCommand(model, toggleBoldCommand, undefined);
    const re = openDocx(toUint8Array(model.doc));
    expect(validate(re).length).toBe(0);
    const runs = paragraphs(re)[0]!.children.filter((c): c is WmlRun => c.kind === "run");
    const bolded = runs.find((r) => getRunFormat(r).bold === true);
    expect(bolded?.pieces.map((p) => (p.kind === "text" ? p.value : "")).join("")).toBe("world");
  });

  it("applies independent formats to overlapping-then-adjacent selections", () => {
    const model = editorFor(createDocx({ paragraphs: ["abcdef"] }));
    select(model, 0, 3); // "abc" bold
    runCommand(model, toggleBoldCommand, undefined);
    // Re-select "def" and italicize.
    const runs = paragraphs(model.doc)[0]!.children.filter((c): c is WmlRun => c.kind === "run");
    // "def" starts at the second run, offset 0..3.
    model.setSelection({
      anchor: { block: 0, inline: runs.length - 1, offset: 0 },
      focus: { block: 0, inline: runs.length - 1, offset: 3 },
    });
    runCommand(model, toggleItalicCommand, undefined);
    const profile = paragraphs(model.doc)[0]!
      .children.filter((c): c is WmlRun => c.kind === "run")
      .map((r) => ({
        t: r.pieces.map((p) => (p.kind === "text" ? p.value : "")).join(""),
        b: getRunFormat(r).bold === true,
        i: getRunFormat(r).italic === true,
      }));
    expect(profile).toEqual([
      { t: "abc", b: true, i: false },
      { t: "def", b: false, i: true },
    ]);
  });
});
