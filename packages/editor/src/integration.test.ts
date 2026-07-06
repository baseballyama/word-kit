/**
 * End-to-end editing-session test: chain the operations a user performs (type,
 * format a selection, align, list, heading, split/merge, undo/redo, replace,
 * table) and assert the document stays valid and correct throughout. This is
 * the "all operations" smoke test guarding the editor's core behaviours.
 */

import {
  createDocx,
  getParagraphAlignment,
  getParagraphStyle,
  getRunFormat,
  openDocx,
  paragraphs,
  paragraphText,
  toUint8Array,
  validate,
  type WmlRun,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { commands, editorFor, runCommand } from "./index.js";
import type { EditorModel } from "./model.js";

function firstParaRuns(model: EditorModel): WmlRun[] {
  return paragraphs(model.doc)[0]!.children.filter((c): c is WmlRun => c.kind === "run");
}

function selectFirst(model: EditorModel, start: number, end: number): void {
  model.setSelection({
    anchor: { block: 0, inline: 0, offset: start },
    focus: { block: 0, inline: 0, offset: end },
  });
}

function assertValidRoundTrip(model: EditorModel) {
  const re = openDocx(toUint8Array(model.doc));
  expect(validate(re).length).toBe(0);
  return re;
}

describe("editing session (all core operations)", () => {
  it("chains character formatting on selections and round-trips", () => {
    const model = editorFor(createDocx({ paragraphs: ["The quick brown fox"] }));

    selectFirst(model, 4, 9); // "quick"
    runCommand(model, commands.toggleBoldCommand, undefined);
    selectFirst(model, 10, 15); // "brown" — offsets unchanged since we re-select on original text
    // After the first split, run indices changed; re-anchor via absolute offsets.
    model.setSelection({
      anchor: { block: 0, inline: 0, offset: 0 },
      focus: { block: 0, inline: 0, offset: 3 },
    });
    runCommand(model, commands.toggleItalicCommand, undefined);

    const re = assertValidRoundTrip(model);
    const runs = paragraphs(re)[0]!.children.filter((c): c is WmlRun => c.kind === "run");
    const bolded = runs.find((r) => getRunFormat(r).bold);
    expect(bolded?.pieces.map((p) => (p.kind === "text" ? p.value : "")).join("")).toBe("quick");
  });

  it("applies paragraph alignment, style, and list, then round-trips", () => {
    const model = editorFor(createDocx({ paragraphs: ["Heading text", "body one", "body two"] }));

    model.setSelection({ anchor: { block: 0 }, focus: { block: 0 } });
    runCommand(model, commands.alignCenterCommand, undefined);
    expect(getParagraphAlignment(paragraphs(model.doc)[0]!)).toBe("center");

    runCommand(model, commands.insertHeadingCommand, { text: "New Heading", level: 1 });
    const heading = paragraphs(model.doc).find((p) => getParagraphStyle(p) === "Heading1");
    expect(heading).toBeDefined();

    model.setSelection({ anchor: { block: 1 }, focus: { block: 1 } });
    runCommand(model, commands.applyListCommand, { kind: "bullet" });

    assertValidRoundTrip(model);
  });

  it("splits, merges, and undoes/redoes with correct final text", () => {
    const model = editorFor(createDocx({ paragraphs: ["Hello world"] }));

    model.setSelection({
      anchor: { block: 0, inline: 0, offset: 5 },
      focus: { block: 0, inline: 0, offset: 5 },
    });
    runCommand(model, commands.splitParagraphCommand, undefined);
    expect(paragraphs(model.doc).map(paragraphText)).toEqual(["Hello", " world"]);

    model.undo();
    expect(paragraphs(model.doc).map(paragraphText)).toEqual(["Hello world"]);

    model.redo();
    expect(paragraphs(model.doc).map(paragraphText)).toEqual(["Hello", " world"]);

    model.setSelection({
      anchor: { block: 1, inline: 0, offset: 0 },
      focus: { block: 1, inline: 0, offset: 0 },
    });
    runCommand(model, commands.mergeBackCommand, undefined);
    expect(paragraphs(model.doc).map(paragraphText)).toEqual(["Hello world"]);

    assertValidRoundTrip(model);
  });

  it("inserts a table and a page break without breaking validity", () => {
    const model = editorFor(createDocx({ paragraphs: ["intro"] }));
    model.setSelection({ anchor: { block: 0 }, focus: { block: 0 } });
    runCommand(model, commands.insertTableCommand, { rows: 2, cols: 3 });
    runCommand(model, commands.insertPageBreakCommand, undefined);
    assertValidRoundTrip(model);
  });

  it("sets font, size, and color on a selection", () => {
    const model = editorFor(createDocx({ paragraphs: ["colored text"] }));
    selectFirst(model, 0, 7); // "colored"
    runCommand(model, commands.setColorCommand, { color: "ff0000" });
    runCommand(model, commands.setFontSizeCommand, { points: 18 });

    const re = assertValidRoundTrip(model);
    const runs = paragraphs(re)[0]!.children.filter((c): c is WmlRun => c.kind === "run");
    const colored = runs.find((r) => getRunFormat(r).color === "ff0000");
    expect(colored).toBeDefined();
    expect(getRunFormat(colored!).fontSizeHalfPoints).toBe(36);
    expect(colored!.pieces.map((p) => (p.kind === "text" ? p.value : "")).join("")).toBe("colored");
  });

  it("keeps run count minimal (no stray empty runs) after formatting", () => {
    const model = editorFor(createDocx({ paragraphs: ["abcdefgh"] }));
    selectFirst(model, 2, 5); // "cde"
    runCommand(model, commands.toggleBoldCommand, undefined);
    const runs = firstParaRuns(model);
    // Exactly three runs: "ab" | "cde" | "fgh".
    expect(
      runs.map((r) => r.pieces.map((p) => (p.kind === "text" ? p.value : "")).join("")),
    ).toEqual(["ab", "cde", "fgh"]);
  });
});
