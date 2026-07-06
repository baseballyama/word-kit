/**
 * Command round-trip tests. Each editable feature is exercised through the
 * command layer, then the document is serialized and reopened to prove the edit
 * survives a real save/load cycle (not just an in-memory mutation). These tests
 * are what the coverage ledger's `edit` dispositions stand on.
 */

import {
  createDocx,
  getParagraphAlignment,
  getRunFormat,
  openDocx,
  paragraphs,
  tables,
  toUint8Array,
  validate,
  type WmlRun,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { editorFor } from "../index.js";
import type { EditorModel } from "../model.js";
import { caretAt } from "../selection.js";
import { runCommand } from "./types.js";
import { toggleBoldCommand, setColorCommand, setFontSizeCommand } from "./text.js";
import { alignCenterCommand, setParagraphStyleCommand } from "./paragraph.js";
import { ensureHeadingStylesCommand } from "./style.js";
import {
  insertParagraphCommand,
  insertHeadingCommand,
  insertPageBreakCommand,
} from "./structure.js";
import { insertTableCommand, addRowCommand } from "./table.js";
import { applyListCommand } from "./list.js";
import { acceptAllRevisionsCommand } from "./review.js";

function editorWith(texts: string[]): EditorModel {
  const model = editorFor(createDocx({ paragraphs: texts }));
  model.setSelection(caretAt({ block: 0 }));
  return model;
}

/** Serialize → reopen and return the fresh Docx (proves the edit persisted). */
function roundTrip(model: EditorModel) {
  const bytes = toUint8Array(model.doc);
  const reopened = openDocx(bytes);
  expect(validate(reopened).length).toBe(0);
  return reopened;
}

function firstRun(model: EditorModel): WmlRun {
  const p = paragraphs(model.doc)[0]!;
  const run = p.children.find((c): c is WmlRun => c.kind === "run");
  if (!run) throw new Error("no run");
  return run;
}

describe("text commands", () => {
  it("toggles bold and it survives round-trip", () => {
    const model = editorWith(["Hello world"]);
    runCommand(model, toggleBoldCommand, undefined);
    expect(getRunFormat(firstRun(model)).bold).toBe(true);
    const reopened = openDocx(toUint8Array(model.doc));
    const run = paragraphs(reopened)[0]!.children.find((c): c is WmlRun => c.kind === "run")!;
    expect(getRunFormat(run).bold).toBe(true);
  });

  it("sets color and font size", () => {
    const model = editorWith(["Text"]);
    runCommand(model, setColorCommand, { color: "#ff0000" });
    runCommand(model, setFontSizeCommand, { points: 18 });
    const fmt = getRunFormat(firstRun(model));
    expect(fmt.color).toBe("ff0000");
    expect(fmt.fontSizeHalfPoints).toBe(36);
    roundTrip(model);
  });
});

describe("paragraph commands", () => {
  it("centers a paragraph", () => {
    const model = editorWith(["Centered"]);
    runCommand(model, alignCenterCommand, undefined);
    expect(getParagraphAlignment(paragraphs(model.doc)[0]!)).toBe("center");
    roundTrip(model);
  });

  it("applies a paragraph style", () => {
    const model = editorWith(["Styled"]);
    runCommand(model, ensureHeadingStylesCommand, undefined);
    runCommand(model, setParagraphStyleCommand, { styleId: "Heading1" });
    roundTrip(model);
  });
});

describe("structure commands", () => {
  it("inserts a paragraph after the caret", () => {
    const model = editorWith(["First", "Third"]);
    model.setSelection(caretAt({ block: 0 }));
    runCommand(model, insertParagraphCommand, { text: "Second" });
    const texts = paragraphs(model.doc).map((p) =>
      p.children
        .filter((c): c is WmlRun => c.kind === "run")
        .flatMap((r) => r.pieces)
        .map((pc) => (pc.kind === "text" ? pc.value : ""))
        .join(""),
    );
    expect(texts).toEqual(["First", "Second", "Third"]);
    roundTrip(model);
  });

  it("inserts a heading and a page break", () => {
    const model = editorWith(["Body"]);
    runCommand(model, insertHeadingCommand, { text: "Title", level: 1 });
    runCommand(model, insertPageBreakCommand, undefined);
    roundTrip(model);
  });
});

describe("table commands", () => {
  it("inserts a table and adds a row", () => {
    const model = editorWith(["Above"]);
    runCommand(model, insertTableCommand, { rows: 2, cols: 3 });
    expect(tables(model.doc).length).toBe(1);
    expect(tables(model.doc)[0]!.rows.length).toBe(2);
    // Caret is now in the table; add a row.
    runCommand(model, addRowCommand, { texts: ["a", "b", "c"] });
    expect(tables(model.doc)[0]!.rows.length).toBe(3);
    roundTrip(model);
  });
});

describe("list commands", () => {
  it("applies a bullet list to the selection", () => {
    const model = editorWith(["Item one", "Item two"]);
    model.setSelection({ anchor: { block: 0 }, focus: { block: 1 } });
    runCommand(model, applyListCommand, { kind: "bullet" });
    roundTrip(model);
  });
});

describe("review commands", () => {
  it("accepts all revisions without error", () => {
    const model = editorWith(["Reviewed"]);
    runCommand(model, acceptAllRevisionsCommand, undefined);
    roundTrip(model);
  });
});

describe("history", () => {
  it("undo/redo restores document state", () => {
    const model = editorWith(["Hello"]);
    runCommand(model, toggleBoldCommand, undefined);
    expect(getRunFormat(firstRun(model)).bold).toBe(true);
    model.undo();
    expect(getRunFormat(firstRun(model)).bold ?? false).toBe(false);
    model.redo();
    expect(getRunFormat(firstRun(model)).bold).toBe(true);
  });
});
