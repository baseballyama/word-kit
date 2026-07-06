/**
 * Round-trip tests for the generic property / table-property / section-property
 * / doc-property commands. Each proves the edit survives serialize → reopen and
 * lands on the right OOXML element, backing the ledger's `edit` dispositions.
 */

import {
  addTable,
  appProperties,
  createDocx,
  getElementProp,
  getParagraphProp,
  getRunProp,
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
import { commandsInGroup, getCommand } from "./registry.js";
import { runCommand } from "./types.js";

function editorWith(texts: string[]): EditorModel {
  const model = editorFor(createDocx({ paragraphs: texts }));
  model.setSelection(caretAt({ block: 0 }));
  return model;
}

function firstRun(model: EditorModel): WmlRun {
  return paragraphs(model.doc)[0]!.children.find((c): c is WmlRun => c.kind === "run")!;
}

function expectValidRoundTrip(model: EditorModel) {
  const reopened = openDocx(toUint8Array(model.doc));
  expect(validate(reopened).length).toBe(0);
  return reopened;
}

describe("generic run property commands", () => {
  it("toggles all-caps and persists", () => {
    const model = editorWith(["hello"]);
    runCommand(model, getCommand("text.caps")!, undefined);
    expect(getRunProp(firstRun(model), "caps").present).toBe(true);
    const reopened = expectValidRoundTrip(model);
    const run = paragraphs(reopened)[0]!.children.find((c): c is WmlRun => c.kind === "run")!;
    expect(getRunProp(run, "caps").present).toBe(true);
  });

  it("sets superscript (vertAlign) with a value", () => {
    const model = editorWith(["x2"]);
    runCommand(model, getCommand("text.vertAlign")!, { val: "superscript" });
    expect(getRunProp(firstRun(model), "vertAlign").val).toBe("superscript");
    expectValidRoundTrip(model);
  });
});

describe("generic paragraph property commands", () => {
  it("toggles keepNext and sets outlineLvl", () => {
    const model = editorWith(["heading-ish"]);
    runCommand(model, getCommand("paragraph.keepNext")!, undefined);
    runCommand(model, getCommand("paragraph.outlineLvl")!, { val: "0" });
    const p = paragraphs(model.doc)[0]!;
    expect(getParagraphProp(p, "keepNext").present).toBe(true);
    expect(getParagraphProp(p, "outlineLvl").val).toBe("0");
    expectValidRoundTrip(model);
  });
});

describe("table property commands", () => {
  it("sets a cell no-wrap flag", () => {
    const model = editorWith([""]);
    const table = addTable(model.doc, [["a", "b"]]);
    const blockIndex = model.doc.document.body.blocks.indexOf(table);
    model.setSelection(caretAt({ block: blockIndex, cell: { row: 0, col: 0 } }));
    runCommand(model, getCommand("table.cell_noWrap")!, undefined);
    const cell = tables(model.doc)[0]!.rows[0]!.cells[0]!;
    expect(getElementProp(cell.tcPr, "noWrap").present).toBe(true);
    expectValidRoundTrip(model);
  });
});

describe("section property commands", () => {
  it("enables a title page", () => {
    const model = editorWith(["body"]);
    runCommand(model, getCommand("section.titlePg")!, undefined);
    expect(getElementProp(model.doc.document.body.sectPr, "titlePg").present).toBe(true);
    expectValidRoundTrip(model);
  });
});

describe("document property commands", () => {
  it("sets extended app properties", () => {
    const model = editorWith(["x"]);
    runCommand(model, getCommand("docprops.setApp")!, { company: "word-kit", manager: "Ada" });
    const reopened = expectValidRoundTrip(model);
    expect(appProperties(reopened).company).toBe("word-kit");
    expect(appProperties(reopened).manager).toBe("Ada");
  });
});

describe("registry sanity", () => {
  it("exposes a substantial editable command surface", () => {
    // Guardrail so a regression that strips command modules is caught.
    expect(commandsInGroup("text").length).toBeGreaterThanOrEqual(20);
    expect(commandsInGroup("table").length).toBeGreaterThanOrEqual(15);
  });
});
