import { describe, expect, it } from "vitest";
import {
  addTable,
  createDocx,
  openDocx,
  setTableCellShading,
  setTableRowAsHeader,
  tables,
  toUint8Array,
} from "./index.js";

describe("setTableCellShading", () => {
  it("writes <w:shd> on the target cell only", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [
      ["a", "b"],
      ["1", "2"],
    ]);
    const t = tables(doc)[0]!;
    const cell = t.rows[0]?.cells[0]!;
    setTableCellShading(cell, { fill: "E0E0E0" });
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    // <w:shd> is present with the expected fill.
    expect(xml).toMatch(/<w:shd[^/]*w:fill="E0E0E0"/);
  });

  it("supports an explicit pattern + color", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["x"]]);
    const t = tables(doc)[0]!;
    const cell = t.rows[0]?.cells[0]!;
    setTableCellShading(cell, { fill: "FFFFFF", pattern: "horzStripe", color: "1F497D" });
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toMatch(/<w:shd[^/]*w:val="horzStripe"[^/]*w:color="1F497D"[^/]*w:fill="FFFFFF"/);
  });

  it("replaces an existing <w:shd> rather than appending a second", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["x"]]);
    const t = tables(doc)[0]!;
    const cell = t.rows[0]?.cells[0]!;
    setTableCellShading(cell, { fill: "AAAAAA" });
    setTableCellShading(cell, { fill: "BBBBBB" });
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toMatch(/w:fill="BBBBBB"/);
    expect(xml).not.toMatch(/w:fill="AAAAAA"/);
  });
});

describe("setTableRowAsHeader", () => {
  it("adds <w:tblHeader/> to the row's trPr", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [
      ["h1", "h2"],
      ["1", "2"],
    ]);
    const t = tables(doc)[0]!;
    const row0 = t.rows[0]!;
    setTableRowAsHeader(row0);
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain("<w:tblHeader/>");
  });

  it("clears the marker when called with false", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["x"]]);
    const t = tables(doc)[0]!;
    const row0 = t.rows[0]!;
    setTableRowAsHeader(row0);
    setTableRowAsHeader(row0, false);
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).not.toContain("<w:tblHeader/>");
  });
});
