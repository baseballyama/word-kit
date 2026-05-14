import { describe, expect, it } from "vitest";
import {
  addTable,
  createDocx,
  openDocx,
  setTableCellVerticalAlign,
  setTableRowHeight,
  tables,
  toUint8Array,
} from "./index.js";

describe("setTableCellVerticalAlign", () => {
  it('writes <w:vAlign w:val="center"> on the cell\'s tcPr', () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["x"]]);
    const t = tables(doc)[0]!;
    const cell = t.rows[0]?.cells[0]!;
    setTableCellVerticalAlign(cell, "center");
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toMatch(/<w:vAlign w:val="center"\/>/);
  });

  it("replaces an existing vAlign rather than appending a second", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["x"]]);
    const t = tables(doc)[0]!;
    const cell = t.rows[0]?.cells[0]!;
    setTableCellVerticalAlign(cell, "top");
    setTableCellVerticalAlign(cell, "bottom");
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    const occurrences = (xml.match(/<w:vAlign/g) || []).length;
    expect(occurrences).toBe(1);
    expect(xml).toMatch(/<w:vAlign w:val="bottom"\/>/);
  });
});

describe("setTableRowHeight", () => {
  it("writes <w:trHeight> with the supplied twips + default atLeast rule", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["a"], ["b"]]);
    const t = tables(doc)[0]!;
    setTableRowHeight(t.rows[0]!, 600);
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toMatch(/<w:trHeight w:val="600" w:hRule="atLeast"\/>/);
  });

  it("supports an explicit exact rule", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["a"]]);
    const t = tables(doc)[0]!;
    setTableRowHeight(t.rows[0]!, 480, "exact");
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toMatch(/<w:trHeight w:val="480" w:hRule="exact"\/>/);
  });

  it("replaces an existing trHeight on the same row", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["a"]]);
    const t = tables(doc)[0]!;
    setTableRowHeight(t.rows[0]!, 200);
    setTableRowHeight(t.rows[0]!, 800);
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    const occurrences = (xml.match(/<w:trHeight/g) || []).length;
    expect(occurrences).toBe(1);
    expect(xml).toMatch(/w:val="800"/);
  });
});
