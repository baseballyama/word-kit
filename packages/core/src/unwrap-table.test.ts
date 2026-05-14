import { describe, expect, it } from "vitest";
import {
  addTable,
  appendParagraph,
  createDocx,
  openDocx,
  paragraphs,
  paragraphText,
  tables,
  toUint8Array,
  unwrapTable,
} from "./index.js";

describe("unwrapTable", () => {
  it("replaces a table with one paragraph per cell in row-major order", () => {
    const doc = createDocx({ paragraphs: ["before"] });
    addTable(doc, [
      ["A", "B"],
      ["1", "2"],
    ]);
    appendParagraph(doc, "after");

    const inserted = unwrapTable(doc, 0);
    expect(inserted).toBeDefined();
    expect(inserted!.map((p) => paragraphText(p))).toEqual(["A", "B", "1", "2"]);
    expect(tables(doc)).toHaveLength(0);
    const allParaTexts = paragraphs(doc).map((p) => paragraphText(p));
    expect(allParaTexts.slice(0, 1)).toEqual(["before"]);
    expect(allParaTexts.slice(-1)).toEqual(["after"]);
    expect(allParaTexts).toContain("A");
    expect(allParaTexts).toContain("2");
  });

  it("returns undefined when the index is out of range", () => {
    const doc = createDocx({ paragraphs: [] });
    expect(unwrapTable(doc, 0)).toBeUndefined();
    addTable(doc, [["only"]]);
    expect(unwrapTable(doc, 7)).toBeUndefined();
  });

  it("indexes only across tables — paragraphs between tables don't shift the count", () => {
    const doc = createDocx({ paragraphs: ["intro"] });
    addTable(doc, [["t0"]]);
    appendParagraph(doc, "mid");
    addTable(doc, [["t1"]]);

    expect(unwrapTable(doc, 1)).toBeDefined();
    expect(tables(doc)).toHaveLength(1);
    // First table (t0) is still there; t1 is unwrapped.
    const t = tables(doc)[0];
    const cell0 = t?.rows[0]?.cells[0]?.paragraphs[0];
    expect(cell0 && paragraphText(cell0)).toBe("t0");
    expect(paragraphs(doc).map((p) => paragraphText(p))).toContain("t1");
  });

  it("survives save+reopen with the formerly-tabled cells as plain paragraphs", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addTable(doc, [
      ["A", "B"],
      ["C", "D"],
    ]);
    unwrapTable(doc, 0);

    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).not.toContain("<w:tbl");
    const texts = paragraphs(reopened).map((p) => paragraphText(p));
    expect(texts).toEqual(expect.arrayContaining(["A", "B", "C", "D"]));
  });
});
