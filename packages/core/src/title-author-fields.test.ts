import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.title / Docx.author shortcuts", () => {
  it("title getter+setter proxies coreProperties.title", () => {
    const doc = Docx.create();
    expect(doc.title).toBeUndefined();
    doc.title = "Hello";
    expect(doc.title).toBe("Hello");
    expect(doc.coreProperties.title).toBe("Hello");
  });

  it("author getter+setter proxies coreProperties.creator", () => {
    const doc = Docx.create();
    doc.author = "Yamada";
    expect(doc.author).toBe("Yamada");
    expect(doc.coreProperties.creator).toBe("Yamada");
  });

  it("merging title and author leaves the other untouched", () => {
    const doc = Docx.create();
    doc.title = "T";
    doc.author = "A";
    expect(doc.coreProperties).toMatchObject({ title: "T", creator: "A" });
  });
});

describe("Docx.fields enumeration", () => {
  it("lists every <w:instrText> instruction in the body", () => {
    const doc = Docx.create({ paragraphs: ["Today: "] });
    const p1 = doc.paragraphs[0];
    if (!p1) return;
    doc.appendField(p1, "DATE", "2026-05-15");
    doc.appendParagraph("Page ");
    const p2 = doc.paragraphs[1];
    if (!p2) return;
    doc.appendField(p2, "PAGE", "1");
    const fields = doc.fields;
    expect(fields.map((f) => f.instruction)).toEqual(["DATE", "PAGE"]);
  });

  it("finds fields inside table cells too", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addTable([["cell"]]);
    const table = doc.tables[0];
    expect(table).toBeDefined();
    if (!table) return;
    const cellPara = table.rows[0]?.cells[0]?.paragraphs[0];
    expect(cellPara).toBeDefined();
    if (!cellPara) return;
    doc.appendField(cellPara, "AUTHOR", "Yamada");
    expect(doc.fields.map((f) => f.instruction)).toContain("AUTHOR");
  });
});
