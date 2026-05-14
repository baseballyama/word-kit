import { describe, expect, it } from "vitest";
import {
  addTable,
  appendField,
  appendParagraph,
  author,
  coreProperties,
  createDocx,
  fields,
  paragraphs,
  setAuthor,
  setTitle,
  tables,
  title,
} from "./docx.js";

describe("Docx.title / Docx.author shortcuts", () => {
  it("title getter+setter proxies coreProperties.title", () => {
    const doc = createDocx();
    expect(title(doc)).toBeUndefined();
    setTitle(doc, "Hello");
    expect(title(doc)).toBe("Hello");
    expect(coreProperties(doc).title).toBe("Hello");
  });

  it("author getter+setter proxies coreProperties.creator", () => {
    const doc = createDocx();
    setAuthor(doc, "Yamada");
    expect(author(doc)).toBe("Yamada");
    expect(coreProperties(doc).creator).toBe("Yamada");
  });

  it("merging title and author leaves the other untouched", () => {
    const doc = createDocx();
    setTitle(doc, "T");
    setAuthor(doc, "A");
    expect(coreProperties(doc)).toMatchObject({ title: "T", creator: "A" });
  });
});

describe("Docx.fields enumeration", () => {
  it("lists every <w:instrText> instruction in the body", () => {
    const doc = createDocx({ paragraphs: ["Today: "] });
    const p1 = paragraphs(doc)[0];
    if (!p1) return;
    appendField(doc, p1, "DATE", "2026-05-15");
    appendParagraph(doc, "Page ");
    const p2 = paragraphs(doc)[1];
    if (!p2) return;
    appendField(doc, p2, "PAGE", "1");
    const flds = fields(doc);
    expect(flds.map((f) => f.instruction)).toEqual(["DATE", "PAGE"]);
  });

  it("finds fields inside table cells too", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["cell"]]);
    const table = tables(doc)[0];
    expect(table).toBeDefined();
    if (!table) return;
    const cellPara = table.rows[0]?.cells[0]?.paragraphs[0];
    expect(cellPara).toBeDefined();
    if (!cellPara) return;
    appendField(doc, cellPara, "AUTHOR", "Yamada");
    expect(fields(doc).map((f) => f.instruction)).toContain("AUTHOR");
  });
});
