import { getPart, writeOpcPackage } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.addTable", () => {
  it("appends a table with the requested cell text", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addTable([
      ["Name", "Score"],
      ["Alice", "90"],
      ["Bob", "85"],
    ]);
    expect(doc.tables).toHaveLength(1);
    const t = doc.tables[0];
    expect(t).toBeDefined();
    expect(t?.rows).toHaveLength(3);
    expect(t?.rows[0]?.cells).toHaveLength(2);
    const firstCell = t?.rows[0]?.cells[0];
    expect(firstCell?.paragraphs[0]?.children[0]).toMatchObject({ kind: "run" });
  });

  it("survives a save+reopen round-trip", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addTable([
      ["A", "B", "C"],
      ["1", "2", "3"],
    ]);
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.tables).toHaveLength(1);
    const t = reopened.tables[0];
    expect(t?.rows).toHaveLength(2);
    expect(t?.rows[0]?.cells).toHaveLength(3);
  });

  it("pads short rows with empty cells", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addTable([["a", "b", "c"], ["x"]]);
    const t = doc.tables[0];
    expect(t?.rows[1]?.cells).toHaveLength(3);
  });

  it("the body keeps paragraphs and tables in the order they were added", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.appendParagraph("before");
    doc.addTable([["x"]]);
    doc.appendParagraph("after");
    const kinds = doc.document.body.blocks.map((b) => b.kind);
    expect(kinds).toEqual(["paragraph", "table", "paragraph"]);
  });

  it("preserves an existing template table through open and re-save", () => {
    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      "<w:body>" +
      "<w:tbl>" +
      '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>' +
      "<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
      "<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>" +
      "</w:tbl>" +
      "</w:body>" +
      "</w:document>";
    // Embed the template into a fresh package so we can open it.
    const seed = Docx.create({ paragraphs: [] });
    const docPart = getPart(seed.opc, "/word/document.xml");
    if (!docPart) throw new Error("no document part");
    docPart.data = new TextEncoder().encode(xml);
    const bytes = writeOpcPackage(seed.opc);
    const reopened = Docx.open(bytes);
    expect(reopened.tables).toHaveLength(1);
    const t = reopened.tables[0];
    expect(t?.rows[0]?.cells[0]?.paragraphs[0]).toBeDefined();
  });
});
