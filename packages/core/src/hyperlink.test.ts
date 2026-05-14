import { getPart, partRelationships, relationshipsByType } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.addHyperlink", () => {
  it("appends a paragraph and registers an external relationship", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addHyperlink("https://example.com/", "Visit example");
    expect(doc.paragraphs).toHaveLength(1);
    const rels = partRelationships(doc.opc, "/word/document.xml");
    const linkRels = relationshipsByType(
      rels,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    );
    expect(linkRels).toHaveLength(1);
    expect(linkRels[0]?.target).toBe("https://example.com/");
    expect(linkRels[0]?.targetMode).toBe("External");
  });

  it("survives save+reopen with hyperlink and rel intact", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addHyperlink("https://example.com/", "Click");
    const reopened = Docx.open(doc.toUint8Array());
    const rels = partRelationships(reopened.opc, "/word/document.xml");
    expect(
      relationshipsByType(
        rels,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      ),
    ).toHaveLength(1);
    const part = getPart(reopened.opc, "/word/document.xml");
    const xml = new TextDecoder().decode(part?.data ?? new Uint8Array());
    expect(xml).toContain("Click");
  });

  it("optional tooltip is preserved", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addHyperlink("https://example.com/", "x", { tooltip: "Hello" });
    const reopened = Docx.open(doc.toUint8Array());
    const part = getPart(reopened.opc, "/word/document.xml");
    const xml = new TextDecoder().decode(part?.data ?? new Uint8Array());
    expect(xml).toContain('w:tooltip="Hello"');
  });
});
