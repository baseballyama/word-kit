import { getPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  addPageNumberFooter,
  appendField,
  createDocx,
  openDocx,
  paragraphs,
  toUint8Array,
} from "./docx.js";

describe("Docx.addPageNumberFooter", () => {
  it("creates a footer with a PAGE field", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addPageNumberFooter(doc);
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    const footerXml = new TextDecoder().decode(
      getPart(reopened.opc, "/word/footer1.xml")?.data ?? new Uint8Array(),
    );
    expect(footerXml).toContain('<w:fldChar w:fldCharType="begin"/>');
    expect(footerXml).toContain("<w:instrText>PAGE</w:instrText>");
    expect(footerXml).toContain('<w:fldChar w:fldCharType="end"/>');
  });

  it("supports prefix and suffix text around the page number", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addPageNumberFooter(doc, "Page ", " of N");
    const footerXml = new TextDecoder().decode(
      getPart(openDocx(toUint8Array(doc)).opc, "/word/footer1.xml")?.data ?? new Uint8Array(),
    );
    expect(footerXml).toContain("Page ");
    expect(footerXml).toContain(" of N");
  });
});

describe("Docx.appendField", () => {
  it("appends fldChar/instrText/fldChar sequence to a paragraph", () => {
    const doc = createDocx({ paragraphs: ["Today: "] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    appendField(doc, para, "DATE", "2026-05-15");
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      getPart(reopened.opc, "/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain('<w:fldChar w:fldCharType="begin"/>');
    expect(xml).toContain("<w:instrText>DATE</w:instrText>");
    expect(xml).toContain('<w:fldChar w:fldCharType="separate"/>');
    expect(xml).toContain("<w:t>2026-05-15</w:t>");
    expect(xml).toContain('<w:fldChar w:fldCharType="end"/>');
  });

  it("works with MERGEFIELD for template fields", () => {
    const doc = createDocx({ paragraphs: ["Hello "] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    appendField(doc, para, "MERGEFIELD name \\* MERGEFORMAT", "«name»");
    const xml = new TextDecoder().decode(
      getPart(openDocx(toUint8Array(doc)).opc, "/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain("MERGEFIELD name");
    expect(xml).toContain("«name»");
  });
});
