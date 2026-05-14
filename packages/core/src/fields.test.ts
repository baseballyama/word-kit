import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.addPageNumberFooter", () => {
  it("creates a footer with a PAGE field", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    doc.addPageNumberFooter();
    const bytes = doc.toUint8Array();
    const reopened = Docx.open(bytes);
    const footerXml = new TextDecoder().decode(
      reopened.opc.getPart("/word/footer1.xml")?.data ?? new Uint8Array(),
    );
    expect(footerXml).toContain('<w:fldChar w:fldCharType="begin"/>');
    expect(footerXml).toContain("<w:instrText>PAGE</w:instrText>");
    expect(footerXml).toContain('<w:fldChar w:fldCharType="end"/>');
  });

  it("supports prefix and suffix text around the page number", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    doc.addPageNumberFooter("Page ", " of N");
    const footerXml = new TextDecoder().decode(
      Docx.open(doc.toUint8Array()).opc.getPart("/word/footer1.xml")?.data ?? new Uint8Array(),
    );
    expect(footerXml).toContain("Page ");
    expect(footerXml).toContain(" of N");
  });
});

describe("Docx.appendField", () => {
  it("appends fldChar/instrText/fldChar sequence to a paragraph", () => {
    const doc = Docx.create({ paragraphs: ["Today: "] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.appendField(para, "DATE", "2026-05-15");
    const reopened = Docx.open(doc.toUint8Array());
    const xml = new TextDecoder().decode(
      reopened.opc.getPart("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain('<w:fldChar w:fldCharType="begin"/>');
    expect(xml).toContain("<w:instrText>DATE</w:instrText>");
    expect(xml).toContain('<w:fldChar w:fldCharType="separate"/>');
    expect(xml).toContain("<w:t>2026-05-15</w:t>");
    expect(xml).toContain('<w:fldChar w:fldCharType="end"/>');
  });

  it("works with MERGEFIELD for template fields", () => {
    const doc = Docx.create({ paragraphs: ["Hello "] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.appendField(para, "MERGEFIELD name \\* MERGEFORMAT", "«name»");
    const xml = new TextDecoder().decode(
      Docx.open(doc.toUint8Array()).opc.getPart("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain("MERGEFIELD name");
    expect(xml).toContain("«name»");
  });
});
