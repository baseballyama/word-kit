import { describe, expect, it } from "vitest";
import {
  appendTextRun,
  Docx,
  setParagraphAlignment,
  setParagraphIndent,
  setParagraphSpacing,
} from "./index.js";

describe("paragraph helpers", () => {
  it("appendTextRun adds a styled run to an existing paragraph", () => {
    const doc = Docx.create({ paragraphs: [] });
    const para = doc.appendParagraph("Hello ");
    appendTextRun(para, "world", { bold: true, color: "1F497D" });
    appendTextRun(para, "!", { italic: true, fontSizeHalfPoints: 28 });
    expect(para.children.length).toBe(3);
    const bytes = doc.toUint8Array();
    const reopened = Docx.open(bytes);
    const part = reopened.opc.getPart("/word/document.xml");
    const xml = new TextDecoder().decode(part?.data ?? new Uint8Array());
    expect(xml).toContain('w:val="1F497D"');
    expect(xml).toContain("<w:b/>");
    expect(xml).toContain("<w:i/>");
    expect(xml).toContain('w:val="28"');
  });

  it("setParagraphAlignment writes <w:jc>", () => {
    const doc = Docx.create({ paragraphs: [] });
    const para = doc.appendParagraph("centered");
    setParagraphAlignment(para, "center");
    const xml = new TextDecoder().decode(
      Docx.open(doc.toUint8Array()).opc.getPart("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain('<w:jc w:val="center"/>');
  });

  it("setParagraphIndent and setParagraphSpacing write the expected children", () => {
    const doc = Docx.create({ paragraphs: [] });
    const para = doc.appendParagraph("indented");
    setParagraphIndent(para, { left: 720, firstLine: 360 });
    setParagraphSpacing(para, { before: 240, after: 240, line: 360, lineRule: "auto" });
    const xml = new TextDecoder().decode(
      Docx.open(doc.toUint8Array()).opc.getPart("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain('w:left="720"');
    expect(xml).toContain('w:firstLine="360"');
    expect(xml).toContain('w:before="240"');
    expect(xml).toContain('w:line="360"');
    expect(xml).toContain('w:lineRule="auto"');
  });

  it("setting the same property twice replaces (not duplicates) it", () => {
    const doc = Docx.create({ paragraphs: [] });
    const para = doc.appendParagraph("alignment");
    setParagraphAlignment(para, "left");
    setParagraphAlignment(para, "right");
    const jcCount = para.pPr?.children.filter(
      (c) => c.kind === "element" && c.name.local === "jc",
    ).length;
    expect(jcCount).toBe(1);
  });
});
