import { describe, expect, it } from "vitest";
import {
  Docx,
  getParagraphAlignment,
  getParagraphStyle,
  PAGE_SIZE_A4,
  setParagraphText,
} from "./index.js";

describe("Docx.appendSectionBreak", () => {
  it("inserts a section break paragraph with sectPr in pPr", () => {
    const doc = Docx.create({ paragraphs: ["section one"] });
    doc.appendSectionBreak("nextPage");
    doc.appendParagraph("section two");
    const bytes = doc.toUint8Array();
    const reopened = Docx.open(bytes);
    const part = reopened.opc.getPart("/word/document.xml");
    const xml = new TextDecoder().decode(part?.data ?? new Uint8Array());
    expect(xml).toContain('<w:type w:val="nextPage"/>');
  });

  it("can override page size and margins for the next section", () => {
    const doc = Docx.create({ paragraphs: ["portrait section"] });
    doc.appendSectionBreak("nextPage", {
      pageSize: { ...PAGE_SIZE_A4, orientation: "landscape" },
    });
    doc.appendParagraph("landscape section");
    const xml = new TextDecoder().decode(
      Docx.open(doc.toUint8Array()).opc.getPart("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain('w:orient="landscape"');
  });
});

describe("Docx.removeParagraph / clearBody", () => {
  it("removes the paragraph at the given index (paragraph-relative)", () => {
    const doc = Docx.create({ paragraphs: ["one", "two", "three"] });
    expect(doc.removeParagraph(1)).toBe(true);
    expect(doc.paragraphs.map((p) => p.children.length).length).toBe(2);
    expect(doc.text).toBe("one\nthree");
  });

  it("returns false when the index is out of range", () => {
    const doc = Docx.create({ paragraphs: ["one"] });
    expect(doc.removeParagraph(5)).toBe(false);
  });

  it("clearBody empties the body", () => {
    const doc = Docx.create({ paragraphs: ["one", "two"] });
    doc.clearBody();
    expect(doc.paragraphs).toHaveLength(0);
    expect(doc.text).toBe("");
  });
});

describe("Docx.removeAllComments", () => {
  it("clears comments.xml entries and strips the markers from the body", () => {
    const doc = Docx.create({ paragraphs: ["paragraph 1"] });
    const p = doc.paragraphs[0];
    if (!p) return;
    doc.addComment(p, { author: "R", text: "First" });
    doc.addComment(p, { author: "R", text: "Second" });
    expect(doc.commentsPart?.comments).toHaveLength(2);
    expect(doc.removeAllComments()).toBe(2);
    expect(doc.commentsPart?.comments).toHaveLength(0);
    // No more rangeStart/rangeEnd or commentReference in the paragraph.
    const bytes = doc.toUint8Array();
    const xml = new TextDecoder().decode(
      Docx.open(bytes).opc.getPart("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).not.toContain("commentRangeStart");
    expect(xml).not.toContain("commentReference");
  });
});

describe("paragraph property accessors", () => {
  it("getParagraphStyle / getParagraphAlignment read pPr children", () => {
    const doc = Docx.create({ paragraphs: [] });
    const para = doc.appendParagraph("hi", { style: "Heading1" });
    expect(getParagraphStyle(para)).toBe("Heading1");
    expect(getParagraphAlignment(para)).toBeUndefined();
  });

  it("setParagraphText replaces existing runs with a single styled run", () => {
    const doc = Docx.create({ paragraphs: ["original"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    setParagraphText(para, "rewritten", { bold: true });
    expect(doc.text).toBe("rewritten");
    expect(para.children).toHaveLength(1);
  });
});
