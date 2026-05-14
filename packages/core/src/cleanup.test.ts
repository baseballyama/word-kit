import { getPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  getParagraphAlignment,
  getParagraphStyle,
  PAGE_SIZE_A4,
  setParagraphText,
} from "./index.js";

import {
  addComment,
  appendParagraph,
  appendSectionBreak,
  clearBody,
  commentsPart,
  createDocx,
  insertParagraphAt,
  openDocx,
  paragraphs,
  removeAllComments,
  removeParagraph,
  text,
  toUint8Array,
} from "./docx.js";
describe("Docx.appendSectionBreak", () => {
  it("inserts a section break paragraph with sectPr in pPr", () => {
    const doc = createDocx({ paragraphs: ["section one"] });
    appendSectionBreak(doc, "nextPage");
    appendParagraph(doc, "section two");
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    const part = getPart(reopened.opc, "/word/document.xml");
    const xml = new TextDecoder().decode(part?.data ?? new Uint8Array());
    expect(xml).toContain('<w:type w:val="nextPage"/>');
  });

  it("can override page size and margins for the next section", () => {
    const doc = createDocx({ paragraphs: ["portrait section"] });
    appendSectionBreak(doc, "nextPage", {
      pageSize: { ...PAGE_SIZE_A4, orientation: "landscape" },
    });
    appendParagraph(doc, "landscape section");
    const xml = new TextDecoder().decode(
      getPart(openDocx(toUint8Array(doc)).opc, "/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain('w:orient="landscape"');
  });
});

describe("Docx.insertParagraphAt", () => {
  it("inserts a paragraph at the given paragraph-relative index", () => {
    const doc = createDocx({ paragraphs: ["A", "B", "C"] });
    insertParagraphAt(doc, 1, "X");
    expect(text(doc)).toBe("A\nX\nB\nC");
  });

  it("appends when the index is past the end", () => {
    const doc = createDocx({ paragraphs: ["A"] });
    insertParagraphAt(doc, 99, "Y");
    expect(text(doc)).toBe("A\nY");
  });

  it("preserves styling options", () => {
    const doc = createDocx({ paragraphs: ["A"] });
    insertParagraphAt(doc, 0, "Heading", { style: "Heading1" });
    const part = getPart(doc.opc, "/word/document.xml");
    // forces a flush before reading via the AST
    toUint8Array(doc);
    expect(paragraphs(doc)[0]?.pPr?.children.length).toBeGreaterThan(0);
    expect(part).toBeDefined();
  });
});

describe("Docx.removeParagraph / clearBody", () => {
  it("removes the paragraph at the given index (paragraph-relative)", () => {
    const doc = createDocx({ paragraphs: ["one", "two", "three"] });
    expect(removeParagraph(doc, 1)).toBe(true);
    expect(paragraphs(doc).map((p) => p.children.length).length).toBe(2);
    expect(text(doc)).toBe("one\nthree");
  });

  it("returns false when the index is out of range", () => {
    const doc = createDocx({ paragraphs: ["one"] });
    expect(removeParagraph(doc, 5)).toBe(false);
  });

  it("clearBody empties the body", () => {
    const doc = createDocx({ paragraphs: ["one", "two"] });
    clearBody(doc);
    expect(paragraphs(doc)).toHaveLength(0);
    expect(text(doc)).toBe("");
  });
});

describe("Docx.removeAllComments", () => {
  it("clears comments.xml entries and strips the markers from the body", () => {
    const doc = createDocx({ paragraphs: ["paragraph 1"] });
    const p = paragraphs(doc)[0];
    if (!p) return;
    addComment(doc, p, { author: "R", text: "First" });
    addComment(doc, p, { author: "R", text: "Second" });
    expect(commentsPart(doc)?.comments).toHaveLength(2);
    expect(removeAllComments(doc)).toBe(2);
    expect(commentsPart(doc)?.comments).toHaveLength(0);
    // No more rangeStart/rangeEnd or commentReference in the paragraph.
    const bytes = toUint8Array(doc);
    const xml = new TextDecoder().decode(
      getPart(openDocx(bytes).opc, "/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).not.toContain("commentRangeStart");
    expect(xml).not.toContain("commentReference");
  });
});

describe("paragraph property accessors", () => {
  it("getParagraphStyle / getParagraphAlignment read pPr children", () => {
    const doc = createDocx({ paragraphs: [] });
    const para = appendParagraph(doc, "hi", { style: "Heading1" });
    expect(getParagraphStyle(para)).toBe("Heading1");
    expect(getParagraphAlignment(para)).toBeUndefined();
  });

  it("setParagraphText replaces existing runs with a single styled run", () => {
    const doc = createDocx({ paragraphs: ["original"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    setParagraphText(para, "rewritten", { bold: true });
    expect(text(doc)).toBe("rewritten");
    expect(para.children).toHaveLength(1);
  });
});
