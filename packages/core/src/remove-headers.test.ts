import { hasPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.removeAllHeaders", () => {
  it("removes every header part and reference", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    doc.addHeader("h1");
    doc.addHeader("h2", "first");
    expect(doc.headers).toHaveLength(2);
    expect(doc.removeAllHeaders()).toBe(2);
    expect(doc.headers).toHaveLength(0);
    expect(hasPart(doc.opc, "/word/header1.xml")).toBe(false);
    expect(hasPart(doc.opc, "/word/header2.xml")).toBe(false);
  });

  it("survives save+reopen", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    doc.addHeader("h");
    doc.removeAllHeaders();
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.headers).toHaveLength(0);
  });

  it("returns 0 when no headers exist", () => {
    const doc = Docx.create();
    expect(doc.removeAllHeaders()).toBe(0);
  });
});

describe("Docx.removeAllFooters", () => {
  it("removes every footer part and reference", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    doc.addFooter("f1");
    doc.addFooter("f2", "even");
    expect(doc.footers).toHaveLength(2);
    expect(doc.removeAllFooters()).toBe(2);
    expect(doc.footers).toHaveLength(0);
  });

  it("works with addPageNumberFooter footers too", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    doc.addPageNumberFooter("Page ");
    expect(doc.footers).toHaveLength(1);
    expect(doc.removeAllFooters()).toBe(1);
    expect(doc.footers).toHaveLength(0);
  });
});
