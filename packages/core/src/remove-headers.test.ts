import { hasPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  addFooter,
  addHeader,
  addPageNumberFooter,
  createDocx,
  footers,
  headers,
  openDocx,
  removeAllFooters,
  removeAllHeaders,
  toUint8Array,
} from "./docx.js";

describe("Docx.removeAllHeaders", () => {
  it("removes every header part and reference", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addHeader(doc, "h1");
    addHeader(doc, "h2", "first");
    expect(headers(doc)).toHaveLength(2);
    expect(removeAllHeaders(doc)).toBe(2);
    expect(headers(doc)).toHaveLength(0);
    expect(hasPart(doc.opc, "/word/header1.xml")).toBe(false);
    expect(hasPart(doc.opc, "/word/header2.xml")).toBe(false);
  });

  it("survives save+reopen", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addHeader(doc, "h");
    removeAllHeaders(doc);
    const reopened = openDocx(toUint8Array(doc));
    expect(headers(reopened)).toHaveLength(0);
  });

  it("returns 0 when no headers exist", () => {
    const doc = createDocx();
    expect(removeAllHeaders(doc)).toBe(0);
  });
});

describe("Docx.removeAllFooters", () => {
  it("removes every footer part and reference", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addFooter(doc, "f1");
    addFooter(doc, "f2", "even");
    expect(footers(doc)).toHaveLength(2);
    expect(removeAllFooters(doc)).toBe(2);
    expect(footers(doc)).toHaveLength(0);
  });

  it("works with addPageNumberFooter footers too", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addPageNumberFooter(doc, "Page ");
    expect(footers(doc)).toHaveLength(1);
    expect(removeAllFooters(doc)).toBe(1);
    expect(footers(doc)).toHaveLength(0);
  });
});
