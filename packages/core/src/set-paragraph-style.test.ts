import { describe, expect, it } from "vitest";
import {
  addStyle,
  appendParagraph,
  createDocx,
  ensureHeadingStyles,
  getParagraphStyle,
  openDocx,
  paragraphs,
  setParagraphStyle,
  toUint8Array,
} from "./index.js";

describe("setParagraphStyle", () => {
  it("applies a styleId on a paragraph that had no pStyle", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    ensureHeadingStyles(doc, 1);
    const p = paragraphs(doc)[0]!;
    expect(getParagraphStyle(p)).toBeUndefined();
    setParagraphStyle(p, "Heading1");
    expect(getParagraphStyle(p)).toBe("Heading1");
  });

  it("replaces an existing pStyle rather than stacking duplicates", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    ensureHeadingStyles(doc, 2);
    const p = paragraphs(doc)[0]!;
    setParagraphStyle(p, "Heading1");
    setParagraphStyle(p, "Heading2");
    expect(getParagraphStyle(p)).toBe("Heading2");
    // Verify only one pStyle exists in pPr.
    const matches = (p.pPr?.children ?? []).filter(
      (c) => c.kind === "element" && c.name.local === "pStyle",
    );
    expect(matches).toHaveLength(1);
  });

  it("clears pStyle when undefined is passed", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addStyle(doc, { type: "paragraph", styleId: "Custom", name: "Custom" });
    const p = paragraphs(doc)[0]!;
    setParagraphStyle(p, "Custom");
    setParagraphStyle(p, undefined);
    expect(getParagraphStyle(p)).toBeUndefined();
  });

  it("does not allocate a pPr when clearing on a paragraph that has none", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const p = paragraphs(doc)[0]!;
    expect(p.pPr).toBeUndefined();
    setParagraphStyle(p, undefined);
    expect(p.pPr).toBeUndefined();
  });

  it("survives a save+reopen", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    ensureHeadingStyles(doc, 1);
    setParagraphStyle(paragraphs(doc)[0]!, "Heading1");
    const reopened = openDocx(toUint8Array(doc));
    expect(getParagraphStyle(paragraphs(reopened)[0]!)).toBe("Heading1");
  });

  it("works mid-document on an appended paragraph", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "callout");
    addStyle(doc, { type: "paragraph", styleId: "Callout", name: "Callout" });
    setParagraphStyle(paragraphs(doc).at(-1)!, "Callout");
    const reopened = openDocx(toUint8Array(doc));
    expect(getParagraphStyle(paragraphs(reopened).at(-1)!)).toBe("Callout");
  });
});
