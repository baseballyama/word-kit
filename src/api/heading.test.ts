import { describe, expect, it } from "vitest";
import {
  appendHeading,
  appendParagraph,
  createDocx,
  ensureHeadingStyles,
  openDocx,
  outline,
  stylesPart,
  toUint8Array,
} from "./docx.js";

describe("Docx.appendHeading + outline", () => {
  it("appendHeading sets pStyle to HeadingN", () => {
    const doc = createDocx({ paragraphs: [] });
    appendHeading(doc, "Chapter 1", 1);
    appendHeading(doc, "Section 1.1", 2);
    appendParagraph(doc, "body");

    const otl = outline(doc);
    expect(otl).toHaveLength(2);
    expect(otl[0]).toMatchObject({ level: 1, text: "Chapter 1" });
    expect(otl[1]).toMatchObject({ level: 2, text: "Section 1.1" });
  });

  it("rejects invalid heading levels", () => {
    const doc = createDocx({ paragraphs: [] });
    expect(() => appendHeading(doc, "x", 0)).toThrow();
    expect(() => appendHeading(doc, "x", 10)).toThrow();
    expect(() => appendHeading(doc, "x", 1.5)).toThrow();
  });

  it("ensureHeadingStyles seeds Heading1..N", () => {
    const doc = createDocx({ paragraphs: [] });
    ensureHeadingStyles(doc, 3);
    const sp = stylesPart(doc);
    expect(sp).toBeDefined();
    if (!sp) return;
    const ids = sp.styles
      .map((s) => s.attrs.find((a) => a.name.local === "styleId")?.value)
      .filter((x): x is string => !!x);
    expect(ids).toContain("Heading1");
    expect(ids).toContain("Heading2");
    expect(ids).toContain("Heading3");
    expect(ids).not.toContain("Heading4");
  });

  it("ensureHeadingStyles is idempotent", () => {
    const doc = createDocx({ paragraphs: [] });
    ensureHeadingStyles(doc, 2);
    ensureHeadingStyles(doc, 2);
    const count = stylesPart(doc)?.styles.filter((s) =>
      s.attrs.some((a) => a.name.local === "styleId" && a.value === "Heading1"),
    ).length;
    expect(count).toBe(1);
  });

  it("outline survives save+reopen", () => {
    const doc = createDocx({ paragraphs: [] });
    ensureHeadingStyles(doc, 3);
    appendHeading(doc, "Title", 1);
    appendHeading(doc, "Section", 2);
    const reopened = openDocx(toUint8Array(doc));
    const otl = outline(reopened);
    expect(otl.map((h) => h.text)).toEqual(["Title", "Section"]);
    expect(otl.map((h) => h.level)).toEqual([1, 2]);
  });
});
