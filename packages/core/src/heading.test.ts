import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.appendHeading + outline", () => {
  it("appendHeading sets pStyle to HeadingN", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.appendHeading("Chapter 1", 1);
    doc.appendHeading("Section 1.1", 2);
    doc.appendParagraph("body");

    const outline = doc.outline();
    expect(outline).toHaveLength(2);
    expect(outline[0]).toMatchObject({ level: 1, text: "Chapter 1" });
    expect(outline[1]).toMatchObject({ level: 2, text: "Section 1.1" });
  });

  it("rejects invalid heading levels", () => {
    const doc = Docx.create({ paragraphs: [] });
    expect(() => doc.appendHeading("x", 0)).toThrow();
    expect(() => doc.appendHeading("x", 10)).toThrow();
    expect(() => doc.appendHeading("x", 1.5)).toThrow();
  });

  it("ensureHeadingStyles seeds Heading1..N", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.ensureHeadingStyles(3);
    const sp = doc.stylesPart;
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
    const doc = Docx.create({ paragraphs: [] });
    doc.ensureHeadingStyles(2);
    doc.ensureHeadingStyles(2);
    const count = doc.stylesPart?.styles.filter((s) =>
      s.attrs.some((a) => a.name.local === "styleId" && a.value === "Heading1"),
    ).length;
    expect(count).toBe(1);
  });

  it("outline survives save+reopen", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.ensureHeadingStyles(3);
    doc.appendHeading("Title", 1);
    doc.appendHeading("Section", 2);
    const reopened = Docx.open(doc.toUint8Array());
    const outline = reopened.outline();
    expect(outline.map((h) => h.text)).toEqual(["Title", "Section"]);
    expect(outline.map((h) => h.level)).toEqual([1, 2]);
  });
});
