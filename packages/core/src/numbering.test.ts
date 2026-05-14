import { numAbstractRef, numId } from "@word-kit/wml";
import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.addBulletList", () => {
  it("creates numbering.xml + abstractNum + num on first use", () => {
    const doc = Docx.create({ paragraphs: [] });
    expect(doc.numberingPart).toBeUndefined();
    doc.addBulletList(["apple", "banana", "cherry"]);
    const part = doc.numberingPart;
    expect(part).toBeDefined();
    if (!part) return;
    expect(part.abstractNums.length).toBeGreaterThan(0);
    expect(part.nums.length).toBeGreaterThan(0);
  });

  it("appends one paragraph per list item with numPr in pPr", () => {
    const doc = Docx.create({ paragraphs: [] });
    const paragraphs = doc.addBulletList(["a", "b"]);
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0]?.pPr).toBeDefined();
    // Look for a <w:numPr> in the pPr.
    const pPr = paragraphs[0]?.pPr;
    const hasNumPr = pPr?.children.some((c) => c.kind === "element" && c.name.local === "numPr");
    expect(hasNumPr).toBe(true);
  });

  it("survives save+reopen and Words sees the items", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addBulletList(["first", "second", "third"]);
    const bytes = doc.toUint8Array();
    const reopened = Docx.open(bytes);
    expect(reopened.numberingPart).toBeDefined();
    expect(reopened.paragraphs).toHaveLength(3);
    expect(reopened.text).toBe("first\nsecond\nthird");
  });

  it("re-uses the same num definition across calls", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addBulletList(["a"]);
    doc.addBulletList(["b"]);
    const part = doc.numberingPart;
    expect(part?.nums.length).toBe(1);
  });
});

describe("Docx.addNumberedList", () => {
  it("uses a decimal abstractNum distinct from the bullet one", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addBulletList(["b"]);
    doc.addNumberedList(["n"]);
    const part = doc.numberingPart;
    expect(part).toBeDefined();
    if (!part) return;
    expect(part.abstractNums.length).toBe(2);
    expect(part.nums.length).toBe(2);
    const refs = part.nums.map(numAbstractRef);
    expect(new Set(refs).size).toBe(2);
    const ids = part.nums.map(numId);
    expect(new Set(ids).size).toBe(2);
  });

  it("save+reopen preserves both lists", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addBulletList(["bullet"]);
    doc.addNumberedList(["one", "two"]);
    const reopened = Docx.open(doc.toUint8Array());
    expect(reopened.paragraphs).toHaveLength(3);
    expect(reopened.text).toBe("bullet\none\ntwo");
  });
});
