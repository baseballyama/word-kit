import { numAbstractRef, numId } from "@word-kit/wml";
import { describe, expect, it } from "vitest";
import {
  addBulletList,
  addNumberedList,
  createDocx,
  numberingPart,
  openDocx,
  paragraphs,
  text,
  toUint8Array,
} from "./docx.js";

describe("Docx.addBulletList", () => {
  it("creates numbering.xml + abstractNum + num on first use", () => {
    const doc = createDocx({ paragraphs: [] });
    expect(numberingPart(doc)).toBeUndefined();
    addBulletList(doc, ["apple", "banana", "cherry"]);
    const part = numberingPart(doc);
    expect(part).toBeDefined();
    if (!part) return;
    expect(part.abstractNums.length).toBeGreaterThan(0);
    expect(part.nums.length).toBeGreaterThan(0);
  });

  it("appends one paragraph per list item with numPr in pPr", () => {
    const doc = createDocx({ paragraphs: [] });
    const paragraphs = addBulletList(doc, ["a", "b"]);
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0]?.pPr).toBeDefined();
    // Look for a <w:numPr> in the pPr.
    const pPr = paragraphs[0]?.pPr;
    const hasNumPr = pPr?.children.some((c) => c.kind === "element" && c.name.local === "numPr");
    expect(hasNumPr).toBe(true);
  });

  it("survives save+reopen and Words sees the items", () => {
    const doc = createDocx({ paragraphs: [] });
    addBulletList(doc, ["first", "second", "third"]);
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    expect(numberingPart(reopened)).toBeDefined();
    expect(paragraphs(reopened)).toHaveLength(3);
    expect(text(reopened)).toBe("first\nsecond\nthird");
  });

  it("re-uses the same num definition across calls", () => {
    const doc = createDocx({ paragraphs: [] });
    addBulletList(doc, ["a"]);
    addBulletList(doc, ["b"]);
    const part = numberingPart(doc);
    expect(part?.nums.length).toBe(1);
  });
});

describe("Docx.addNumberedList", () => {
  it("uses a decimal abstractNum distinct from the bullet one", () => {
    const doc = createDocx({ paragraphs: [] });
    addBulletList(doc, ["b"]);
    addNumberedList(doc, ["n"]);
    const part = numberingPart(doc);
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
    const doc = createDocx({ paragraphs: [] });
    addBulletList(doc, ["bullet"]);
    addNumberedList(doc, ["one", "two"]);
    const reopened = openDocx(toUint8Array(doc));
    expect(paragraphs(reopened)).toHaveLength(3);
    expect(text(reopened)).toBe("bullet\none\ntwo");
  });
});
