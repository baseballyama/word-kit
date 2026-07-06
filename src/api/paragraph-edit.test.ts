/**
 * Paragraph split/merge and run-range isolation — the caret-level structural
 * primitives behind the editor's Enter, Backspace, and selection formatting.
 */

import { describe, expect, it } from "vitest";
import {
  createDocx,
  getRunFormat,
  isolateParagraphRunRange,
  mergeParagraphIntoPrevious,
  openDocx,
  paragraphs,
  paragraphText,
  runTextLength,
  setRunFormat,
  splitParagraphAt,
  toUint8Array,
  validate,
  type WmlRun,
} from "../index.js";

function runTexts(doc: ReturnType<typeof createDocx>, block: number): string[] {
  const para = doc.document.body.blocks[block];
  if (!para || para.kind !== "paragraph") return [];
  return para.children
    .filter((c): c is WmlRun => c.kind === "run")
    .map((r) => r.pieces.map((p) => (p.kind === "text" ? p.value : "")).join(""));
}

describe("splitParagraphAt", () => {
  it("splits at an offset, inherits pPr, and round-trips", () => {
    const doc = createDocx({ paragraphs: ["Hello world"] });
    const nb = splitParagraphAt(doc, 0, 0, 5);
    expect(nb).toBe(1);
    expect(paragraphs(doc).map(paragraphText)).toEqual(["Hello", " world"]);
    const re = openDocx(toUint8Array(doc));
    expect(validate(re).length).toBe(0);
  });

  it("returns -1 for a non-paragraph block", () => {
    const doc = createDocx({ paragraphs: ["x"] });
    expect(splitParagraphAt(doc, 99, 0, 0)).toBe(-1);
  });
});

describe("mergeParagraphIntoPrevious", () => {
  it("merges into the previous paragraph and reports the join", () => {
    const doc = createDocx({ paragraphs: ["Hello", " world"] });
    const pos = mergeParagraphIntoPrevious(doc, 1);
    expect(pos).toEqual({ block: 0, inline: 0, offset: 5 });
    expect(paragraphs(doc).map(paragraphText)).toEqual(["Hello world"]);
  });

  it("returns null with no preceding paragraph", () => {
    const doc = createDocx({ paragraphs: ["only"] });
    expect(mergeParagraphIntoPrevious(doc, 0)).toBeNull();
  });
});

describe("runTextLength / isolateParagraphRunRange", () => {
  it("counts visible characters", () => {
    const doc = createDocx({ paragraphs: ["abcdef"] });
    const run = paragraphs(doc)[0]!.children.find((c): c is WmlRun => c.kind === "run")!;
    expect(runTextLength(run)).toBe(6);
  });

  it("isolates a middle range into its own run", () => {
    const doc = createDocx({ paragraphs: ["Hello world foo"] });
    const para = paragraphs(doc)[0]!;
    const runs = isolateParagraphRunRange(para, 6, 11);
    for (const r of runs) setRunFormat(r, { bold: true });
    expect(runTexts(doc, 0)).toEqual(["Hello ", "world", " foo"]);
    const re = openDocx(toUint8Array(doc));
    expect(validate(re).length).toBe(0);
    const bolded = paragraphs(re)[0]!
      .children.filter((c): c is WmlRun => c.kind === "run")
      .find((r) => getRunFormat(r).bold);
    expect(bolded?.pieces.map((p) => (p.kind === "text" ? p.value : "")).join("")).toBe("world");
  });

  it("returns nothing for an empty range", () => {
    const doc = createDocx({ paragraphs: ["abc"] });
    expect(isolateParagraphRunRange(paragraphs(doc)[0]!, 2, 2)).toEqual([]);
  });
});
