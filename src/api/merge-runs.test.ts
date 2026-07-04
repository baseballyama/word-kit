import { describe, expect, it } from "vitest";
import {
  appendParagraph,
  appendTextRun,
  createDocx,
  mergeAdjacentRuns,
  paragraphs,
  paragraphText,
} from "./index.js";

function emptyParagraph(doc: ReturnType<typeof createDocx>) {
  // appendParagraph seeds an unstyled run with an empty text piece. We start
  // from there so tests that count merges only have to reason about runs we
  // explicitly add on top of that seed run.
  appendParagraph(doc, "");
  const last = paragraphs(doc).at(-1)!;
  return last;
}

describe("mergeAdjacentRuns", () => {
  it("collapses adjacent unstyled runs", () => {
    const doc = createDocx({ paragraphs: [] });
    const target = emptyParagraph(doc);
    appendTextRun(target, "Hello, ");
    appendTextRun(target, "world!");
    // Three unstyled runs total (seed + Hello + world) → 2 merges.
    expect(mergeAdjacentRuns(target)).toBe(2);
    expect(target.children).toHaveLength(1);
    expect(paragraphText(target)).toBe("Hello, world!");
  });

  it("merges runs with identical formatting but not when formatting differs", () => {
    const doc = createDocx({ paragraphs: [] });
    const target = emptyParagraph(doc);
    appendTextRun(target, "A", { bold: true });
    appendTextRun(target, "B", { bold: true });
    appendTextRun(target, "C", { italic: true });
    // Seed unstyled run sits alone, A+B merge, C is its own run.
    expect(mergeAdjacentRuns(target)).toBe(1);
    expect(target.children).toHaveLength(3);
    expect(paragraphText(target)).toBe("ABC");
  });

  it("returns 0 when there is nothing to merge", () => {
    const doc = createDocx({ paragraphs: [] });
    const target = emptyParagraph(doc);
    appendTextRun(target, "alone", { bold: true });
    // Seed run (no rPr) and the bold run differ in formatting → no merge.
    expect(mergeAdjacentRuns(target)).toBe(0);
  });

  it("does not merge across a non-run inline (eg <w:hyperlink>)", () => {
    const doc = createDocx({ paragraphs: [] });
    const target = emptyParagraph(doc);
    appendTextRun(target, "before", { bold: true });
    target.children.push({
      kind: "raw",
      node: {
        kind: "element",
        name: {
          uri: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
          local: "hyperlink",
          prefix: "w",
        },
        attrs: [],
        children: [],
        xmlSpace: "default",
        selfClosing: true,
      },
    });
    appendTextRun(target, "after", { bold: true });
    // The hyperlink raw inline interrupts adjacency between the two bold runs.
    // Seed run and "before" differ in formatting too → 0 merges total.
    expect(mergeAdjacentRuns(target)).toBe(0);
    expect(target.children).toHaveLength(4);
  });

  it("merges a chain of three identically-formatted runs into one", () => {
    const doc = createDocx({ paragraphs: [] });
    const target = emptyParagraph(doc);
    appendTextRun(target, "X", { bold: true });
    appendTextRun(target, "Y", { bold: true });
    appendTextRun(target, "Z", { bold: true });
    // Seed unstyled run alone, then bold X+Y+Z merge into one (2 merges).
    expect(mergeAdjacentRuns(target)).toBe(2);
    expect(target.children).toHaveLength(2);
    expect(paragraphText(target)).toBe("XYZ");
  });
});
