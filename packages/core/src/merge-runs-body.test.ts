import { describe, expect, it } from "vitest";
import {
  addTable,
  appendParagraph,
  appendTextRun,
  createDocx,
  mergeAdjacentRunsInBody,
  paragraphs,
  paragraphText,
  tables,
} from "./index.js";

describe("mergeAdjacentRunsInBody", () => {
  it("collapses runs across every body paragraph and reports the total", () => {
    const doc = createDocx({ paragraphs: ["one"] });
    const p1 = paragraphs(doc).at(-1)!;
    appendTextRun(p1, " two");
    appendTextRun(p1, " three");

    appendParagraph(doc, "alpha");
    const p2 = paragraphs(doc).at(-1)!;
    appendTextRun(p2, " beta");

    const merges = mergeAdjacentRunsInBody(doc);
    // Each paragraph: seed run + the new appended runs all share no rPr,
    // so they collapse — 2 in p1 (3 runs → 1 = 2 merges) + 1 in p2 = 3.
    expect(merges).toBe(3);
    expect(paragraphText(p1)).toBe("one two three");
    expect(paragraphText(p2)).toBe("alpha beta");
  });

  it("descends into table cells", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["cell"]]);
    const t = tables(doc)[0]!;
    const cellPara = t.rows[0]?.cells[0]?.paragraphs[0]!;
    appendTextRun(cellPara, " more");
    appendTextRun(cellPara, " text");

    const merges = mergeAdjacentRunsInBody(doc);
    // cell starts with one run "cell"; we appended two more — no rPr on
    // any → 2 merges.
    expect(merges).toBe(2);
    expect(paragraphText(cellPara)).toBe("cell more text");
  });

  it("returns 0 when nothing in the body needs merging", () => {
    const doc = createDocx({ paragraphs: ["solo"] });
    expect(mergeAdjacentRunsInBody(doc)).toBe(0);
  });
});
