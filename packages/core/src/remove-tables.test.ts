import { describe, expect, it } from "vitest";
import {
  addTable,
  appendParagraph,
  createDocx,
  openDocx,
  paragraphs,
  removeAllTables,
  removeTable,
  tables,
  toUint8Array,
} from "./docx.js";

describe("removeTable", () => {
  it("removes the index-th table and keeps surrounding paragraphs", () => {
    const doc = createDocx({ paragraphs: ["before"] });
    addTable(doc, [["a", "b"]]);
    appendParagraph(doc, "between");
    addTable(doc, [["c", "d"]]);
    appendParagraph(doc, "after");

    expect(tables(doc)).toHaveLength(2);
    expect(removeTable(doc, 0)).toBe(true);
    expect(tables(doc)).toHaveLength(1);
    // Surviving table should be the second one (cells c/d).
    const remaining = tables(doc)[0]!;
    const firstCellRun = remaining.rows[0]?.cells[0]?.paragraphs[0]?.children[0];
    if (firstCellRun?.kind === "run") {
      const text = firstCellRun.pieces
        .filter((p) => p.kind === "text")
        .map((p) => p.value)
        .join("");
      expect(text).toBe("c");
    } else {
      throw new Error("expected a run with text content");
    }
    // Paragraphs untouched.
    expect(paragraphs(doc).map((p) => p.children).length).toBeGreaterThanOrEqual(3);
  });

  it("returns false when index is out of range", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["x"]]);
    expect(removeTable(doc, 7)).toBe(false);
    expect(tables(doc)).toHaveLength(1);
  });

  it("survives save+reopen with the survivor intact", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addTable(doc, [["1"]]);
    addTable(doc, [["2"]]);
    removeTable(doc, 0);
    const reopened = openDocx(toUint8Array(doc));
    expect(tables(reopened)).toHaveLength(1);
  });
});

describe("removeAllTables", () => {
  it("drops every table and returns the count", () => {
    const doc = createDocx({ paragraphs: ["intro"] });
    addTable(doc, [["a"]]);
    appendParagraph(doc, "mid");
    addTable(doc, [["b"]]);
    addTable(doc, [["c"]]);
    expect(removeAllTables(doc)).toBe(3);
    expect(tables(doc)).toHaveLength(0);
    // Paragraphs kept.
    expect(paragraphs(doc).length).toBeGreaterThanOrEqual(2);
  });

  it("returns 0 and doesn't mark dirty when there are no tables", () => {
    const doc = createDocx({ paragraphs: ["only text"] });
    expect(removeAllTables(doc)).toBe(0);
  });
});
