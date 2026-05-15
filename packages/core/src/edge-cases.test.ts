// Boundary-input audit. Most APIs accept "reasonable" input in
// happy-path tests; this file feeds them empty / single-element /
// large / unusual values and asserts they don't blow up and the
// produced docx still validates.

import { describe, expect, it } from "vitest";
import {
  addBulletList,
  addComment,
  addNumberedList,
  addStyle,
  addTable,
  appendParagraph,
  createDocx,
  openDocx,
  paragraphs,
  removeStyle,
  setHyperlinkUrl,
  tables,
  toUint8Array,
  validate,
} from "./index.js";

function expectClean(doc: ReturnType<typeof createDocx>): ReturnType<typeof createDocx> {
  const reopened = openDocx(toUint8Array(doc));
  const errors = validate(reopened).filter((i) => i.level === "error");
  if (errors.length > 0) {
    throw new Error(
      `validate() returned errors:\n${errors.map((e) => `  [${e.code}] ${e.message}`).join("\n")}`,
    );
  }
  return reopened;
}

describe("createDocx edge cases", () => {
  it("paragraphs: undefined → seeds one empty paragraph", () => {
    const doc = createDocx();
    expectClean(doc);
    expect(paragraphs(doc).length).toBeGreaterThanOrEqual(1);
  });

  it("paragraphs: [] → seeds zero paragraphs (caller is providing the body)", () => {
    const doc = createDocx({ paragraphs: [] });
    expectClean(doc);
    expect(paragraphs(doc).length).toBe(0);
  });

  it('paragraphs: [""] → one empty paragraph', () => {
    const doc = createDocx({ paragraphs: [""] });
    expectClean(doc);
    expect(paragraphs(doc).length).toBe(1);
  });
});

describe("appendParagraph edge cases", () => {
  it("empty string", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "");
    expectClean(doc);
  });

  it("very long string (~64 KB)", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "x".repeat(64 * 1024));
    expectClean(doc);
  });

  it("string with only whitespace", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "   \t\n   ");
    expectClean(doc);
  });

  it("string containing every common whitespace incl. NBSP and emoji", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "spaces   tabs \t newlines \n end 🎉");
    expectClean(doc);
  });

  it("CJK with combining marks (Korean Hangul, Japanese Kana, Chinese)", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "한국어 日本語 中文 ja: が゙ ぱ");
    expectClean(doc);
  });

  it("RTL text (Arabic + Hebrew)", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "مرحبا بالعالم שלום עולם");
    expectClean(doc);
  });

  it("appending many paragraphs in a tight loop", () => {
    const doc = createDocx({ paragraphs: [] });
    for (let i = 0; i < 1000; i++) appendParagraph(doc, `line ${i}`);
    const reopened = expectClean(doc);
    expect(paragraphs(reopened).length).toBe(1000);
  });
});

describe("addBulletList / addNumberedList edge cases", () => {
  it("empty array — no bullets, no errors", () => {
    const doc = createDocx({ paragraphs: [] });
    addBulletList(doc, []);
    expectClean(doc);
  });

  it("empty-string items render as empty bullets", () => {
    const doc = createDocx({ paragraphs: [] });
    addBulletList(doc, ["", "", ""]);
    expectClean(doc);
  });

  it("numbered list with one very long item", () => {
    const doc = createDocx({ paragraphs: [] });
    addNumberedList(doc, ["X".repeat(8 * 1024)]);
    expectClean(doc);
  });
});

describe("addTable edge cases", () => {
  it("empty rows array — no table created (or empty table; assert clean)", () => {
    const doc = createDocx({ paragraphs: ["before"] });
    addTable(doc, []);
    expectClean(doc);
  });

  it("ragged rows are padded to the widest column count", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["a", "b", "c"], ["1"], ["x", "y"]]);
    const reopened = expectClean(doc);
    const t = tables(reopened)[0]!;
    // All rows should have 3 cells (widest row).
    for (const row of t.rows) {
      expect(row.cells.length).toBe(3);
    }
  });

  it("single-cell single-row table", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["only"]]);
    expectClean(doc);
  });

  it("100-row table", () => {
    const doc = createDocx({ paragraphs: [] });
    const rows: string[][] = [];
    for (let i = 0; i < 100; i++) rows.push([`r${i}c0`, `r${i}c1`]);
    addTable(doc, rows);
    const reopened = expectClean(doc);
    expect(tables(reopened)[0]?.rows.length).toBe(100);
  });
});

describe("addComment edge cases", () => {
  it("empty author / initials / text", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addComment(doc, paragraphs(doc)[0]!, { author: "", initials: "", text: "" });
    expectClean(doc);
  });

  it("multi-line comment text", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addComment(doc, paragraphs(doc)[0]!, {
      author: "R",
      initials: "R",
      text: "line one\nline two\nline three",
    });
    expectClean(doc);
  });
});

describe("addStyle / removeStyle edge cases", () => {
  it("addStyle then add same id again (overwrites the existing definition)", () => {
    const doc = createDocx({ paragraphs: [] });
    addStyle(doc, { type: "paragraph", styleId: "S", name: "S", bold: true });
    addStyle(doc, { type: "paragraph", styleId: "S", name: "S", italic: true });
    expectClean(doc);
  });

  it("removeStyle on an unknown id returns false and is otherwise a no-op", () => {
    const doc = createDocx({ paragraphs: [] });
    expect(removeStyle(doc, "NoSuchStyle")).toBe(false);
    expectClean(doc);
  });
});

describe("setHyperlinkUrl edge cases", () => {
  it("predicate that returns the same URL is treated as no-op", () => {
    const doc = createDocx({ paragraphs: [] });
    expect(setHyperlinkUrl(doc, () => null)).toBe(0);
    expectClean(doc);
  });
});
