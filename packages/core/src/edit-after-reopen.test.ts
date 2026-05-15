// "Edit a reopened document" audit. The library has accumulated a set of
// add* / append* helpers; each must work as well after a save+open
// round-trip as it does on a freshly-built doc, because the realistic
// workflow is `openDocx(template) → mutate → toUint8Array(out)`. Bugs
// that only show up post-round-trip would be invisible to the existing
// build-from-scratch tests.
//
// Each test follows the same shape:
//   1. build base
//   2. round-trip (toUint8Array → openDocx)
//   3. apply a mutation on the reopened doc
//   4. round-trip again, assert validate() returns no errors AND the
//      mutation actually landed (length / part presence / xml content).

import { hasPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  acceptAllRevisions,
  addBookmark,
  addBulletList,
  addComment,
  addEndnote,
  addFooter,
  addFootnote,
  addHeader,
  addHyperlink,
  addImage,
  addInternalHyperlink,
  addNumberedList,
  addStyle,
  addTable,
  addTableOfContents,
  appendField,
  appendHeading,
  appendMergeField,
  appendParagraph,
  bookmarks,
  commentsPart,
  createDocx,
  endnotesPart,
  externalHyperlinks,
  footers,
  footnotesPart,
  headers,
  images,
  listStyles,
  openDocx,
  paragraphs,
  removeStyle,
  setHyperlinkUrl,
  tables,
  toUint8Array,
  validate,
} from "./index.js";

const TINY_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xfa, 0xcf, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);

function expectClean(bytes: Uint8Array): ReturnType<typeof openDocx> {
  const reopened = openDocx(bytes);
  const issues = validate(reopened);
  const errors = issues.filter((i) => i.level === "error");
  if (errors.length > 0) {
    throw new Error(
      `validate() returned errors:\n${errors.map((e) => `  [${e.code}] ${e.message}`).join("\n")}`,
    );
  }
  return reopened;
}

function reopen(doc: ReturnType<typeof createDocx>): ReturnType<typeof createDocx> {
  return openDocx(toUint8Array(doc));
}

describe("editing a reopened document", () => {
  it("addComment on a reopened doc with no prior comments", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    const reopened = reopen(seed);
    addComment(reopened, paragraphs(reopened)[0]!, {
      author: "R",
      initials: "R",
      text: "added after reopen",
    });
    const final = expectClean(toUint8Array(reopened));
    expect(commentsPart(final)?.comments.length).toBe(1);
  });

  it("addComment on a reopened doc that already has a comment (id collision avoidance)", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    addComment(seed, paragraphs(seed)[0]!, { author: "R", initials: "R", text: "first" });
    const reopened = reopen(seed);
    appendParagraph(reopened, "second body");
    addComment(reopened, paragraphs(reopened).at(-1)!, {
      author: "R",
      initials: "R",
      text: "second",
    });
    const final = expectClean(toUint8Array(reopened));
    expect(commentsPart(final)?.comments.length).toBe(2);
  });

  it("addFootnote on a reopened doc with no prior footnotes", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    const reopened = reopen(seed);
    addFootnote(reopened, paragraphs(reopened)[0]!, "added after reopen");
    const final = expectClean(toUint8Array(reopened));
    // separator + continuationSeparator + 1 user note
    expect(footnotesPart(final)?.footnotes.length).toBeGreaterThanOrEqual(3);
  });

  it("addEndnote on a reopened doc", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    const reopened = reopen(seed);
    addEndnote(reopened, paragraphs(reopened)[0]!, "after reopen");
    const final = expectClean(toUint8Array(reopened));
    expect(endnotesPart(final)?.footnotes.length).toBeGreaterThanOrEqual(3);
  });

  it("addImage on a reopened doc", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    const reopened = reopen(seed);
    addImage(reopened, TINY_PNG, { widthEmu: 914_400, heightEmu: 914_400 });
    const final = expectClean(toUint8Array(reopened));
    expect(images(final)).toHaveLength(1);
  });

  it("addHeader / addFooter on a reopened doc", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    const reopened = reopen(seed);
    addHeader(reopened, "added header");
    addFooter(reopened, "added footer");
    const bytes = toUint8Array(reopened);
    const final = expectClean(bytes);
    expect(headers(final)).toHaveLength(1);
    expect(footers(final)).toHaveLength(1);
  });

  it("addHyperlink on a reopened doc allocates a fresh rel id", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    addHyperlink(seed, "https://stage.example.com/", "stage");
    const reopened = reopen(seed);
    addHyperlink(reopened, "https://prod.example.com/", "prod");
    const final = expectClean(toUint8Array(reopened));
    const targets = externalHyperlinks(final)
      .map((h) => h.target)
      .toSorted();
    expect(targets).toEqual(["https://prod.example.com/", "https://stage.example.com/"]);
  });

  it("addInternalHyperlink + addBookmark on a reopened doc", () => {
    const seed = createDocx({ paragraphs: ["intro", "anchor"] });
    addBookmark(seed, "ch1", paragraphs(seed)[1]!);
    const reopened = reopen(seed);
    addInternalHyperlink(reopened, "ch1", "jump");
    const final = expectClean(toUint8Array(reopened));
    expect(bookmarks(final).map((b) => b.name)).toContain("ch1");
  });

  it("addTable + addBulletList + addNumberedList on a reopened doc", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    const reopened = reopen(seed);
    addTable(reopened, [
      ["A", "B"],
      ["1", "2"],
    ]);
    addBulletList(reopened, ["x", "y"]);
    addNumberedList(reopened, ["one", "two"]);
    const final = expectClean(toUint8Array(reopened));
    expect(tables(final)).toHaveLength(1);
    expect(hasPart(final.opc, "/word/numbering.xml")).toBe(true);
  });

  it("addStyle + removeStyle on a reopened doc", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    addStyle(seed, { type: "paragraph", styleId: "X1", name: "X1" });
    const reopened = reopen(seed);
    addStyle(reopened, { type: "paragraph", styleId: "X2", name: "X2" });
    expect(removeStyle(reopened, "X1")).toBe(true);
    const final = expectClean(toUint8Array(reopened));
    const ids = listStyles(final).map((s) => s.styleId);
    expect(ids).toContain("X2");
    expect(ids).not.toContain("X1");
  });

  it("addTableOfContents + appendMergeField + appendField on a reopened doc", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    const reopened = reopen(seed);
    addTableOfContents(reopened);
    appendParagraph(reopened, "Dear ");
    appendMergeField(reopened, paragraphs(reopened).at(-1)!, "FirstName");
    appendParagraph(reopened, "Date: ");
    appendField(reopened, paragraphs(reopened).at(-1)!, "DATE", "[date]");
    const bytes = toUint8Array(reopened);
    const final = expectClean(bytes);
    const xml = new TextDecoder().decode(
      final.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain('TOC \\o "1-3"');
    expect(xml).toContain("MERGEFIELD FirstName");
    expect(xml).toContain("DATE");
  });

  it("acceptAllRevisions on a reopened doc with no revisions is a no-op", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    const reopened = reopen(seed);
    expect(acceptAllRevisions(reopened)).toBe(0);
    expectClean(toUint8Array(reopened));
  });

  it("appendHeading after reopen still uses the seeded Heading style", () => {
    const seed = createDocx({ paragraphs: [] });
    const reopened = reopen(seed);
    appendHeading(reopened, "After reopen", 1);
    const final = expectClean(toUint8Array(reopened));
    const xml = new TextDecoder().decode(
      final.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain('<w:pStyle w:val="Heading1"');
  });

  it("setHyperlinkUrl after reopen rewrites the rel target on disk", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    addHyperlink(seed, "https://stage.example.com/x", "x");
    const reopened = reopen(seed);
    expect(setHyperlinkUrl(reopened, (t) => t.replace("stage.", ""))).toBe(1);
    const final = expectClean(toUint8Array(reopened));
    expect(externalHyperlinks(final)[0]?.target).toBe("https://example.com/x");
  });
});
