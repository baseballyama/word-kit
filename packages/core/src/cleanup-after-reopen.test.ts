// `removeAllComments` was buggy when called *after a round-trip* — the
// original strip logic only matched the raw-inline shape produced by
// addComment, missing the WmlRun shape produced by parseRun. This file
// exercises every removeAllX helper through a save+reopen+remove cycle
// and asserts validate() returns no errors. If a future helper develops
// the same blind spot, this catches it before it ships.

import { describe, expect, it } from "vitest";
import {
  addBookmark,
  addComment,
  addEndnote,
  addFooter,
  addFootnote,
  addHeader,
  addHyperlink,
  addImage,
  appendParagraph,
  createDocx,
  openDocx,
  paragraphs,
  removeAllBookmarks,
  removeAllComments,
  removeAllEndnotes,
  removeAllFooters,
  removeAllFootnotes,
  removeAllHeaders,
  removeAllHyperlinks,
  removeAllImages,
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

function expectClean(bytes: Uint8Array): void {
  const reopened = openDocx(bytes);
  const issues = validate(reopened);
  const errors = issues.filter((i) => i.level === "error");
  if (errors.length > 0) {
    throw new Error(
      `validate() returned errors:\n${errors.map((e) => `  [${e.code}] ${e.message}`).join("\n")}`,
    );
  }
}

describe("removeAllX after a reopen → validate clean", () => {
  it("removeAllComments", () => {
    const doc = createDocx({ paragraphs: ["with comment"] });
    addComment(doc, paragraphs(doc)[0]!, {
      author: "R",
      initials: "R",
      text: "hi",
    });
    const reopened = openDocx(toUint8Array(doc));
    expect(removeAllComments(reopened)).toBe(1);
    expectClean(toUint8Array(reopened));
  });

  it("removeAllFootnotes", () => {
    const doc = createDocx({ paragraphs: ["with footnote"] });
    addFootnote(doc, paragraphs(doc)[0]!, "Source.");
    const reopened = openDocx(toUint8Array(doc));
    expect(removeAllFootnotes(reopened)).toBeGreaterThan(0);
    expectClean(toUint8Array(reopened));
  });

  it("removeAllEndnotes", () => {
    const doc = createDocx({ paragraphs: ["with endnote"] });
    addEndnote(doc, paragraphs(doc)[0]!, "Glossary.");
    const reopened = openDocx(toUint8Array(doc));
    expect(removeAllEndnotes(reopened)).toBeGreaterThan(0);
    expectClean(toUint8Array(reopened));
  });

  it("removeAllImages", () => {
    const doc = createDocx({ paragraphs: ["with image"] });
    addImage(doc, TINY_PNG, {
      widthEmu: 914_400,
      heightEmu: 914_400,
      altText: "x",
    });
    const reopened = openDocx(toUint8Array(doc));
    expect(removeAllImages(reopened)).toBe(1);
    expectClean(toUint8Array(reopened));
  });

  it("removeAllBookmarks", () => {
    const doc = createDocx({ paragraphs: ["with bookmark"] });
    addBookmark(doc, "section1", paragraphs(doc)[0]!);
    const reopened = openDocx(toUint8Array(doc));
    expect(removeAllBookmarks(reopened)).toBe(1);
    expectClean(toUint8Array(reopened));
  });

  it("removeAllHyperlinks", () => {
    const doc = createDocx({ paragraphs: ["intro"] });
    addHyperlink(doc, "https://example.com/", "click");
    const reopened = openDocx(toUint8Array(doc));
    expect(removeAllHyperlinks(reopened)).toBe(1);
    expectClean(toUint8Array(reopened));
  });

  it("removeAllHeaders", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addHeader(doc, "page header");
    const reopened = openDocx(toUint8Array(doc));
    expect(removeAllHeaders(reopened)).toBe(1);
    expectClean(toUint8Array(reopened));
  });

  it("removeAllFooters", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addFooter(doc, "page footer");
    const reopened = openDocx(toUint8Array(doc));
    expect(removeAllFooters(reopened)).toBe(1);
    expectClean(toUint8Array(reopened));
  });

  it("a chained cleanup pass leaves no validation errors", () => {
    const doc = createDocx({ paragraphs: ["mixed"] });
    addHeader(doc, "header");
    addFooter(doc, "footer");
    addComment(doc, paragraphs(doc)[0]!, { author: "R", initials: "R", text: "x" });
    addFootnote(doc, paragraphs(doc)[0]!, "fn");
    addEndnote(doc, paragraphs(doc)[0]!, "en");
    addBookmark(doc, "anchor", paragraphs(doc)[0]!);
    addHyperlink(doc, "https://example.com/", "link");
    addImage(doc, TINY_PNG, { widthEmu: 914_400, heightEmu: 914_400 });
    appendParagraph(doc, "tail");

    const reopened = openDocx(toUint8Array(doc));
    removeAllComments(reopened);
    removeAllFootnotes(reopened);
    removeAllEndnotes(reopened);
    removeAllImages(reopened);
    removeAllBookmarks(reopened);
    removeAllHyperlinks(reopened);
    removeAllHeaders(reopened);
    removeAllFooters(reopened);
    expectClean(toUint8Array(reopened));
  });
});
