// Re-save stability audit. After a save+open, calling toUint8Array
// again WITHOUT any mutation should produce identical bytes (modulo
// metadata that contains a timestamp). This is the property templating
// pipelines rely on: opening a docx and re-emitting it as-is should not
// silently change its contents.
//
// We measure stability by comparing the bytes of two consecutive
// save+open cycles. If they differ, the library is non-deterministic
// or re-flushes side parts unnecessarily, and we'd want to know.

import { describe, expect, it } from "vitest";
import {
  addBookmark,
  addBulletList,
  addComment,
  addEndnote,
  addFooter,
  addFootnote,
  addHeader,
  addHyperlink,
  addImage,
  addStyle,
  addTable,
  appendHeading,
  appendParagraph,
  createDocx,
  openDocx,
  paragraphs,
  toUint8Array,
} from "./index.js";

const TINY_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xfa, 0xcf, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);

function reSaveCycle(bytes: Uint8Array): Uint8Array {
  return toUint8Array(openDocx(bytes));
}

describe("idempotent re-save", () => {
  it("a freshly created minimal doc round-trips to the same bytes twice", () => {
    const seed = createDocx({ paragraphs: ["hello"] });
    const bytesA = toUint8Array(seed);
    const bytesB = reSaveCycle(bytesA);
    const bytesC = reSaveCycle(bytesB);
    // First save bakes the doc-properties (created timestamp etc.) so
    // bytesA may differ from bytesB — that's expected since the seed has
    // a fresh `created`. The interesting invariant is that subsequent
    // re-saves of *the same* bytes are stable.
    expect(bytesB).toEqual(bytesC);
  });

  it("a doc with every side part round-trips byte-stable", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    appendHeading(seed, "Title", 1);
    appendParagraph(seed, "body 1");
    addBulletList(seed, ["a", "b"]);
    addTable(seed, [
      ["A", "B"],
      ["1", "2"],
    ]);
    addImage(seed, TINY_PNG, { widthEmu: 914_400, heightEmu: 914_400 });
    addHeader(seed, "Header");
    addFooter(seed, "Footer");
    addHyperlink(seed, "https://example.com/", "link");
    addBookmark(seed, "anchor", paragraphs(seed)[0]!);
    addComment(seed, paragraphs(seed)[0]!, { author: "R", initials: "R", text: "c" });
    addFootnote(seed, paragraphs(seed)[0]!, "fn");
    addEndnote(seed, paragraphs(seed)[0]!, "en");
    addStyle(seed, { type: "paragraph", styleId: "MyHeading", name: "MyHeading" });

    const bytesA = toUint8Array(seed);
    const bytesB = reSaveCycle(bytesA);
    const bytesC = reSaveCycle(bytesB);
    expect(bytesB).toEqual(bytesC);
  });

  it("re-saving without any mutation does not grow the file", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    addTable(seed, [["x"]]);
    const bytesA = toUint8Array(seed);
    const bytesB = reSaveCycle(bytesA);
    const bytesC = reSaveCycle(bytesB);
    // Stable size between B and C (allow A→B to vary slightly).
    expect(bytesC.length).toBe(bytesB.length);
  });

  it("a single mutation produces exactly one new byte stream then stabilises", () => {
    const seed = createDocx({ paragraphs: ["body"] });
    const bytesA = toUint8Array(seed);
    const reopened = openDocx(bytesA);
    appendParagraph(reopened, "extra");
    const bytesB = toUint8Array(reopened);
    const bytesC = reSaveCycle(bytesB);
    expect(bytesC).toEqual(bytesB);
  });
});
