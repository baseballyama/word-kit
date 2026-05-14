import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

const TINY_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xfa, 0xcf, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);

describe("Docx.removeAllImages", () => {
  it("drops every media part and image rel", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addImage(TINY_PNG, { widthEmu: 1000, heightEmu: 1000 });
    doc.addImage(TINY_PNG, { widthEmu: 1000, heightEmu: 1000 });
    expect(doc.images).toHaveLength(2);
    expect(doc.removeAllImages()).toBe(2);
    expect(doc.images).toHaveLength(0);
    expect(doc.opc.hasPart("/word/media/image1.png")).toBe(false);
  });

  it("returns 0 when there are no images", () => {
    const doc = Docx.create();
    expect(doc.removeAllImages()).toBe(0);
  });
});

describe("Docx.removeAllBookmarks", () => {
  it("drops every bookmark start/end pair", () => {
    const doc = Docx.create({ paragraphs: ["A", "B", "C"] });
    const [p1, p2, p3] = doc.paragraphs;
    if (!p1 || !p2 || !p3) return;
    doc.addBookmark("a", p1);
    doc.addBookmark("b", p2);
    doc.addBookmark("c", p3);
    expect(doc.bookmarks).toHaveLength(3);
    expect(doc.removeAllBookmarks()).toBe(3);
    expect(doc.bookmarks).toHaveLength(0);
  });
});
