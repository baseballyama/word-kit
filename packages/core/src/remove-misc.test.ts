import { hasPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  addBookmark,
  addImage,
  bookmarks,
  createDocx,
  images,
  paragraphs,
  removeAllBookmarks,
  removeAllImages,
} from "./docx.js";

const TINY_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xfa, 0xcf, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);

describe("Docx.removeAllImages", () => {
  it("drops every media part and image rel", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, TINY_PNG, { widthEmu: 1000, heightEmu: 1000 });
    addImage(doc, TINY_PNG, { widthEmu: 1000, heightEmu: 1000 });
    expect(images(doc)).toHaveLength(2);
    expect(removeAllImages(doc)).toBe(2);
    expect(images(doc)).toHaveLength(0);
    expect(hasPart(doc.opc, "/word/media/image1.png")).toBe(false);
  });

  it("returns 0 when there are no images", () => {
    const doc = createDocx();
    expect(removeAllImages(doc)).toBe(0);
  });
});

describe("Docx.removeAllBookmarks", () => {
  it("drops every bookmark start/end pair", () => {
    const doc = createDocx({ paragraphs: ["A", "B", "C"] });
    const [p1, p2, p3] = paragraphs(doc);
    if (!p1 || !p2 || !p3) return;
    addBookmark(doc, "a", p1);
    addBookmark(doc, "b", p2);
    addBookmark(doc, "c", p3);
    expect(bookmarks(doc)).toHaveLength(3);
    expect(removeAllBookmarks(doc)).toBe(3);
    expect(bookmarks(doc)).toHaveLength(0);
  });
});
