import { getPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  addFooter,
  addHeader,
  addImage,
  createDocx,
  footers,
  headers,
  images,
  openDocx,
  replaceImage,
  toUint8Array,
} from "./docx.js";

const TINY_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xfa, 0xcf, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);

const OTHER_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x42,
]);

describe("Docx.headers / Docx.footers", () => {
  it("enumerates header parts with their text content", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addHeader(doc, "Hello Header");
    addHeader(doc, "Second", "first");
    const heads = headers(doc);
    expect(heads).toHaveLength(2);
    expect(heads[0]?.text).toContain("Hello Header");
    expect(heads[1]?.text).toContain("Second");
  });

  it("enumerates footer parts with their text content", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addFooter(doc, "Page Footer");
    const foots = footers(doc);
    expect(foots).toHaveLength(1);
    expect(foots[0]?.text).toContain("Page Footer");
  });
});

describe("Docx.images / Docx.replaceImage", () => {
  it("lists media parts and returns each part's bytes", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, TINY_PNG, { widthEmu: 1000, heightEmu: 1000 });
    const imgs = images(doc);
    expect(imgs).toHaveLength(1);
    expect(imgs[0]?.partName).toBe("/word/media/image1.png");
    expect(imgs[0]?.contentType).toBe("image/png");
    expect(imgs[0]?.data.length).toBe(TINY_PNG.length);
  });

  it("replaces an existing image's bytes in place", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, TINY_PNG, { widthEmu: 1000, heightEmu: 1000 });
    expect(replaceImage(doc, "/word/media/image1.png", OTHER_PNG)).toBe(true);
    const part = getPart(doc.opc, "/word/media/image1.png");
    expect(part?.data.length).toBe(OTHER_PNG.length);

    // Survives save+reopen with the new bytes:
    const reopened = openDocx(toUint8Array(doc));
    expect(getPart(reopened.opc, "/word/media/image1.png")?.data.length).toBe(OTHER_PNG.length);
  });

  it("replaceImage returns false for unknown parts", () => {
    const doc = createDocx({ paragraphs: [] });
    expect(replaceImage(doc, "/word/media/imageXX.png", OTHER_PNG)).toBe(false);
  });
});
