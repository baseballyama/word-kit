import { describe, expect, it } from "vitest";
import {
  addImage,
  createDocx,
  imageReferences,
  images,
  openDocx,
  replaceImageByAltText,
  toUint8Array,
} from "./index.js";

const PNG_A = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xfa, 0xcf, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);

const PNG_B = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xff, 0x00, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);

describe("imageReferences", () => {
  it("returns one entry per inline drawing with alt text", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, PNG_A, { widthEmu: 914_400, heightEmu: 914_400, altText: "Logo" });
    addImage(doc, PNG_A, { widthEmu: 914_400, heightEmu: 914_400, altText: "Banner" });
    const refs = imageReferences(doc);
    expect(refs.map((r) => r.altText).toSorted()).toEqual(["Banner", "Logo"]);
    for (const r of refs) {
      expect(r.relId).toMatch(/^rId\d+$/);
      expect(r.partName).toMatch(/^\/word\/media\//);
    }
  });

  it("survives a save+reopen — alt text round-trips", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, PNG_A, { widthEmu: 914_400, heightEmu: 914_400, altText: "Logo" });
    const reopened = openDocx(toUint8Array(doc));
    const refs = imageReferences(reopened);
    expect(refs.map((r) => r.altText)).toEqual(["Logo"]);
  });
});

describe("replaceImageByAltText", () => {
  it("swaps the bytes of every image whose alt text matches", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, PNG_A, { widthEmu: 914_400, heightEmu: 914_400, altText: "Logo" });
    addImage(doc, PNG_A, { widthEmu: 914_400, heightEmu: 914_400, altText: "Banner" });
    const n = replaceImageByAltText(doc, "Logo", PNG_B);
    expect(n).toBe(1);
    const after = images(doc);
    const logoPart = imageReferences(doc).find((r) => r.altText === "Logo")?.partName;
    const logoImage = after.find((p) => p.partName === logoPart);
    expect(logoImage?.data).toEqual(PNG_B);
  });

  it("returns 0 for an unknown alt text", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, PNG_A, { widthEmu: 914_400, heightEmu: 914_400, altText: "Logo" });
    expect(replaceImageByAltText(doc, "NoSuchAlt", PNG_B)).toBe(0);
  });

  it("works after a save+reopen", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, PNG_A, { widthEmu: 914_400, heightEmu: 914_400, altText: "Logo" });
    const reopened = openDocx(toUint8Array(doc));
    const n = replaceImageByAltText(reopened, "Logo", PNG_B);
    expect(n).toBe(1);
    const final = openDocx(toUint8Array(reopened));
    expect(images(final)[0]?.data).toEqual(PNG_B);
  });
});
