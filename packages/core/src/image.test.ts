import { hasPart, partRelationships, relationshipsByType } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  addImage,
  createDocx,
  insertImageInto,
  openDocx,
  paragraphs,
  toUint8Array,
} from "./docx.js";

// A 1x1 transparent PNG.
const TINY_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xfa, 0xcf, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);

describe("Docx.insertImageInto", () => {
  it("appends an image run to an existing paragraph", () => {
    const doc = createDocx({ paragraphs: ["Logo:"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    const childrenBefore = para.children.length;
    insertImageInto(doc, para, TINY_PNG, { widthEmu: 914400, heightEmu: 914400 });
    expect(para.children.length).toBe(childrenBefore + 1);
    expect(hasPart(doc.opc, "/word/media/image1.png")).toBe(true);
  });
});

describe("Docx.addImage", () => {
  it("attaches an image as /word/media/imageN.png and adds a rel", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, TINY_PNG, { widthEmu: 914400, heightEmu: 914400 });
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    expect(hasPart(reopened.opc, "/word/media/image1.png")).toBe(true);
    const rels = partRelationships(reopened.opc, "/word/document.xml");
    const imageRels = relationshipsByType(
      rels,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    );
    expect(imageRels).toHaveLength(1);
    expect(imageRels[0]?.target).toBe("media/image1.png");
  });

  it("emits a <w:drawing> inside the inserted paragraph", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, TINY_PNG, { widthEmu: 914400, heightEmu: 914400, altText: "test alt" });
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    const para = paragraphs(reopened)[0];
    expect(para).toBeDefined();
    const run = para?.children[0];
    expect(run?.kind).toBe("run");
    if (run?.kind !== "run") return;
    const drawing = run.pieces[0];
    expect(drawing?.kind).toBe("drawing");
  });

  it("allocates fresh part names when called multiple times", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, TINY_PNG, { widthEmu: 100000, heightEmu: 100000 });
    addImage(doc, TINY_PNG, { widthEmu: 100000, heightEmu: 100000 });
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    expect(hasPart(reopened.opc, "/word/media/image1.png")).toBe(true);
    expect(hasPart(reopened.opc, "/word/media/image2.png")).toBe(true);
  });

  it("throws if the content type cannot be detected and none is supplied", () => {
    const doc = createDocx({ paragraphs: [] });
    const garbage = new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7]);
    expect(() => addImage(doc, garbage, { widthEmu: 100, heightEmu: 100 })).toThrow();
  });

  it("accepts an explicit content type override", () => {
    const doc = createDocx({ paragraphs: [] });
    addImage(doc, new Uint8Array([1, 2, 3]), {
      widthEmu: 100,
      heightEmu: 100,
      contentType: "image/png",
    });
    expect(hasPart(doc.opc, "/word/media/image1.png")).toBe(true);
  });
});
