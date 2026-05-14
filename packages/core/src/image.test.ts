import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

// A 1x1 transparent PNG.
const TINY_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xfa, 0xcf, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);

describe("Docx.addImage", () => {
  it("attaches an image as /word/media/imageN.png and adds a rel", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addImage(TINY_PNG, { widthEmu: 914400, heightEmu: 914400 });
    const bytes = doc.toUint8Array();
    const reopened = Docx.open(bytes);
    expect(reopened.opc.hasPart("/word/media/image1.png")).toBe(true);
    const rels = reopened.opc.partRelationships("/word/document.xml");
    const imageRels = rels.byType(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    );
    expect(imageRels).toHaveLength(1);
    expect(imageRels[0]?.target).toBe("media/image1.png");
  });

  it("emits a <w:drawing> inside the inserted paragraph", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addImage(TINY_PNG, { widthEmu: 914400, heightEmu: 914400, altText: "test alt" });
    const bytes = doc.toUint8Array();
    const reopened = Docx.open(bytes);
    const para = reopened.paragraphs[0];
    expect(para).toBeDefined();
    const run = para?.children[0];
    expect(run?.kind).toBe("run");
    if (run?.kind !== "run") return;
    const drawing = run.pieces[0];
    expect(drawing?.kind).toBe("drawing");
  });

  it("allocates fresh part names when called multiple times", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addImage(TINY_PNG, { widthEmu: 100000, heightEmu: 100000 });
    doc.addImage(TINY_PNG, { widthEmu: 100000, heightEmu: 100000 });
    const bytes = doc.toUint8Array();
    const reopened = Docx.open(bytes);
    expect(reopened.opc.hasPart("/word/media/image1.png")).toBe(true);
    expect(reopened.opc.hasPart("/word/media/image2.png")).toBe(true);
  });

  it("throws if the content type cannot be detected and none is supplied", () => {
    const doc = Docx.create({ paragraphs: [] });
    const garbage = new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7]);
    expect(() => doc.addImage(garbage, { widthEmu: 100, heightEmu: 100 })).toThrow();
  });

  it("accepts an explicit content type override", () => {
    const doc = Docx.create({ paragraphs: [] });
    doc.addImage(new Uint8Array([1, 2, 3]), {
      widthEmu: 100,
      heightEmu: 100,
      contentType: "image/png",
    });
    expect(doc.opc.hasPart("/word/media/image1.png")).toBe(true);
  });
});
