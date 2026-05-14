import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.appendPageBreak", () => {
  it("adds a paragraph with a page-break run", () => {
    const doc = Docx.create({ paragraphs: ["before"] });
    doc.appendPageBreak();
    doc.appendParagraph("after");
    expect(doc.paragraphs).toHaveLength(3);
    const bytes = doc.toUint8Array();
    const reopened = Docx.open(bytes);
    const part = reopened.opc.getPart("/word/document.xml");
    const xml = new TextDecoder().decode(part?.data ?? new Uint8Array());
    expect(xml).toContain('<w:br w:type="page"/>');
  });
});

describe("Docx.addBookmark + addInternalHyperlink", () => {
  it("places bookmarkStart/End around a paragraph and an anchor hyperlink to it", () => {
    const doc = Docx.create({ paragraphs: ["chapter heading"] });
    const target = doc.paragraphs[0];
    if (!target) return;
    const id = doc.addBookmark("ch1", target);
    expect(id).toBe(0);
    doc.appendParagraph("body");
    doc.addInternalHyperlink("ch1", "Jump to Chapter 1");

    // Verify start/end placement
    const first = target.children[0];
    const last = target.children.at(-1);
    expect(first?.kind).toBe("raw");
    expect(last?.kind).toBe("raw");
    if (first?.kind === "raw") expect(first.node.name.local).toBe("bookmarkStart");
    if (last?.kind === "raw") expect(last.node.name.local).toBe("bookmarkEnd");

    // Verify save+reopen preserves both
    const reopened = Docx.open(doc.toUint8Array());
    const part = reopened.opc.getPart("/word/document.xml");
    const xml = new TextDecoder().decode(part?.data ?? new Uint8Array());
    expect(xml).toContain('w:name="ch1"');
    expect(xml).toContain('w:anchor="ch1"');
  });

  it("allocates non-overlapping bookmark ids across multiple calls", () => {
    const doc = Docx.create({ paragraphs: ["a", "b", "c"] });
    const [p1, p2, p3] = doc.paragraphs;
    if (!p1 || !p2 || !p3) return;
    const id1 = doc.addBookmark("a", p1);
    const id2 = doc.addBookmark("b", p2);
    const id3 = doc.addBookmark("c", p3);
    expect(new Set([id1, id2, id3]).size).toBe(3);
  });
});
