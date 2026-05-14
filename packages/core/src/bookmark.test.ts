import { getPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  addBookmark,
  addInternalHyperlink,
  appendPageBreak,
  appendParagraph,
  createDocx,
  openDocx,
  paragraphs,
  toUint8Array,
} from "./docx.js";

describe("Docx.appendPageBreak", () => {
  it("adds a paragraph with a page-break run", () => {
    const doc = createDocx({ paragraphs: ["before"] });
    appendPageBreak(doc);
    appendParagraph(doc, "after");
    expect(paragraphs(doc)).toHaveLength(3);
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    const part = getPart(reopened.opc, "/word/document.xml");
    const xml = new TextDecoder().decode(part?.data ?? new Uint8Array());
    expect(xml).toContain('<w:br w:type="page"/>');
  });
});

describe("Docx.addBookmark + addInternalHyperlink", () => {
  it("places bookmarkStart/End around a paragraph and an anchor hyperlink to it", () => {
    const doc = createDocx({ paragraphs: ["chapter heading"] });
    const target = paragraphs(doc)[0];
    if (!target) return;
    const id = addBookmark(doc, "ch1", target);
    expect(id).toBe(0);
    appendParagraph(doc, "body");
    addInternalHyperlink(doc, "ch1", "Jump to Chapter 1");

    // Verify start/end placement
    const first = target.children[0];
    const last = target.children.at(-1);
    expect(first?.kind).toBe("raw");
    expect(last?.kind).toBe("raw");
    if (first?.kind === "raw") expect(first.node.name.local).toBe("bookmarkStart");
    if (last?.kind === "raw") expect(last.node.name.local).toBe("bookmarkEnd");

    // Verify save+reopen preserves both
    const reopened = openDocx(toUint8Array(doc));
    const part = getPart(reopened.opc, "/word/document.xml");
    const xml = new TextDecoder().decode(part?.data ?? new Uint8Array());
    expect(xml).toContain('w:name="ch1"');
    expect(xml).toContain('w:anchor="ch1"');
  });

  it("allocates non-overlapping bookmark ids across multiple calls", () => {
    const doc = createDocx({ paragraphs: ["a", "b", "c"] });
    const [p1, p2, p3] = paragraphs(doc);
    if (!p1 || !p2 || !p3) return;
    const id1 = addBookmark(doc, "a", p1);
    const id2 = addBookmark(doc, "b", p2);
    const id3 = addBookmark(doc, "c", p3);
    expect(new Set([id1, id2, id3]).size).toBe(3);
  });
});
