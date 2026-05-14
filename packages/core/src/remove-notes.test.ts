import { getPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import { Docx } from "./docx.js";

describe("Docx.removeAllFootnotes", () => {
  it("removes user footnotes but keeps separator entries", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addFootnote(para, "one");
    doc.addFootnote(para, "two");
    expect(doc.footnotesPart?.footnotes).toHaveLength(4); // 2 separators + 2 user
    expect(doc.removeAllFootnotes()).toBe(2);
    expect(doc.footnotesPart?.footnotes).toHaveLength(2); // only separators left
  });

  it("strips <w:footnoteReference> runs from the body", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addFootnote(para, "ref");
    doc.removeAllFootnotes();
    const reopened = Docx.open(doc.toUint8Array());
    const xml = new TextDecoder().decode(
      getPart(reopened.opc, "/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).not.toContain("footnoteReference");
  });

  it("returns 0 when there are no user footnotes", () => {
    const doc = Docx.create();
    expect(doc.removeAllFootnotes()).toBe(0);
  });
});

describe("Docx.removeAllEndnotes", () => {
  it("removes user endnotes but keeps separator entries", () => {
    const doc = Docx.create({ paragraphs: ["body"] });
    const para = doc.paragraphs[0];
    if (!para) return;
    doc.addEndnote(para, "one");
    expect(doc.removeAllEndnotes()).toBe(1);
    expect(doc.endnotesPart?.footnotes).toHaveLength(2); // separators only
  });
});
