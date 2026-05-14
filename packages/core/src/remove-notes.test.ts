import { getPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  addEndnote,
  addFootnote,
  createDocx,
  endnotesPart,
  footnotesPart,
  openDocx,
  paragraphs,
  removeAllEndnotes,
  removeAllFootnotes,
  toUint8Array,
} from "./docx.js";

describe("Docx.removeAllFootnotes", () => {
  it("removes user footnotes but keeps separator entries", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addFootnote(doc, para, "one");
    addFootnote(doc, para, "two");
    expect(footnotesPart(doc)?.footnotes).toHaveLength(4); // 2 separators + 2 user
    expect(removeAllFootnotes(doc)).toBe(2);
    expect(footnotesPart(doc)?.footnotes).toHaveLength(2); // only separators left
  });

  it("strips <w:footnoteReference> runs from the body", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addFootnote(doc, para, "ref");
    removeAllFootnotes(doc);
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      getPart(reopened.opc, "/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).not.toContain("footnoteReference");
  });

  it("returns 0 when there are no user footnotes", () => {
    const doc = createDocx();
    expect(removeAllFootnotes(doc)).toBe(0);
  });
});

describe("Docx.removeAllEndnotes", () => {
  it("removes user endnotes but keeps separator entries", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const para = paragraphs(doc)[0];
    if (!para) return;
    addEndnote(doc, para, "one");
    expect(removeAllEndnotes(doc)).toBe(1);
    expect(endnotesPart(doc)?.footnotes).toHaveLength(2); // separators only
  });
});
