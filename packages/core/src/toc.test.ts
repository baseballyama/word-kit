import { describe, expect, it } from "vitest";
import {
  addTableOfContents,
  appendMergeField,
  appendParagraph,
  createDocx,
  openDocx,
  paragraphs,
  toUint8Array,
} from "./docx.js";

describe("addTableOfContents", () => {
  it("inserts a single paragraph holding a TOC field with the default 1-3 range", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    const before = paragraphs(doc).length;
    addTableOfContents(doc);
    const after = paragraphs(doc).length;
    expect(after).toBe(before + 1);
    const bytes = toUint8Array(doc);
    const xml = new TextDecoder().decode(bytes);
    // Re-open + serialise the body, then check the field instruction text
    // shows up. Splitting by part isn't necessary since toUint8Array zips
    // everything; we grep the bytes via Docx.open instead.
    const reopened = openDocx(bytes);
    const xmlText = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xmlText).toContain('TOC \\o "1-3" \\h \\z \\u');
    expect(xml.length).toBeGreaterThan(0);
  });

  it("honours a custom heading-level range and placeholder", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addTableOfContents(doc, {
      headingLevels: { from: 2, to: 4 },
      placeholderText: "[ToC here]",
    });
    const reopened = openDocx(toUint8Array(doc));
    const xmlText = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xmlText).toContain('TOC \\o "2-4" \\h \\z \\u');
    expect(xmlText).toContain("[ToC here]");
  });

  it("rejects invalid heading-level ranges", () => {
    const doc = createDocx({ paragraphs: [] });
    expect(() => addTableOfContents(doc, { headingLevels: { from: 0, to: 3 } })).toThrow(
      /Invalid TOC/,
    );
    expect(() => addTableOfContents(doc, { headingLevels: { from: 3, to: 2 } })).toThrow(
      /Invalid TOC/,
    );
    expect(() => addTableOfContents(doc, { headingLevels: { from: 1, to: 10 } })).toThrow(
      /Invalid TOC/,
    );
  });
});

describe("appendMergeField", () => {
  it("emits a MERGEFIELD complex-field run sequence with Word-style placeholder", () => {
    const doc = createDocx({ paragraphs: [""] });
    const p = paragraphs(doc)[0];
    if (!p) return;
    appendParagraph(doc, "Hello ");
    const target = paragraphs(doc)[1];
    if (!target) return;
    appendMergeField(doc, target, "FirstName");
    const reopened = openDocx(toUint8Array(doc));
    const xmlText = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xmlText).toContain("MERGEFIELD FirstName \\* MERGEFORMAT");
    expect(xmlText).toContain("«FirstName»");
  });

  it("allows overriding the placeholder text", () => {
    const doc = createDocx({ paragraphs: [""] });
    const p = paragraphs(doc)[0];
    if (!p) return;
    appendMergeField(doc, p, "Total", "$0.00");
    const reopened = openDocx(toUint8Array(doc));
    const xmlText = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xmlText).toContain("MERGEFIELD Total \\* MERGEFORMAT");
    expect(xmlText).toContain("$0.00");
  });

  it("rejects field names that aren't simple identifiers", () => {
    const doc = createDocx({ paragraphs: [""] });
    const p = paragraphs(doc)[0];
    if (!p) return;
    expect(() => appendMergeField(doc, p, "")).toThrow(/MERGEFIELD/);
    expect(() => appendMergeField(doc, p, "with space")).toThrow(/MERGEFIELD/);
    expect(() => appendMergeField(doc, p, "1startsWithDigit")).toThrow(/MERGEFIELD/);
  });
});
