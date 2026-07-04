import { describe, expect, it } from "vitest";
import {
  appendLineBreak,
  appendParagraph,
  createDocx,
  openDocx,
  paragraphs,
  toUint8Array,
} from "./docx.js";

describe("appendLineBreak", () => {
  it("appends a soft <w:br/> to an existing paragraph by default", () => {
    const doc = createDocx({ paragraphs: ["hello"] });
    const p = paragraphs(doc)[0]!;
    appendLineBreak(doc, p);
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    // Plain <w:br/> with no type attribute.
    expect(xml).toContain("<w:br/>");
    // No `w:type="page"` came along for the ride.
    expect(xml).not.toMatch(/<w:br[^/]*type="page"/);
  });

  it('emits w:type="page" for a mid-paragraph page break', () => {
    const doc = createDocx({ paragraphs: ["before"] });
    const p = paragraphs(doc)[0]!;
    appendLineBreak(doc, p, "page");
    appendParagraph(doc, "after");
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toMatch(/<w:br\s+w:type="page"\/>/);
  });

  it('emits w:type="column" for a column break', () => {
    const doc = createDocx({ paragraphs: ["col"] });
    const p = paragraphs(doc)[0]!;
    appendLineBreak(doc, p, "column");
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toMatch(/<w:br\s+w:type="column"\/>/);
  });
});
