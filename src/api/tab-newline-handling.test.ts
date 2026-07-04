// User text often contains \t and \n. Word's <w:t> renders them
// literally as the underlying Unicode code points and does NOT show
// visual tab stops or line breaks. The OOXML-correct shape is to
// emit <w:tab/> and <w:br/> children inside the same run:
//
//   "foo\tbar"   →   <w:r><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r>
//   "foo\nbar"   →   <w:r><w:t>foo</w:t><w:br/><w:t>bar</w:t></w:r>
//
// This file pins that contract.

import { describe, expect, it } from "vitest";
import {
  addBulletList,
  appendParagraph,
  appendTextRun,
  createDocx,
  openDocx,
  paragraphs,
  paragraphText,
  toUint8Array,
} from "./index.js";

function reopenXml(doc: ReturnType<typeof createDocx>): string {
  const reopened = openDocx(toUint8Array(doc));
  return new TextDecoder().decode(
    reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
  );
}

describe("tab and newline handling", () => {
  it("appendParagraph splits at \\t into <w:tab/> within a single run", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "foo\tbar");
    const xml = reopenXml(doc);
    expect(xml).toContain("<w:t>foo</w:t><w:tab/><w:t>bar</w:t>");
  });

  it("appendParagraph splits at \\n into <w:br/> within a single run", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "foo\nbar");
    const xml = reopenXml(doc);
    expect(xml).toContain("<w:t>foo</w:t><w:br/><w:t>bar</w:t>");
  });

  it("\\r\\n is treated as one break, not two", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "foo\r\nbar");
    const xml = reopenXml(doc);
    expect(xml).toContain("<w:t>foo</w:t><w:br/><w:t>bar</w:t>");
    expect(xml).not.toMatch(/<w:br\/><w:br\/>/);
    // No raw CR should leak into element text. The XML declaration's own
    // newline is `\r\n`, so we check the body specifically.
    const bodyMatch = xml.match(/<w:body>(.*)<\/w:body>/s);
    expect(bodyMatch?.[1]).not.toContain("\r");
  });

  it("paragraphText round-trips \\t and \\n as the underlying characters", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "foo\tbar\nbaz");
    const reopened = openDocx(toUint8Array(doc));
    const text = paragraphText(paragraphs(reopened)[0]!);
    expect(text).toBe("foo\tbar\nbaz");
  });

  it("addBulletList items with embedded \\n produce <w:br/>, not a new bullet", () => {
    const doc = createDocx({ paragraphs: [] });
    addBulletList(doc, ["one\ntwo"]);
    const xml = reopenXml(doc);
    expect(xml).toContain("<w:t>one</w:t><w:br/><w:t>two</w:t>");
    const numPrCount = (xml.match(/<w:numPr>/g) || []).length;
    expect(numPrCount).toBe(1);
  });

  it("appendTextRun formatting wraps the segmented pieces inside one run", () => {
    const doc = createDocx({ paragraphs: [""] });
    appendParagraph(doc, "");
    const target = paragraphs(doc).at(-1)!;
    appendTextRun(target, "a\tb", { bold: true });
    const xml = reopenXml(doc);
    // Exactly one <w:b/> attached to the run that wraps both text pieces.
    expect(xml).toMatch(/<w:r><w:rPr><w:b\/><\/w:rPr><w:t>a<\/w:t><w:tab\/><w:t>b<\/w:t><\/w:r>/);
  });

  it("a leading tab does not crash and starts the run with <w:tab/>", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "\thead");
    const xml = reopenXml(doc);
    expect(xml).toMatch(/<w:r><w:tab\/><w:t>head<\/w:t><\/w:r>/);
  });

  it("consecutive tabs emit consecutive <w:tab/> elements", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "a\t\tb");
    const xml = reopenXml(doc);
    expect(xml).toContain("<w:t>a</w:t><w:tab/><w:tab/><w:t>b</w:t>");
  });
});
