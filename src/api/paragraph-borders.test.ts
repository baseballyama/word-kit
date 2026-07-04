import { describe, expect, it } from "vitest";
import {
  appendParagraph,
  createDocx,
  openDocx,
  paragraphs,
  setParagraphBorders,
  setParagraphShading,
  toUint8Array,
} from "./index.js";

function bodyXml(doc: ReturnType<typeof createDocx>): string {
  const reopened = openDocx(toUint8Array(doc));
  return new TextDecoder().decode(
    reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
  );
}

describe("setParagraphBorders", () => {
  it("draws all four sides by default", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "boxed");
    setParagraphBorders(paragraphs(doc).at(-1)!, {});
    const xml = bodyXml(doc);
    expect(xml).toContain("<w:pBdr>");
    for (const s of ["top", "left", "bottom", "right"]) {
      expect(xml).toContain(`<w:${s} w:val="single"`);
    }
  });

  it('supports a single bottom rule via sides:["bottom"]', () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "heading");
    setParagraphBorders(paragraphs(doc).at(-1)!, { sides: ["bottom"], sizeEighthsOfPoint: 12 });
    const xml = bodyXml(doc);
    expect(xml).toContain("<w:pBdr>");
    expect(xml).toMatch(/<w:bottom[^/]*w:sz="12"/);
    expect(xml).not.toContain("<w:top w:val");
    expect(xml).not.toContain("<w:left w:val");
    expect(xml).not.toContain("<w:right w:val");
  });

  it('"none" style is honoured', () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "p");
    setParagraphBorders(paragraphs(doc).at(-1)!, { style: "none" });
    const xml = bodyXml(doc);
    expect(xml).toMatch(/<w:top[^/]*w:val="none"/);
  });

  it("replaces an existing pBdr instead of stacking duplicates", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "p");
    const p = paragraphs(doc).at(-1)!;
    setParagraphBorders(p, { style: "single" });
    setParagraphBorders(p, { style: "dashed" });
    const xml = bodyXml(doc);
    const occurrences = (xml.match(/<w:pBdr>/g) || []).length;
    expect(occurrences).toBe(1);
    expect(xml).toMatch(/<w:top[^/]*w:val="dashed"/);
    expect(xml).not.toMatch(/<w:top[^/]*w:val="single"/);
  });
});

describe("setParagraphShading", () => {
  it("writes a single <w:shd> with the requested fill", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "callout");
    setParagraphShading(paragraphs(doc).at(-1)!, { fill: "FFE0E0" });
    const xml = bodyXml(doc);
    expect(xml).toMatch(/<w:shd[^/]*w:fill="FFE0E0"/);
  });

  it("replaces an existing shd rather than appending another", () => {
    const doc = createDocx({ paragraphs: [] });
    appendParagraph(doc, "callout");
    const p = paragraphs(doc).at(-1)!;
    setParagraphShading(p, { fill: "AAAAAA" });
    setParagraphShading(p, { fill: "BBBBBB" });
    const xml = bodyXml(doc);
    const occurrences = (xml.match(/<w:shd/g) || []).length;
    expect(occurrences).toBe(1);
    expect(xml).toMatch(/w:fill="BBBBBB"/);
  });
});
