import { describe, expect, it } from "vitest";
import { addTable, createDocx, openDocx, setTableBorders, tables, toUint8Array } from "./index.js";

describe("setTableBorders", () => {
  it("applies a default single-line border on all six sides", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [
      ["a", "b"],
      ["c", "d"],
    ]);
    const t = tables(doc)[0]!;
    setTableBorders(t, {});
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain("<w:tblBorders>");
    for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"]) {
      expect(xml).toContain(`<w:${side} w:val="single"`);
    }
  });

  it("supports outer-only by passing inside: false", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["x"]]);
    const t = tables(doc)[0]!;
    setTableBorders(t, { inside: false });
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain("<w:top");
    expect(xml).toContain("<w:bottom");
    expect(xml).not.toContain("<w:insideH");
    expect(xml).not.toContain("<w:insideV");
  });

  it("honours style / size / color overrides", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["x"]]);
    const t = tables(doc)[0]!;
    setTableBorders(t, { style: "double", sizeEighthsOfPoint: 12, color: "1F497D" });
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toMatch(/<w:top[^/]*w:val="double"[^/]*w:sz="12"[^/]*w:color="1F497D"/);
  });

  it("replaces an existing <w:tblBorders> rather than appending a second one", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["x"]]);
    const t = tables(doc)[0]!;
    setTableBorders(t, { style: "single" });
    setTableBorders(t, { style: "dashed" });
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    const occurrences = (xml.match(/<w:tblBorders>/g) || []).length;
    expect(occurrences).toBe(1);
    expect(xml).toMatch(/<w:top[^/]*w:val="dashed"/);
    expect(xml).not.toMatch(/<w:top[^/]*w:val="single"/);
  });

  it('"none" style writes through to the bytes', () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["x"]]);
    const t = tables(doc)[0]!;
    setTableBorders(t, { style: "none" });
    const reopened = openDocx(toUint8Array(doc));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toMatch(/<w:top[^/]*w:val="none"/);
  });
});
