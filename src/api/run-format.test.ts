import { describe, expect, it } from "vitest";
import {
  appendParagraph,
  appendTextRun,
  clearRunFormat,
  createDocx,
  openDocx,
  paragraphs,
  setRunFormat,
  toUint8Array,
  type WmlRun,
} from "./index.js";

function firstRun(doc: ReturnType<typeof createDocx>): WmlRun {
  const para = paragraphs(doc)[0];
  if (!para) throw new Error("no paragraph");
  for (const child of para.children) {
    if (child.kind === "run") return child;
  }
  throw new Error("no run in paragraph");
}

describe("setRunFormat", () => {
  it("adds bold/italic/underline/color to a run without an rPr", () => {
    const doc = createDocx({ paragraphs: ["plain"] });
    const run = firstRun(doc);
    expect(run.rPr).toBeUndefined();
    setRunFormat(run, { bold: true, italic: true, underline: "single", color: "FF0000" });
    expect(run.rPr).toBeDefined();
    const localNames = (run.rPr?.children ?? []).map((c) => c.kind === "element" && c.name.local);
    expect(localNames).toEqual(expect.arrayContaining(["b", "i", "u", "color"]));
  });

  it("replaces existing color rather than appending a duplicate", () => {
    const doc = createDocx({ paragraphs: [""] });
    appendParagraph(doc, "");
    const target = paragraphs(doc)[1]!;
    const run = appendTextRun(target, "hello", { color: "000000" });
    setRunFormat(run, { color: "FF0000" });
    const colors = (run.rPr?.children ?? []).filter(
      (c) => c.kind === "element" && c.name.local === "color",
    );
    expect(colors).toHaveLength(1);
    expect(colors[0]).toMatchObject({ attrs: [{ value: "FF0000" }] });
  });

  it("bool=false removes the corresponding element", () => {
    const doc = createDocx({ paragraphs: [""] });
    appendParagraph(doc, "");
    const target = paragraphs(doc)[1]!;
    const run = appendTextRun(target, "hi", { bold: true, italic: true });
    setRunFormat(run, { bold: false });
    const bs = (run.rPr?.children ?? []).filter(
      (c) => c.kind === "element" && c.name.local === "b",
    );
    expect(bs).toHaveLength(0);
    // italic preserved
    const is = (run.rPr?.children ?? []).filter(
      (c) => c.kind === "element" && c.name.local === "i",
    );
    expect(is).toHaveLength(1);
  });

  it('underline="none" removes the underline element', () => {
    const doc = createDocx({ paragraphs: [""] });
    appendParagraph(doc, "");
    const target = paragraphs(doc)[1]!;
    const run = appendTextRun(target, "hi", { underline: "single" });
    setRunFormat(run, { underline: "none" });
    const us = (run.rPr?.children ?? []).filter(
      (c) => c.kind === "element" && c.name.local === "u",
    );
    expect(us).toHaveLength(0);
  });

  it("merges rFonts attrs instead of replacing the whole element", () => {
    const doc = createDocx({ paragraphs: [""] });
    appendParagraph(doc, "");
    const target = paragraphs(doc)[1]!;
    const run = appendTextRun(target, "hi", { font: "Arial", fontEastAsia: "Yu Mincho" });
    setRunFormat(run, { font: "Helvetica" });
    const rFonts = (run.rPr?.children ?? []).find(
      (c) => c.kind === "element" && c.name.local === "rFonts",
    );
    expect(rFonts).toBeDefined();
    const attrs = (rFonts as { attrs: { name: { local: string }; value: string }[] }).attrs;
    expect(attrs.find((a) => a.name.local === "ascii")?.value).toBe("Helvetica");
    expect(attrs.find((a) => a.name.local === "hAnsi")?.value).toBe("Helvetica");
    // eastAsia preserved from the prior call.
    expect(attrs.find((a) => a.name.local === "eastAsia")?.value).toBe("Yu Mincho");
  });

  it("survives a save+reopen with the formatting in the bytes", () => {
    const doc = createDocx({ paragraphs: ["plain"] });
    const run = firstRun(doc);
    setRunFormat(run, { bold: true, color: "1F497D" });
    const reopened = openDocx(toUint8Array(doc));
    const xmlText = new TextDecoder().decode(
      reopened.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xmlText).toContain("<w:b/>");
    expect(xmlText).toContain('<w:color w:val="1F497D"');
  });

  it("is a no-op when every field is undefined", () => {
    const doc = createDocx({ paragraphs: ["plain"] });
    const run = firstRun(doc);
    setRunFormat(run, {});
    expect(run.rPr).toBeUndefined();
  });
});

describe("clearRunFormat", () => {
  it("removes the run's <w:rPr> entirely", () => {
    const doc = createDocx({ paragraphs: [""] });
    appendParagraph(doc, "");
    const target = paragraphs(doc)[1]!;
    const run = appendTextRun(target, "hi", { bold: true, italic: true });
    expect(run.rPr).toBeDefined();
    clearRunFormat(run);
    expect(run.rPr).toBeUndefined();
  });
});
