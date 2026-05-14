import { getPart, hasPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import {
  addFooter,
  addHeader,
  createDocx,
  openDocx,
  setPageMargins,
  setPageOrientation,
  setPageSize,
  text,
  toUint8Array,
} from "./docx.js";
import { MARGINS_NORMAL, PAGE_SIZE_A4 } from "./index.js";

describe("Docx.addHeader / addFooter", () => {
  it("creates a header part, registers a rel, and references it in sectPr", () => {
    const doc = createDocx({ paragraphs: [] });
    const relId = addHeader(doc, "Document header");
    expect(relId).toMatch(/^rId\d+$/);
    expect(hasPart(doc.opc, "/word/header1.xml")).toBe(true);
    expect(doc.document.body.sectPr).toBeDefined();
    const sectPr = doc.document.body.sectPr;
    expect(
      sectPr?.children.some((c) => c.kind === "element" && c.name.local === "headerReference"),
    ).toBe(true);
  });

  it("creates a footer part the same way", () => {
    const doc = createDocx({ paragraphs: [] });
    addFooter(doc, "Document footer");
    expect(hasPart(doc.opc, "/word/footer1.xml")).toBe(true);
    const sectPr = doc.document.body.sectPr;
    expect(
      sectPr?.children.some((c) => c.kind === "element" && c.name.local === "footerReference"),
    ).toBe(true);
  });

  it("save+reopen preserves the header content", () => {
    const doc = createDocx({ paragraphs: ["body"] });
    addHeader(doc, "Top of page");
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    const headerPart = getPart(reopened.opc, "/word/header1.xml");
    expect(headerPart).toBeDefined();
    const headerXml = new TextDecoder().decode(headerPart?.data ?? new Uint8Array());
    expect(headerXml).toContain("Top of page");
  });

  it("allocates header2.xml / footer2.xml when called twice", () => {
    const doc = createDocx({ paragraphs: [] });
    addHeader(doc, "h1", "default");
    addHeader(doc, "h2", "first");
    expect(hasPart(doc.opc, "/word/header1.xml")).toBe(true);
    expect(hasPart(doc.opc, "/word/header2.xml")).toBe(true);
  });
});

describe("Docx page size + margins", () => {
  it("setPageSize writes a pgSz onto the body sectPr", () => {
    const doc = createDocx({ paragraphs: [] });
    setPageSize(doc, PAGE_SIZE_A4);
    const sectPr = doc.document.body.sectPr;
    expect(sectPr).toBeDefined();
    const pgSz = sectPr?.children.find((c) => c.kind === "element" && c.name.local === "pgSz");
    expect(pgSz).toBeDefined();
    if (pgSz?.kind !== "element") return;
    const w = pgSz.attrs.find((a) => a.name.local === "w")?.value;
    expect(w).toBe("11906");
  });

  it("setPageMargins writes a pgMar onto the body sectPr", () => {
    const doc = createDocx({ paragraphs: [] });
    setPageMargins(doc, MARGINS_NORMAL);
    const sectPr = doc.document.body.sectPr;
    const pgMar = sectPr?.children.find((c) => c.kind === "element" && c.name.local === "pgMar");
    expect(pgMar).toBeDefined();
  });

  it("setPageOrientation('landscape') swaps width/height appropriately", () => {
    const doc = createDocx({ paragraphs: [] });
    setPageSize(doc, PAGE_SIZE_A4);
    setPageOrientation(doc, "landscape");
    const sectPr = doc.document.body.sectPr;
    const pgSz = sectPr?.children.find((c) => c.kind === "element" && c.name.local === "pgSz");
    if (pgSz?.kind !== "element") {
      expect.fail("no pgSz");
      return;
    }
    const w = Number.parseInt(pgSz.attrs.find((a) => a.name.local === "w")?.value ?? "0", 10);
    const h = Number.parseInt(pgSz.attrs.find((a) => a.name.local === "h")?.value ?? "0", 10);
    expect(w).toBeGreaterThan(h);
    expect(pgSz.attrs.find((a) => a.name.local === "orient")?.value).toBe("landscape");
  });

  it("round-trips section changes through save+reopen", () => {
    const doc = createDocx({ paragraphs: ["hi"] });
    setPageSize(doc, PAGE_SIZE_A4);
    setPageOrientation(doc, "landscape");
    setPageMargins(doc, MARGINS_NORMAL);
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    expect(reopened.document.body.sectPr).toBeDefined();
    expect(text(reopened)).toBe("hi");
  });
});
