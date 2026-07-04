// Automated coverage for the "styled-base template" workflow:
// designed `.docx` is opened, fresh content is appended, the
// produced docx still validates AND the new paragraphs reference
// the template's styles by id — so Word picks up the design at
// render time without any per-call style redefinition.

import { describe, expect, it } from "vitest";
import {
  addBulletList,
  addHeader,
  addPageNumberFooter,
  addStyle,
  appendParagraph,
  createDocx,
  ensureHeadingStyles,
  findStyleIdByName,
  getParagraphStyle,
  headers,
  footers,
  listStyles,
  openDocx,
  paragraphs,
  setCoreProperties,
  setPageMargins,
  setPageOrientation,
  setPageSize,
  toUint8Array,
  validate,
  type Docx,
  MARGINS_NORMAL,
  PAGE_SIZE_A4,
} from "./index.js";

function buildDesignedTemplate(): Uint8Array {
  const tpl = createDocx({ paragraphs: [] });
  setPageSize(tpl, PAGE_SIZE_A4);
  setPageMargins(tpl, MARGINS_NORMAL);
  setPageOrientation(tpl, "portrait");
  setCoreProperties(tpl, { title: "Acme Reports — Template", creator: "design team" });
  addHeader(tpl, "Acme Reports");
  addPageNumberFooter(tpl, "Page ", "");
  addStyle(tpl, {
    type: "paragraph",
    styleId: "AcmeTitle",
    name: "Acme Title",
    qFormat: true,
    bold: true,
    fontSizeHalfPoints: 44,
    color: "1F497D",
  });
  addStyle(tpl, {
    type: "paragraph",
    styleId: "AcmeSubtitle",
    name: "Acme Subtitle",
    qFormat: true,
    italic: true,
    fontSizeHalfPoints: 28,
    color: "4F81BD",
  });
  addStyle(tpl, {
    type: "paragraph",
    styleId: "AcmeBody",
    name: "Acme Body",
    qFormat: true,
    fontSizeHalfPoints: 22,
  });
  addStyle(tpl, {
    type: "paragraph",
    styleId: "AcmeCallout",
    name: "Acme Callout",
    qFormat: true,
    bold: true,
    color: "C00000",
    fontSizeHalfPoints: 24,
  });
  return toUint8Array(tpl);
}

function expectClean(doc: Docx): Docx {
  const reopened = openDocx(toUint8Array(doc));
  const errors = validate(reopened).filter((i) => i.level === "error");
  if (errors.length > 0) {
    throw new Error(
      `validate() returned errors:\n${errors.map((e) => `  [${e.code}] ${e.message}`).join("\n")}`,
    );
  }
  return reopened;
}

describe("styled-base template: open → appendParagraph with style id → save", () => {
  it("the template carries the four custom styles + header/footer + page setup", () => {
    const doc = openDocx(buildDesignedTemplate());
    const ids = listStyles(doc).map((s) => s.styleId);
    expect(ids).toEqual(
      expect.arrayContaining(["AcmeTitle", "AcmeSubtitle", "AcmeBody", "AcmeCallout"]),
    );
    expect(headers(doc)).toHaveLength(1);
    expect(footers(doc)).toHaveLength(1);
  });

  it("a freshly-appended paragraph carries the template's style id verbatim", () => {
    const doc = openDocx(buildDesignedTemplate());
    appendParagraph(doc, "Quarterly Status — Q3 2026", { style: "AcmeTitle" });
    const final = expectClean(doc);
    const last = paragraphs(final).at(-1)!;
    expect(getParagraphStyle(last)).toBe("AcmeTitle");
  });

  it("custom styles survive save+reopen — listStyles matches the template", () => {
    const doc = openDocx(buildDesignedTemplate());
    appendParagraph(doc, "Body 1", { style: "AcmeBody" });
    addBulletList(doc, ["one", "two"]);
    const final = expectClean(doc);
    const ids = listStyles(final).map((s) => s.styleId);
    for (const id of ["AcmeTitle", "AcmeSubtitle", "AcmeBody", "AcmeCallout"]) {
      expect(ids).toContain(id);
    }
  });

  it("findStyleIdByName resolves the template's localised style names", () => {
    const doc = openDocx(buildDesignedTemplate());
    expect(findStyleIdByName(doc, "Acme Title")).toBe("AcmeTitle");
    expect(findStyleIdByName(doc, "Acme Body")).toBe("AcmeBody");
    expect(findStyleIdByName(doc, "Nonexistent")).toBeUndefined();
  });

  it("header / footer survive the open → append → save cycle", () => {
    const doc = openDocx(buildDesignedTemplate());
    appendParagraph(doc, "body", { style: "AcmeBody" });
    const final = expectClean(doc);
    expect(headers(final)[0]?.text).toContain("Acme Reports");
    expect(footers(final)[0]?.text).toBeDefined();
  });

  it("page setup (size + margins + orientation) carries through", () => {
    const doc = openDocx(buildDesignedTemplate());
    appendParagraph(doc, "body", { style: "AcmeBody" });
    const final = expectClean(doc);
    const xml = new TextDecoder().decode(
      final.opc.parts.get("/word/document.xml")?.data ?? new Uint8Array(),
    );
    // A4 portrait dimensions — w=11906 h=16838 twips.
    expect(xml).toMatch(/<w:pgSz[^/]*w:w="11906"[^/]*w:h="16838"/);
  });

  it("can layer extra styles on top of the template (addStyle is additive)", () => {
    const doc = openDocx(buildDesignedTemplate());
    addStyle(doc, {
      type: "paragraph",
      styleId: "AcmeQuote",
      name: "Acme Quote",
      qFormat: true,
      italic: true,
      color: "606060",
    });
    appendParagraph(doc, "Quoted text", { style: "AcmeQuote" });
    const final = expectClean(doc);
    expect(listStyles(final).map((s) => s.styleId)).toContain("AcmeQuote");
    expect(getParagraphStyle(paragraphs(final).at(-1)!)).toBe("AcmeQuote");
  });

  it("dangling style reference (caller typo) is flagged as a warning, not an error", () => {
    const doc = openDocx(buildDesignedTemplate());
    appendParagraph(doc, "Heading", { style: "DoesNotExist" });
    const reopened = openDocx(toUint8Array(doc));
    const issues = validate(reopened);
    expect(issues.some((i) => i.code === "style-ref-missing" && i.level === "warning")).toBe(true);
    expect(issues.filter((i) => i.level === "error")).toHaveLength(0);
  });

  it("ensureHeadingStyles on top of a custom-styled template adds Heading 1/2/3 without conflict", () => {
    const doc = openDocx(buildDesignedTemplate());
    ensureHeadingStyles(doc, 3);
    const final = expectClean(doc);
    const ids = listStyles(final).map((s) => s.styleId);
    expect(ids).toEqual(expect.arrayContaining(["Heading1", "Heading2", "Heading3", "AcmeTitle"]));
  });

  it("a typical end-to-end build: cover paragraphs + body + bulleted list", () => {
    const doc = openDocx(buildDesignedTemplate());
    appendParagraph(doc, "Quarterly Status — Q3 2026", { style: "AcmeTitle" });
    appendParagraph(doc, "Internal — Engineering", { style: "AcmeSubtitle" });
    appendParagraph(doc, "Highlights", { style: "AcmeCallout" });
    appendParagraph(doc, "Shipped v2.0.", { style: "AcmeBody" });
    appendParagraph(doc, "Next steps", { style: "AcmeCallout" });
    addBulletList(doc, ["Real-Word verification", "1.0 release after dogfood"]);
    expectClean(doc);
  });
});
