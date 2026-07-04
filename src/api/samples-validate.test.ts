// Mirrors the building blocks used by `scripts/generate-samples.mjs` —
// produces each docx purely in memory and asserts that:
//
//   1. It opens cleanly (no parse errors).
//   2. validate() returns no `level: "error"` issues.
//
// Goal: catch regressions where a feature combination would produce a docx
// that Word flags as needing repair, *before* the human has to open every
// sample file. The samples themselves (under ./samples/) remain useful for
// real-Word verification because automated tests can't tell us what Word's
// renderer will actually do — but they should at least never produce a docx
// the validator rejects on its own.

import { describe, expect, it } from "vitest";
import {
  acceptAllRevisions,
  addBookmark,
  addBulletList,
  addComment,
  addEndnote,
  addFooter,
  addFootnote,
  addHeader,
  addHyperlink,
  addImage,
  addInternalHyperlink,
  addNumberedList,
  addPageNumberFooter,
  addStyle,
  addTable,
  addTableOfContents,
  appendField,
  appendHeading,
  appendLineBreak,
  appendMergeField,
  appendPageBreak,
  appendParagraph,
  appendTextRun,
  clone,
  createDocx,
  ensureHeadingStyles,
  externalHyperlinks,
  findTextEverywhere,
  images,
  MARGINS_NORMAL,
  openDocx,
  PAGE_SIZE_A4,
  paragraphs,
  removeAllComments,
  replaceImage,
  replaceText,
  replaceTextEverywhere,
  setCoreProperties,
  setHyperlinkUrl,
  setPageMargins,
  setPageOrientation,
  setPageSize,
  setParagraphAlignment,
  setParagraphBorders,
  setParagraphShading,
  setRunFormat,
  setTableBorders,
  setTableCellShading,
  setTableCellVerticalAlign,
  setTableRowAsHeader,
  setTableRowHeight,
  setTitle,
  tables,
  toUint8Array,
  validate,
  type Docx,
} from "./index.js";

const TINY_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xfa, 0xcf, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);

function freshDoc(): Docx {
  const doc = createDocx({ paragraphs: [] });
  setPageSize(doc, PAGE_SIZE_A4);
  setPageMargins(doc, MARGINS_NORMAL);
  setCoreProperties(doc, {
    title: "sample",
    creator: "word-kit samples test",
    created: new Date().toISOString(),
  });
  return doc;
}

function expectClean(doc: Docx): void {
  // Round-trip through bytes so we exercise the writer + reader pair too.
  const bytes = toUint8Array(doc);
  expect(bytes.length).toBeGreaterThan(0);
  const reopened = openDocx(bytes);
  const issues = validate(reopened);
  const errors = issues.filter((i) => i.level === "error");
  if (errors.length > 0) {
    throw new Error(
      `validate() returned errors after round-trip:\n${errors
        .map((e) => `  [${e.code}] ${e.message}`)
        .join("\n")}`,
    );
  }
}

describe("sample combinations validate clean after round-trip", () => {
  it("01-minimal", () => {
    const doc = freshDoc();
    appendParagraph(doc, "Hello, world.");
    expectClean(doc);
  });

  it("02-headings-toc", () => {
    const doc = freshDoc();
    addTableOfContents(doc);
    ensureHeadingStyles(doc, 2);
    appendHeading(doc, "Chapter 1", 1);
    appendParagraph(doc, "Body for chapter 1.");
    appendHeading(doc, "Section 1.1", 2);
    appendParagraph(doc, "Subsection text.");
    appendHeading(doc, "Chapter 2", 1);
    expectClean(doc);
  });

  it("03-tables", () => {
    const doc = freshDoc();
    addTable(doc, [
      ["Name", "Score"],
      ["Alice", "90"],
    ]);
    setTableBorders(tables(doc).at(-1)!, {});
    addTable(doc, [["Q1", "Q2"]]);
    setTableBorders(tables(doc).at(-1)!, {
      style: "double",
      sizeEighthsOfPoint: 12,
      inside: false,
    });
    addTable(doc, [
      ["Field", "Value"],
      ["Name", "山田"],
    ]);
    const t = tables(doc).at(-1)!;
    setTableBorders(t, {});
    setTableRowAsHeader(t.rows[0]!);
    for (const cell of t.rows[0]!.cells) {
      setTableCellShading(cell, { fill: "E0E0E0" });
      setTableCellVerticalAlign(cell, "center");
    }
    setTableRowHeight(t.rows[0]!, 600, "atLeast");
    expectClean(doc);
  });

  it("04-lists", () => {
    const doc = freshDoc();
    addBulletList(doc, ["Apples", "Oranges", "Bananas"]);
    addNumberedList(doc, ["Plan", "Build", "Ship"]);
    expectClean(doc);
  });

  it("05-images", () => {
    const doc = freshDoc();
    addImage(doc, TINY_PNG, {
      widthEmu: 914_400,
      heightEmu: 914_400,
      altText: "Tiny test pixel",
    });
    expectClean(doc);
  });

  it("06-headers-footers", () => {
    const doc = freshDoc();
    setPageOrientation(doc, "portrait");
    addHeader(doc, "Confidential — Acme");
    addPageNumberFooter(doc, "Page ", "");
    for (let i = 1; i <= 5; i++) appendParagraph(doc, `Body line ${i}`);
    appendPageBreak(doc);
    appendParagraph(doc, "After break");
    expectClean(doc);
  });

  it("07-hyperlinks-bookmarks", () => {
    const doc = freshDoc();
    addHyperlink(doc, "https://example.com/", "external link");
    addInternalHyperlink(doc, "section1", "jump to section");
    appendParagraph(doc, "Section target:");
    addBookmark(doc, "section1", paragraphs(doc).at(-1)!);
    expectClean(doc);
  });

  it("08-comments-footnotes-endnotes", () => {
    const doc = freshDoc();
    appendParagraph(doc, "Commented sentence.");
    addComment(doc, paragraphs(doc).at(-1)!, {
      author: "Reviewer",
      initials: "R",
      text: "Confirm.",
    });
    appendParagraph(doc, "Footnoted sentence.");
    addFootnote(doc, paragraphs(doc).at(-1)!, "Source.");
    appendParagraph(doc, "Endnoted sentence.");
    addEndnote(doc, paragraphs(doc).at(-1)!, "Glossary.");
    expectClean(doc);
  });

  it("09-tracked-changes-accept-all", () => {
    const doc = freshDoc();
    appendParagraph(doc, "Plain body.");
    acceptAllRevisions(doc);
    expectClean(doc);
  });

  it("10-mergefields", () => {
    const doc = freshDoc();
    appendParagraph(doc, "Dear ");
    appendMergeField(doc, paragraphs(doc).at(-1)!, "FirstName");
    expectClean(doc);
  });

  it("11-formatting-mix", () => {
    const doc = freshDoc();
    appendParagraph(doc, "Heading-style paragraph");
    setParagraphBorders(paragraphs(doc).at(-1)!, {
      sides: ["bottom"],
      sizeEighthsOfPoint: 12,
    });
    appendParagraph(doc, "Highlighted callout");
    setParagraphShading(paragraphs(doc).at(-1)!, { fill: "FFFFCC" });
    appendParagraph(doc, "");
    const p = paragraphs(doc).at(-1)!;
    setParagraphAlignment(p, "center");
    const r = appendTextRun(p, "Bold red centered");
    setRunFormat(r, { bold: true, color: "C00000" });
    appendParagraph(doc, "Line one.");
    const lineP = paragraphs(doc).at(-1)!;
    appendLineBreak(doc, lineP);
    appendTextRun(lineP, "Line two within the same paragraph.");
    appendPageBreak(doc);
    appendParagraph(doc, "After break");
    expectClean(doc);
  });

  it("12-fields-misc", () => {
    const doc = freshDoc();
    appendParagraph(doc, "Today's date: ");
    appendField(doc, paragraphs(doc).at(-1)!, "DATE", "[date]");
    appendParagraph(doc, "Current page: ");
    appendField(doc, paragraphs(doc).at(-1)!, "PAGE", "1");
    expectClean(doc);
  });

  it("13-styles-custom", () => {
    const doc = freshDoc();
    addStyle(doc, {
      type: "paragraph",
      styleId: "AcmeHeading",
      name: "AcmeHeading",
      basedOn: "Normal",
      qFormat: true,
      bold: true,
      fontSizeHalfPoints: 32,
      color: "1F497D",
    });
    appendParagraph(doc, "Acme heading", { style: "AcmeHeading" });
    expectClean(doc);
  });
});

describe("template-pair samples validate clean after round-trip", () => {
  function buildBaseTemplate(build: (doc: Docx) => void): Uint8Array {
    const tpl = freshDoc();
    build(tpl);
    return toUint8Array(tpl);
  }

  it("20-template-text-replace", () => {
    const tplBytes = buildBaseTemplate((tpl) => {
      appendParagraph(tpl, "Dear {{name}},");
      appendParagraph(tpl, "Today is {{date}}.");
    });
    const filled = openDocx(tplBytes);
    setTitle(filled, "filled");
    const values: Record<string, string> = { name: "山田太郎", date: "2026/05/15" };
    replaceText(filled, /\{\{(\w+)\}\}/g, (m) => values[m.captures[0] ?? ""] ?? m.text);
    expectClean(filled);
  });

  it("21-template-cross-part", () => {
    const tplBytes = buildBaseTemplate((tpl) => {
      addHeader(tpl, "{{company}} — Confidential");
      addFooter(tpl, "Reviewed by {{reviewer}}");
      appendParagraph(tpl, "Project status: in progress.");
      addComment(tpl, paragraphs(tpl).at(-1)!, {
        author: "Reviewer",
        initials: "R",
        text: "Confirm with {{owner}}.",
      });
    });
    const filled = openDocx(tplBytes);
    const values: Record<string, string> = {
      company: "Acme Inc.",
      reviewer: "Yamada",
      owner: "PM",
    };
    replaceTextEverywhere(filled, /\{\{(\w+)\}\}/g, (m) => values[m.captures[0] ?? ""] ?? m.text);
    expectClean(filled);
  });

  it("22-template-mail-merge-fanout", () => {
    const tplBytes = buildBaseTemplate((tpl) => {
      appendHeading(tpl, "Order receipt for {{name}}", 1);
      appendParagraph(tpl, "Order #{{order}}, total {{total}}.");
    });
    const filled = openDocx(tplBytes);
    const apply = (doc: Docx, c: Record<string, string>): void => {
      replaceTextEverywhere(doc, /\{\{(\w+)\}\}/g, (m) => c[m.captures[0] ?? ""] ?? m.text);
      setTitle(doc, `Order receipt for ${c.name}`);
    };
    apply(filled, { name: "Alice", order: "1001", total: "¥3,400" });
    expectClean(filled);
    const cloned = clone(filled);
    apply(cloned, { name: "Bob", order: "1002", total: "¥1,200" });
    expectClean(cloned);
  });

  it("23-template-image-replace", () => {
    const tplBytes = buildBaseTemplate((tpl) => {
      appendParagraph(tpl, "Logo placeholder:");
      addImage(tpl, TINY_PNG, {
        widthEmu: 914_400,
        heightEmu: 914_400,
        altText: "placeholder",
      });
    });
    const filled = openDocx(tplBytes);
    const first = images(filled)[0];
    if (first) replaceImage(filled, first.partName, TINY_PNG);
    expectClean(filled);
  });

  it("24-template-hyperlink-rewrite", () => {
    const tplBytes = buildBaseTemplate((tpl) => {
      addHyperlink(tpl, "https://stage.example.com/report", "Report");
      addHyperlink(tpl, "https://stage.example.com/dashboard", "Dashboard");
    });
    const filled = openDocx(tplBytes);
    setHyperlinkUrl(filled, (target) =>
      target.startsWith("https://stage.example.com/")
        ? "https://example.com/" + target.slice("https://stage.example.com/".length)
        : null,
    );
    expectClean(filled);
  });

  it("25-template-cleanup-after-fill", () => {
    const tplBytes = buildBaseTemplate((tpl) => {
      appendParagraph(tpl, "Hello {{name}}.");
      addComment(tpl, paragraphs(tpl).at(-1)!, {
        author: "Reviewer",
        initials: "R",
        text: "Reviewer: please update",
      });
    });
    const filled = openDocx(tplBytes);
    replaceText(filled, /\{\{name\}\}/g, () => "Yamada");
    removeAllComments(filled);
    expectClean(filled);
  });

  it("26-template-audit", () => {
    const doc = freshDoc();
    appendParagraph(doc, "Dear {{name}},");
    appendParagraph(doc, "Your invoice total is {{amount}}.");
    addHyperlink(doc, "https://example.com/", "site");
    appendHeading(doc, "Audit", 2);
    const placeholders = new Set<string>();
    for (const { matches } of findTextEverywhere(doc, /\{\{(\w+)\}\}/g)) {
      for (const m of matches) {
        const k = m.captures[0];
        if (k !== undefined) placeholders.add(k);
      }
    }
    appendParagraph(doc, "Placeholders found:");
    for (const k of [...placeholders].toSorted()) appendParagraph(doc, `  • ${k}`);
    appendParagraph(doc, "External hyperlinks found:");
    for (const h of externalHyperlinks(doc)) appendParagraph(doc, `  • ${h.relId} → ${h.target}`);
    expectClean(doc);
  });
});
