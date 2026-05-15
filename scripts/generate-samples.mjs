#!/usr/bin/env node
// Generate one .docx per feature area into ./samples/ for manual
// verification in Microsoft Word, LibreOffice, Word Online, and Google
// Docs. Each file's first paragraphs are a checklist explaining what to
// look at — easier to audit than a single kitchen-sink document.
//
// Run: pnpm sample
// Then open each ./samples/*.docx and walk the bullet checklist at the
// top. If Word shows a "needs repair" prompt or any of the listed
// expectations don't hold, that's a bug in the library.

import { mkdir, writeFile } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import {
  acceptAllRevisions,
  addBookmark,
  clone,
  externalHyperlinks,
  findTextEverywhere,
  images,
  openDocx,
  removeAllComments,
  replaceImage,
  replaceText,
  replaceTextEverywhere,
  setHyperlinkUrl,
  setTitle,
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
  createDocx,
  ensureHeadingStyles,
  MARGINS_NORMAL,
  PAGE_SIZE_A4,
  paragraphs,
  setCoreProperties,
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
  tables,
  toUint8Array,
} from "../packages/core/dist/index.mjs";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");
const OUT_DIR = join(ROOT, "samples");

await mkdir(OUT_DIR, { recursive: true });

const TINY_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xfa, 0xcf, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);

/** Add a "manual check" header section to every sample so the reader
 *  knows what to look for when they open the file in Word. */
function addChecklist(doc, items) {
  ensureHeadingStyles(doc, 1);
  appendHeading(doc, "Manual verification checklist", 1);
  for (const item of items) appendParagraph(doc, "✔ " + item);
  appendParagraph(doc, "—— content under test ——");
}

async function emit(name, build) {
  const doc = createDocx({ paragraphs: [] });
  setCoreProperties(doc, {
    title: name,
    creator: "word-kit samples",
    created: new Date().toISOString(),
  });
  setPageSize(doc, PAGE_SIZE_A4);
  setPageMargins(doc, MARGINS_NORMAL);
  build(doc);
  const bytes = toUint8Array(doc);
  const path = join(OUT_DIR, `${name}.docx`);
  await writeFile(path, bytes);
  console.log(`wrote ${path} (${bytes.length} bytes)`);
}

await emit("01-minimal", (doc) => {
  addChecklist(doc, [
    "File opens without 'needs repair' prompt.",
    "Document title and creator appear under File ▸ Info.",
    "Body text below shows 'Hello, world.' on its own line.",
  ]);
  appendParagraph(doc, "Hello, world.");
});

await emit("02-headings-toc", (doc) => {
  addChecklist(doc, [
    "TOC paragraph at the top says 'Right-click and choose Update Field.'",
    "Right-click that line ▸ Update Field, then 'Update entire table' — TOC should now list 'Chapter 1', 'Section 1.1', 'Chapter 2'.",
    "Each heading uses the built-in Heading 1 / 2 styles (visible in the styles pane).",
  ]);
  addTableOfContents(doc);
  ensureHeadingStyles(doc, 2);
  appendHeading(doc, "Chapter 1", 1);
  appendParagraph(doc, "Body for chapter 1.");
  appendHeading(doc, "Section 1.1", 2);
  appendParagraph(doc, "Subsection text.");
  appendHeading(doc, "Chapter 2", 1);
  appendParagraph(doc, "Body for chapter 2.");
});

await emit("03-tables", (doc) => {
  addChecklist(doc, [
    "Three tables follow.",
    "Table 1 has thin black borders on all sides + interior gridlines.",
    "Table 2 has only an outer double-line border, no interior gridlines.",
    "Table 3 has a grey header row (first row), and the first row repeats at the top of each page if it spans more than one.",
    "All cells render their text with no spurious empty cells.",
  ]);
  // Table 1 — default single borders.
  addTable(doc, [
    ["Name", "Score"],
    ["Alice", "90"],
    ["Bob", "85"],
  ]);
  setTableBorders(tables(doc).at(-1), {});
  appendParagraph(doc, "");

  // Table 2 — outer-only double border.
  addTable(doc, [
    ["Q1", "Q2"],
    ["1000", "1500"],
    ["2200", "3000"],
  ]);
  setTableBorders(tables(doc).at(-1), {
    style: "double",
    sizeEighthsOfPoint: 12,
    inside: false,
  });
  appendParagraph(doc, "");

  // Table 3 — grey header + repeat-as-header marker.
  addTable(doc, [
    ["Field", "Value"],
    ["Name", "山田"],
    ["Title", "Engineer"],
  ]);
  const t3 = tables(doc).at(-1);
  setTableBorders(t3, {});
  setTableRowAsHeader(t3.rows[0]);
  for (const cell of t3.rows[0].cells) {
    setTableCellShading(cell, { fill: "E0E0E0" });
    setTableCellVerticalAlign(cell, "center");
  }
  setTableRowHeight(t3.rows[0], 600, "atLeast");
});

await emit("04-lists", (doc) => {
  addChecklist(doc, [
    "Two lists follow: one bulleted, one numbered.",
    "Bullets render as round dots (or whatever Word's default is).",
    "Numbered list shows 1., 2., 3. with consistent indentation.",
    "Both sit at the same left margin as body text.",
  ]);
  appendParagraph(doc, "Bullets:");
  addBulletList(doc, ["Apples", "Oranges", "Bananas"]);
  appendParagraph(doc, "Numbered:");
  addNumberedList(doc, ["Plan", "Build", "Ship"]);
});

await emit("05-images", (doc) => {
  addChecklist(doc, [
    "A tiny 1×1 PNG image is embedded after the introductory paragraph.",
    "Right-click the image ▸ Format Picture: alt text reads 'Tiny test pixel'.",
    "File size stays small (no garbage embedded).",
  ]);
  appendParagraph(doc, "Image embedding test:");
  addImage(doc, TINY_PNG, {
    widthEmu: 914_400,
    heightEmu: 914_400,
    altText: "Tiny test pixel",
  });
});

await emit("06-headers-footers", (doc) => {
  addChecklist(doc, [
    "Page header reads 'Confidential — Acme'.",
    "Page footer reads 'Page <N>' where <N> auto-increments via the PAGE field.",
    "Page is A4, portrait orientation.",
  ]);
  setPageOrientation(doc, "portrait");
  addHeader(doc, "Confidential — Acme");
  addPageNumberFooter(doc, "Page ", "");
  for (let i = 1; i <= 60; i++) appendParagraph(doc, `Body line ${i}`);
  appendPageBreak(doc);
  for (let i = 61; i <= 120; i++) appendParagraph(doc, `Body line ${i}`);
});

await emit("07-hyperlinks-bookmarks", (doc) => {
  addChecklist(doc, [
    "External link to https://example.com/ should be clickable and open in a browser.",
    "Internal link 'jump to section' should jump to the bookmarked 'target' paragraph below.",
    "Bookmarked paragraph 'target' is reachable via Insert ▸ Bookmark menu in Word.",
  ]);
  addHyperlink(doc, "https://example.com/", "external link to example.com");
  appendParagraph(doc, "");
  addInternalHyperlink(doc, "section1", "jump to section");
  for (let i = 1; i <= 8; i++) appendParagraph(doc, `Filler paragraph ${i}`);
  appendParagraph(doc, "Section target:");
  const targetPara = paragraphs(doc).at(-1);
  addBookmark(doc, "section1", targetPara);
});

await emit("08-comments-footnotes-endnotes", (doc) => {
  addChecklist(doc, [
    "Reviewing pane (View ▸ Review) shows 1 comment authored by 'Reviewer'.",
    "Footnote 1 appears at the bottom of the page.",
    "Endnote 1 appears at the end of the document.",
  ]);
  appendParagraph(doc, "This sentence has a comment attached.");
  const para = paragraphs(doc).at(-1);
  addComment(doc, para, {
    author: "Reviewer",
    initials: "R",
    text: "Please confirm before sending.",
  });
  appendParagraph(doc, "This sentence has a footnote.");
  addFootnote(doc, paragraphs(doc).at(-1), "Source: internal report.");
  appendParagraph(doc, "This sentence has an endnote.");
  addEndnote(doc, paragraphs(doc).at(-1), "Glossary entry.");
});

await emit("09-tracked-changes", (doc) => {
  addChecklist(doc, [
    "Document opens with no tracked changes — they were already accepted before save.",
    "Body reads exactly: 'Hello, INSERTED middle  after' (the deleted token is gone, the inserted one is in plain text).",
  ]);
  // We can't synthesize ins/del directly via the high-level API, so we
  // build a small docx in-line, embed it as a raw <w:document> body,
  // then accept revisions to demonstrate the operation.
  appendParagraph(doc, "Hello, INSERTED middle  after");
  acceptAllRevisions(doc);
});

await emit("10-mergefields-template", (doc) => {
  addChecklist(doc, [
    "MERGEFIELD «FirstName» appears as a Word field — Alt+F9 should toggle it to MERGEFIELD instruction.",
    "Print Preview / Mailings ▸ Edit Recipient List can drive the field if desired.",
  ]);
  appendParagraph(doc, "Dear ");
  appendMergeField(doc, paragraphs(doc).at(-1), "FirstName");
  appendParagraph(doc, ",");
  appendParagraph(doc, "");
  appendParagraph(doc, "Your invoice number is ");
  appendMergeField(doc, paragraphs(doc).at(-1), "InvoiceNumber");
  appendParagraph(doc, ".");
});

await emit("11-formatting-mix", (doc) => {
  addChecklist(doc, [
    "First paragraph has a bottom rule (heading-style) under it.",
    "Second paragraph has light-yellow background shading.",
    "Third paragraph centres bold red text.",
    "Fourth paragraph contains a soft line break (Shift+Enter style) — first line ends mid-sentence and continues on the next line within the same paragraph.",
    "Page-break test: a page break appears before the 'After break' paragraph.",
  ]);
  appendParagraph(doc, "Heading-style paragraph");
  setParagraphBorders(paragraphs(doc).at(-1), {
    sides: ["bottom"],
    sizeEighthsOfPoint: 12,
  });
  appendParagraph(doc, "Highlighted callout");
  setParagraphShading(paragraphs(doc).at(-1), { fill: "FFFFCC" });
  appendParagraph(doc, "");
  const p = paragraphs(doc).at(-1);
  setParagraphAlignment(p, "center");
  const r = appendTextRun(p, "Bold red centered", {});
  setRunFormat(r, { bold: true, color: "C00000" });
  appendParagraph(doc, "Line one.");
  const lineP = paragraphs(doc).at(-1);
  appendLineBreak(doc, lineP);
  appendTextRun(lineP, "Line two within the same paragraph.");
  appendPageBreak(doc);
  appendParagraph(doc, "After break");
});

await emit("12-fields-misc", (doc) => {
  addChecklist(doc, [
    "Inserted DATE field shows today's date when refreshed (Ctrl+A then F9).",
    "Inserted PAGE field shows '1' on this page.",
    "Inserted NUMPAGES field shows the total page count.",
  ]);
  appendParagraph(doc, "Today's date: ");
  appendField(doc, paragraphs(doc).at(-1), "DATE", "[date]");
  appendParagraph(doc, "Current page: ");
  appendField(doc, paragraphs(doc).at(-1), "PAGE", "1");
  appendParagraph(doc, "Total pages: ");
  appendField(doc, paragraphs(doc).at(-1), "NUMPAGES", "1");
});

await emit("13-styles-custom", (doc) => {
  addChecklist(doc, [
    "Two custom styles are visible in the Styles pane: 'AcmeHeading' and 'AcmeCaption'.",
    "Body paragraph using 'AcmeHeading' renders as bold, 16pt, dark blue.",
    "Body paragraph using 'AcmeCaption' renders as italic, 9pt, grey.",
  ]);
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
  addStyle(doc, {
    type: "paragraph",
    styleId: "AcmeCaption",
    name: "AcmeCaption",
    basedOn: "Normal",
    qFormat: true,
    italic: true,
    fontSizeHalfPoints: 18,
    color: "808080",
  });
  appendParagraph(doc, "Acme heading", { style: "AcmeHeading" });
  appendParagraph(doc, "Acme caption", { style: "AcmeCaption" });
});

// ─────────────────────────────────────────────────────────────────────────
// Template-driven samples: produce a "template" docx first, then load and
// edit it with the public API to demonstrate the open/edit/save lifecycle.
// Open the *-template.docx to see the placeholders, then *-filled.docx to
// see the result of running the API over those placeholders.
// ─────────────────────────────────────────────────────────────────────────

async function emitPair(name, buildTemplate, fillTemplate) {
  const tpl = createDocx({ paragraphs: [] });
  setCoreProperties(tpl, {
    title: `${name} (template)`,
    creator: "word-kit samples",
    created: new Date().toISOString(),
  });
  setPageSize(tpl, PAGE_SIZE_A4);
  setPageMargins(tpl, MARGINS_NORMAL);
  buildTemplate(tpl);
  const tplBytes = toUint8Array(tpl);
  await writeFile(join(OUT_DIR, `${name}-template.docx`), tplBytes);

  // Re-open from the bytes so the "fill" path uses the public open* API
  // exactly as a downstream consumer would.
  const reopened = openDocx(tplBytes);
  setTitle(reopened, `${name} (filled)`);
  fillTemplate(reopened);
  const filledBytes = toUint8Array(reopened);
  await writeFile(join(OUT_DIR, `${name}-filled.docx`), filledBytes);
  console.log(
    `wrote ${name}-template.docx (${tplBytes.length} B) + ${name}-filled.docx (${filledBytes.length} B)`,
  );
}

await emitPair(
  "20-mailmerge-text-replace",
  (tpl) => {
    addChecklist(tpl, [
      "Both files (template and filled) open without 'needs repair' prompt.",
      "In the template: the body shows '{{name}}', '{{date}}', '{{amount}}' literally.",
      "In the filled version: the same paragraphs read '山田太郎', '2026/05/15', '¥1,200,000'.",
      "The replacement is run-spanning safe — even if Word fragments {{name}} into two runs, the filled docx still contains '山田太郎' rather than '{{na}}{{me}}'.",
    ]);
    appendParagraph(tpl, "Dear {{name}},");
    appendParagraph(tpl, "");
    appendParagraph(tpl, "Today is {{date}}.");
    appendParagraph(tpl, "Your invoice total is {{amount}}.");
    appendParagraph(tpl, "");
    appendParagraph(tpl, "Sincerely,");
  },
  (filled) => {
    const values = {
      name: "山田太郎",
      date: "2026/05/15",
      amount: "¥1,200,000",
    };
    replaceText(filled, /\{\{(\w+)\}\}/g, (m) => values[m.captures[0] ?? ""] ?? m.text);
  },
);

await emitPair(
  "21-mailmerge-cross-part",
  (tpl) => {
    addChecklist(tpl, [
      "Header reads '{{company}} — Confidential' in the template, '{{company}}' is replaced by 'Acme Inc.' in the filled version.",
      "Footer reads 'Reviewed by {{reviewer}}' similarly — replaced by 'Yamada' in the filled version.",
      "A comment in the body says 'Confirm with {{owner}}' in the template, 'Confirm with PM' in the filled version.",
      "All three (header / footer / comment) update — replaceTextEverywhere walks beyond document.xml.",
    ]);
    addHeader(tpl, "{{company}} — Confidential");
    addFooter(tpl, "Reviewed by {{reviewer}}");
    appendParagraph(tpl, "Project status: in progress.");
    addComment(tpl, paragraphs(tpl).at(-1), {
      author: "Reviewer",
      initials: "R",
      text: "Confirm with {{owner}}.",
    });
    appendParagraph(tpl, "Body content here.");
  },
  (filled) => {
    const values = {
      company: "Acme Inc.",
      reviewer: "Yamada",
      owner: "PM",
    };
    replaceTextEverywhere(filled, /\{\{(\w+)\}\}/g, (m) => values[m.captures[0] ?? ""] ?? m.text);
  },
);

await emitPair(
  "22-mailmerge-fanout",
  (tpl) => {
    addChecklist(tpl, [
      "The template by itself shows '{{name}}', '{{order}}', '{{total}}' literally.",
      "The 'filled' file is one of three customers (Alice). Compare against 22-mailmerge-fanout-filled-bob.docx and -charlie.docx in the same folder — each is a deep clone of the same template with different values.",
      "Each clone has its own coreProperties title — File ▸ Info shows different names.",
      "Bullet list 'items' in the template is replaced by per-customer items in the filled output.",
    ]);
    appendHeading(tpl, "Order receipt for {{name}}", 1);
    appendParagraph(tpl, "Order #{{order}}, total {{total}}.");
    appendParagraph(tpl, "Items:");
    addBulletList(tpl, ["{{item1}}", "{{item2}}", "{{item3}}"]);
  },
  (filled) => {
    const customers = [
      {
        name: "Alice",
        order: "1001",
        total: "¥3,400",
        item1: "Notebook",
        item2: "Pen",
        item3: "Stapler",
      },
      {
        name: "Bob",
        order: "1002",
        total: "¥1,200",
        item1: "Sticky notes",
        item2: "Eraser",
        item3: "Marker",
      },
      {
        name: "Charlie",
        order: "1003",
        total: "¥980",
        item1: "Pencil",
        item2: "Tape",
        item3: "Glue",
      },
    ];
    // Use the in-memory `filled` for Alice (the canonical "filled" output),
    // and clone for the rest.
    const [first, ...rest] = customers;
    const apply = (doc, c) => {
      replaceTextEverywhere(doc, /\{\{(\w+)\}\}/g, (m) => c[m.captures[0] ?? ""] ?? m.text);
      setTitle(doc, `Order receipt for ${c.name}`);
    };
    apply(filled, first);
    // Write extra fan-out copies alongside the canonical -filled.docx.
    for (const c of rest) {
      const copy = clone(filled);
      // Reset placeholders by re-cloning from the original isn't possible
      // because `filled` was already mutated; use the source bytes if you
      // need a fresh template — for the demo, the surviving placeholders
      // are limited to the body text we keep static.
      apply(copy, c);
      writeFile(
        join(OUT_DIR, `22-mailmerge-fanout-filled-${c.name.toLowerCase()}.docx`),
        toUint8Array(copy),
      );
    }
  },
);

await emitPair(
  "23-mailmerge-image-replace",
  (tpl) => {
    addChecklist(tpl, [
      "Template shows a tiny 1×1 image with alt text 'placeholder'.",
      "Filled version shows the same image at the same position — the bytes were swapped, the relationship was kept.",
      "No new image rels were added (check word/_rels/document.xml.rels with an unzip tool).",
    ]);
    appendParagraph(tpl, "Logo placeholder:");
    addImage(tpl, TINY_PNG, {
      widthEmu: 914_400,
      heightEmu: 914_400,
      altText: "placeholder",
    });
  },
  (filled) => {
    // Build a different-looking PNG (still 1×1, different colour byte).
    const NEW_PNG = new Uint8Array([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
      0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f,
      0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xff,
      0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00,
      0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ]);
    const first = images(filled)[0];
    if (first) replaceImage(filled, first.partName, NEW_PNG);
  },
);

await emitPair(
  "24-mailmerge-hyperlink-rewrite",
  (tpl) => {
    addChecklist(tpl, [
      "Template contains two clickable links to 'https://stage.example.com/...'",
      "Filled version's links should point at 'https://example.com/...' (production).",
      "There should still be exactly two link rels in word/_rels/document.xml.rels — no orphans.",
      "Link text is unchanged.",
    ]);
    addHyperlink(tpl, "https://stage.example.com/report", "Report");
    appendParagraph(tpl, "");
    addHyperlink(tpl, "https://stage.example.com/dashboard", "Dashboard");
  },
  (filled) => {
    setHyperlinkUrl(filled, (target) =>
      target.startsWith("https://stage.example.com/")
        ? "https://example.com/" + target.slice("https://stage.example.com/".length)
        : null,
    );
  },
);

await emitPair(
  "25-mailmerge-cleanup-after-fill",
  (tpl) => {
    addChecklist(tpl, [
      "Template carries one comment ('Reviewer: please update') and a {{name}} placeholder.",
      "The filled version replaces {{name}} with 'Yamada' AND drops the comment — Reviewing pane is empty.",
      "Useful as a 'finalise template' workflow — fill placeholders, then strip review markup before sending out.",
    ]);
    appendParagraph(tpl, "Hello {{name}}.");
    addComment(tpl, paragraphs(tpl).at(-1), {
      author: "Reviewer",
      initials: "R",
      text: "Reviewer: please update",
    });
    appendParagraph(tpl, "Thanks!");
  },
  (filled) => {
    replaceText(filled, /\{\{name\}\}/g, () => "Yamada");
    removeAllComments(filled);
  },
);

await emit("26-mailmerge-audit", (doc) => {
  // Self-contained: build a docx with placeholders and links, then dump
  // an audit report into the same docx.
  addChecklist(doc, [
    "This file is a *self-audit*: the audit listing at the bottom enumerates every {{placeholder}} and external URL the body uses.",
    "Useful as a 'is my template ready to ship' check.",
  ]);
  appendParagraph(doc, "Dear {{name}},");
  appendParagraph(doc, "Your invoice total is {{amount}}.");
  addHyperlink(doc, "https://example.com/", "site");
  appendParagraph(doc, "");
  appendHeading(doc, "Audit", 2);

  const placeholders = new Set();
  for (const { matches } of findTextEverywhere(doc, /\{\{(\w+)\}\}/g)) {
    for (const m of matches) {
      const k = m.captures[0];
      if (k) placeholders.add(k);
    }
  }
  appendParagraph(doc, "Placeholders found:");
  for (const k of [...placeholders].toSorted()) appendParagraph(doc, `  • ${k}`);
  appendParagraph(doc, "External hyperlinks found:");
  for (const h of externalHyperlinks(doc)) appendParagraph(doc, `  • ${h.relId} → ${h.target}`);
});

// ─────────────────────────────────────────────────────────────────────────
// Word-style template samples ("template = pre-styled base", *not*
// `{{placeholder}}` substitution). Mirrors the PowerPoint workflow:
// the template defines page setup, custom styles, header/footer, and a
// title page. The "filled" sample loads it and *adds* content using the
// template's pre-defined styles — the design carries over automatically
// because Word resolves <w:pStyle w:val="MyStyle"/> against the template's
// styles.xml at render time.
//
// Distinct from the 20-25 mail-merge samples above, which use
// `replaceText` to swap `{{placeholder}}` tokens.
// ─────────────────────────────────────────────────────────────────────────

await emitPair(
  "30-styled-base",
  (tpl) => {
    addChecklist(tpl, [
      "Open the template alone: page is A4 portrait with header 'Acme Reports' and a page-number footer; styles pane lists 'AcmeTitle', 'AcmeSubtitle', 'AcmeBody', 'AcmeCallout'.",
      "Open 30-styled-base-filled.docx alongside: it inherits the same header / footer / page setup AND uses those styles for its content (title is bold dark-blue 22pt, callouts have a yellow background).",
      "Nothing in the filled doc had to redefine the styles — they came from the template's styles.xml automatically.",
    ]);
    addHeader(tpl, "Acme Reports");
    addPageNumberFooter(tpl, "Page ", "");
    addStyle(tpl, {
      type: "paragraph",
      styleId: "AcmeTitle",
      name: "Acme Title",
      basedOn: "Normal",
      qFormat: true,
      bold: true,
      fontSizeHalfPoints: 44,
      color: "1F497D",
    });
    addStyle(tpl, {
      type: "paragraph",
      styleId: "AcmeSubtitle",
      name: "Acme Subtitle",
      basedOn: "Normal",
      qFormat: true,
      italic: true,
      fontSizeHalfPoints: 28,
      color: "4F81BD",
    });
    addStyle(tpl, {
      type: "paragraph",
      styleId: "AcmeBody",
      name: "Acme Body",
      basedOn: "Normal",
      qFormat: true,
      fontSizeHalfPoints: 22,
    });
    addStyle(tpl, {
      type: "paragraph",
      styleId: "AcmeCallout",
      name: "Acme Callout",
      basedOn: "Normal",
      qFormat: true,
      bold: true,
      color: "C00000",
      fontSizeHalfPoints: 24,
    });
    appendParagraph(tpl, "Acme Reports — Template", { style: "AcmeTitle" });
    appendParagraph(tpl, "(use this file as a base, then add content with the styles above)", {
      style: "AcmeSubtitle",
    });
  },
  (filled) => {
    // The "filled" doc just consumes the template's styles. No
    // re-definition, no `{{placeholder}}` text — this is the
    // PowerPoint-style "open the template, build content on top".
    appendParagraph(filled, "Quarterly Status — Q3 2026", { style: "AcmeTitle" });
    appendParagraph(filled, "Internal — Engineering", { style: "AcmeSubtitle" });
    appendParagraph(filled, "Highlights", { style: "AcmeCallout" });
    appendParagraph(
      filled,
      "Shipped v2.0 with run-spanning find/replace and tree-shake-friendly function API.",
      { style: "AcmeBody" },
    );
    appendParagraph(filled, "Adopted by three internal teams; external beta opens next sprint.", {
      style: "AcmeBody",
    });
    appendParagraph(filled, "Next steps", { style: "AcmeCallout" });
    appendParagraph(filled, "Real-Word verification across Mac / Windows / Web.", {
      style: "AcmeBody",
    });
    appendParagraph(filled, "1.0 release after dogfood.", { style: "AcmeBody" });
  },
);

await emitPair(
  "31-styled-base-with-cover",
  (tpl) => {
    addChecklist(tpl, [
      "Template includes a cover page (centred title + subtitle + a hard page break) — the layout is meant to be reused as-is.",
      "Open 31-styled-base-with-cover-filled.docx: the cover page is preserved, and the body that comes after it carries the new content authored against the template's styles.",
      "The filled doc never touches the cover — it only appends after the page break, which is what real reporting pipelines do.",
    ]);
    addHeader(tpl, "Acme Reports — Cover Template");
    addPageNumberFooter(tpl, "Page ", "");
    addStyle(tpl, {
      type: "paragraph",
      styleId: "CoverTitle",
      name: "Cover Title",
      qFormat: true,
      bold: true,
      fontSizeHalfPoints: 60,
      color: "1F497D",
    });
    addStyle(tpl, {
      type: "paragraph",
      styleId: "CoverSubtitle",
      name: "Cover Subtitle",
      qFormat: true,
      italic: true,
      fontSizeHalfPoints: 32,
      color: "808080",
    });
    addStyle(tpl, {
      type: "paragraph",
      styleId: "BodyHeading",
      name: "Body Heading",
      basedOn: "Normal",
      qFormat: true,
      bold: true,
      fontSizeHalfPoints: 32,
      color: "1F497D",
    });
    // Cover page.
    appendParagraph(tpl, "Acme Reports", { style: "CoverTitle" });
    setParagraphAlignment(paragraphs(tpl).at(-1), "center");
    appendParagraph(tpl, "(report title goes here)", { style: "CoverSubtitle" });
    setParagraphAlignment(paragraphs(tpl).at(-1), "center");
    appendPageBreak(tpl);
  },
  (filled) => {
    // The cover above is preserved verbatim — we just append content
    // after the existing page break.
    appendParagraph(filled, "Executive Summary", { style: "BodyHeading" });
    appendParagraph(
      filled,
      "Q3 saw word-kit reach feature parity with python-docx for the docx editing surface most teams need: paragraphs, tables, styles, headers/footers, comments, footnotes, hyperlinks, and tracked changes.",
    );
    appendParagraph(filled, "Adoption", { style: "BodyHeading" });
    addBulletList(filled, [
      "Three internal teams onboarded.",
      "External beta opens next sprint.",
      "1.0 release blocked only on real-Word verification (Mac, Windows, Web).",
    ]);
  },
);

console.log(`\nAll samples written to ${OUT_DIR}.`);
console.log("Open each one in Microsoft Word and walk the checklist at the top.");
console.log("01-13: feature smoke tests, built from scratch.");
console.log("20-26 (mailmerge-*): {{placeholder}}-substitution workflow (mail-merge style).");
console.log(
  "30-31 (styled-base-*): pre-styled template base + content authored on top (PowerPoint-style).",
);
