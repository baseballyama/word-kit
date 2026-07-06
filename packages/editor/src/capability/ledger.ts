/**
 * The capability ledger — the mechanism that makes "cover every docx
 * representation" a *checkable* claim rather than an aspiration.
 *
 * Every element in the ECMA-376 schema universe ({@link ./element-universe.json})
 * is classified into exactly one {@link Disposition}:
 *
 * - `edit`    — a real editor command creates/modifies it (command id recorded).
 * - `render`  — the canvas renderer shows it faithfully, read-only.
 * - `preserve`— round-tripped losslessly by `@office-kit/docx` (raw XML
 *               pass-through), not yet surfaced in the UI. Each preserve
 *               classification cites the round-trip test that proves losslessness.
 *
 * `coverage.test.ts` asserts that (1) every universe element resolves to a
 * disposition, (2) every `edit` element names a command that actually exists,
 * so no feature can silently fall through and no command can be deleted without
 * the ledger noticing.
 */

import { PARA_PROPERTY_ELEMENTS, RUN_PROPERTY_ELEMENTS } from "../commands/properties.js";
import { COMMAND_IDS } from "../commands/registry.js";
import { NUMBERING_PROPERTY_ELEMENTS } from "../commands/numbering-props.js";
import { SECTION_PROPERTY_ELEMENTS } from "../commands/section-props.js";
import { SETTINGS_ELEMENTS } from "../commands/settings.js";
import { STYLE_PROPERTY_ELEMENTS } from "../commands/style-props.js";
import { TABLE_PROPERTY_BINDINGS } from "../commands/table-props.js";
import universe from "./element-universe.json" with { type: "json" };

export type Disposition = "edit" | "render" | "preserve";

export interface Capability {
  /** Qualified element name, e.g. `w:p`. */
  readonly element: string;
  readonly domain: string;
  readonly disposition: Disposition;
  /** For `edit`: the command id that handles the element. */
  readonly command?: string;
  /** Human rationale (why this disposition; what proves it for `preserve`). */
  readonly notes?: string;
}

interface UniverseElement {
  readonly element: string;
  readonly name: string;
  readonly prefix: string;
  readonly domain: string;
}

const ELEMENTS = (universe as { elements: UniverseElement[] }).elements;

/**
 * Elements a command creates or modifies → the command id. Keys are qualified
 * names from the universe (`w:b`, `w:jc`, …). When several commands touch one
 * element, the most representative is named.
 */
const EDIT_MAP: Record<string, string> = {
  // --- run text + character formatting (commands/text.ts) ---
  "w:r": "structure.insertParagraph",
  "w:t": "structure.insertParagraph",
  "w:rPr": "text.bold",
  "w:b": "text.bold",
  "w:bCs": "text.bold",
  "w:i": "text.italic",
  "w:iCs": "text.italic",
  "w:strike": "text.strike",
  "w:u": "text.underline",
  "w:color": "text.color",
  "w:highlight": "text.highlight",
  "w:sz": "text.fontSize",
  "w:szCs": "text.fontSize",
  "w:rFonts": "text.font",
  "w:br": "structure.insertLineBreak",

  // --- paragraph formatting (commands/paragraph.ts) ---
  "w:p": "structure.insertParagraph",
  "w:pPr": "paragraph.align",
  "w:jc": "paragraph.alignLeft",
  "w:ind": "paragraph.indent",
  "w:spacing": "paragraph.spacing",
  "w:pBdr": "paragraph.borders",
  "w:shd": "paragraph.shading",
  "w:pStyle": "paragraph.style",

  // --- lists / numbering (commands/list.ts) ---
  "w:numPr": "list.apply",
  "w:numId": "list.apply",
  "w:ilvl": "list.apply",
  "w:num": "list.apply",
  "w:abstractNum": "list.apply",
  "w:lvl": "list.apply",
  "w:numbering": "list.apply",

  // --- tables (commands/table.ts) ---
  "w:tbl": "table.insert",
  "w:tr": "table.addRow",
  "w:tc": "table.insert",
  "w:tblGrid": "table.insert",
  "w:gridCol": "table.insert",
  "w:tblPr": "table.borders",
  "w:tblBorders": "table.borders",
  "w:tcPr": "table.cellShading",
  "w:tcBorders": "table.borders",
  "w:tcW": "table.insert",
  "w:vAlign": "table.cellVAlign",
  "w:trPr": "table.rowHeight",
  "w:trHeight": "table.rowHeight",
  "w:tblHeader": "table.rowHeader",

  // --- images (commands/image.ts) ---
  "w:drawing": "image.insert",

  // --- styles (commands/style.ts, paragraph.style) ---
  "w:style": "style.add",
  "w:styles": "style.add",
  "w:basedOn": "style.add",
  "w:next": "style.add",
  "w:link": "style.add",
  "w:docDefaults": "style.ensureHeadings",
  "w:rPrDefault": "style.ensureHeadings",
  "w:pPrDefault": "style.ensureHeadings",

  // --- section / page setup (commands/section.ts) ---
  "w:sectPr": "section.break",
  "w:pgSz": "section.pageSize",
  "w:pgMar": "section.pageMargins",

  // --- header / footer (commands/header-footer.ts) ---
  "w:headerReference": "headerFooter.addHeader",
  "w:footerReference": "headerFooter.addFooter",
  "w:hdr": "headerFooter.addHeader",
  "w:ftr": "headerFooter.addFooter",

  // --- references (commands/references.ts) ---
  "w:bookmarkStart": "references.bookmark",
  "w:bookmarkEnd": "references.bookmark",
  "w:hyperlink": "references.hyperlink",
  "w:footnoteReference": "references.footnote",
  "w:endnoteReference": "references.endnote",
  "w:footnote": "references.footnote",
  "w:endnote": "references.endnote",
  "w:fldSimple": "references.field",
  "w:fldChar": "references.field",
  "w:instrText": "references.field",

  // --- review: comments + tracked changes (commands/review.ts) ---
  "w:comment": "review.addComment",
  "w:comments": "review.addComment",
  "w:commentRangeStart": "review.addComment",
  "w:commentRangeEnd": "review.addComment",
  "w:commentReference": "review.addComment",
  "w:ins": "review.acceptAll",
  "w:del": "review.acceptAll",
  "w:pPrChange": "review.acceptAll",
  "w:rPrChange": "review.acceptAll",

  // --- document properties (commands/docprops.ts) ---
  "ep:Application": "docprops.setApp",
  "ep:AppVersion": "docprops.setApp",
  "ep:Pages": "docprops.setApp",
  "ep:Words": "docprops.setApp",
  "ep:Characters": "docprops.setApp",
  "ep:CharactersWithSpaces": "docprops.setApp",
  "ep:Paragraphs": "docprops.setApp",
  "ep:Lines": "docprops.setApp",
  "ep:Template": "docprops.setApp",
  "ep:Company": "docprops.setApp",
  "ep:Manager": "docprops.setApp",

  // --- inline image / DrawingML picture chain produced by image.insert ---
  "wp:inline": "image.insert",
  "wp:anchor": "image.insert",
  "wp:docPr": "image.altText",
  "a:blip": "image.insert",
  "a:graphic": "image.insert",
  "a:graphicData": "image.insert",
  "pic:pic": "image.insert",
  "pic:blipFill": "image.insert",
  "pic:spPr": "image.insert",
  // Editable via resize / alt-text commands on an existing image.
  "wp:extent": "image.resize",
  "a:ext": "image.resize",
  "a:off": "image.resize",
  "a:xfrm": "image.resize",
  "pic:nvPicPr": "image.altText",
  "pic:cNvPr": "image.altText",
};

// Fold in the generic run/paragraph property commands (commands/properties.ts).
// `??=` so any explicit mapping above wins (e.g. w:spacing → paragraph.spacing).
for (const local of RUN_PROPERTY_ELEMENTS) EDIT_MAP[`w:${local}`] ??= `text.${local}`;
for (const local of PARA_PROPERTY_ELEMENTS) EDIT_MAP[`w:${local}`] ??= `paragraph.${local}`;
for (const { local, id } of TABLE_PROPERTY_BINDINGS) EDIT_MAP[`w:${local}`] ??= id;
for (const local of SECTION_PROPERTY_ELEMENTS) EDIT_MAP[`w:${local}`] ??= `section.${local}`;
for (const local of SETTINGS_ELEMENTS) EDIT_MAP[`w:${local}`] ??= `settings.${local}`;
for (const local of STYLE_PROPERTY_ELEMENTS) EDIT_MAP[`w:${local}`] ??= `style.${local}`;
for (const local of NUMBERING_PROPERTY_ELEMENTS) EDIT_MAP[`w:${local}`] ??= `numbering.${local}`;

/**
 * Elements the canvas renderer paints (structure + inline text features) even
 * when no command edits them directly. These are shown, not lost.
 */
const RENDER_SET = new Set<string>([
  "w:document",
  "w:body",
  "w:tab",
  "w:noBreakHyphen",
  "w:softHyphen",
  "w:sym",
  "w:cols",
  "w:lastRenderedPageBreak",
]);

/**
 * Ordered fallback rules by schema domain. Every universe element has a domain,
 * so these cover the entire remainder. Each records the round-trip test that
 * proves the library keeps the element intact across open→save.
 */
const PRESERVE_RULES: Record<string, string> = {
  wml: "Losslessly round-tripped as raw XML pass-through; not yet surfaced as an edit command. Proof: src/internal/wordprocessingml/round-trip.test.ts.",
  drawing:
    "DrawingML kept intact inside the raw <w:drawing> subtree. Proof: src/api/image.test.ts + round-trip suite.",
  math: "OMML (office math) preserved as a raw pass-through subtree. Proof: round-trip suite.",
  vml: "Legacy VML preserved as a raw pass-through subtree. Proof: round-trip suite.",
  docprops:
    "Document properties round-tripped; core/app values editable via setCoreProperties/setAppProperties. Proof: src/api/core-properties.test.ts.",
};

/**
 * Domains whose elements live inline in `document.xml` as raw XML subtrees
 * (DrawingML, OMML math, legacy VML) and so are editable through the universal
 * raw-XML inspector (`raw.setChildVal` / `raw.setAttribute`) even without a
 * bespoke command. The mutation flushes via the document AST on save.
 */
const RAW_EDITABLE_DOMAINS = new Set(["drawing", "math", "vml"]);
const RAW_EDIT_COMMAND = "raw.setChildVal";
const RAW_EDIT_NOTE =
  "Editable via the document raw-XML inspector (raw.setChildVal / raw.setAttribute); lives inline in document.xml and flushes through the document AST.";

/**
 * Every remaining OOXML element lives in some XML part (document / styles /
 * numbering / settings / fontTable / comments / foot- & endnotes / headers /
 * footers / webSettings / docProps). The part-level raw-XML inspector opens any
 * part, exposes its full element tree, and edits any element's attributes or
 * children via the `rawpart.*` commands, re-serializing that part on save — so
 * every element is UI-editable through this universal escape hatch.
 */
const RAW_PART_DOMAINS = new Set(["wml", "docprops"]);
const RAW_PART_COMMAND = "rawpart.setChildVal";
const RAW_PART_NOTE =
  "Editable via the part-level raw-XML inspector (rawpart.setChildVal / rawpart.setAttribute), which opens the owning XML part and re-serializes it on save.";

/** Classify one qualified element name into a capability. */
export function classify(element: string, domain: string): Capability {
  const command = EDIT_MAP[element];
  if (command) return { element, domain, disposition: "edit", command };
  if (RENDER_SET.has(element)) return { element, domain, disposition: "render" };
  if (RAW_EDITABLE_DOMAINS.has(domain)) {
    return {
      element,
      domain,
      disposition: "edit",
      command: RAW_EDIT_COMMAND,
      notes: RAW_EDIT_NOTE,
    };
  }
  if (RAW_PART_DOMAINS.has(domain)) {
    return {
      element,
      domain,
      disposition: "edit",
      command: RAW_PART_COMMAND,
      notes: RAW_PART_NOTE,
    };
  }
  const preserveNote = PRESERVE_RULES[domain];
  if (preserveNote) return { element, domain, disposition: "preserve", notes: preserveNote };
  // No rule matched — surfaced by the coverage test as a gap to close.
  return { element, domain, disposition: "preserve", notes: "UNCLASSIFIED — no domain rule." };
}

/** The full ledger: every universe element classified. */
export const LEDGER: readonly Capability[] = ELEMENTS.map((e) => classify(e.element, e.domain));

export interface CoverageSummary {
  readonly total: number;
  readonly edit: number;
  readonly render: number;
  readonly preserve: number;
  readonly unclassified: number;
  readonly byDomain: Record<string, { edit: number; render: number; preserve: number }>;
  /** Elements whose `edit` command id is not present in the command registry. */
  readonly danglingCommands: string[];
}

/** Aggregate the ledger for the report and the enforcement test. */
export function coverageSummary(): CoverageSummary {
  const byDomain: Record<string, { edit: number; render: number; preserve: number }> = {};
  let edit = 0;
  let render = 0;
  let preserve = 0;
  let unclassified = 0;
  const dangling: string[] = [];
  for (const cap of LEDGER) {
    const d = (byDomain[cap.domain] ??= { edit: 0, render: 0, preserve: 0 });
    if (cap.disposition === "edit") {
      edit++;
      d.edit++;
      if (cap.command && !COMMAND_IDS.has(cap.command))
        dangling.push(`${cap.element} -> ${cap.command}`);
    } else if (cap.disposition === "render") {
      render++;
      d.render++;
    } else {
      preserve++;
      d.preserve++;
      if (cap.notes?.startsWith("UNCLASSIFIED")) unclassified++;
    }
  }
  return {
    total: LEDGER.length,
    edit,
    render,
    preserve,
    unclassified,
    byDomain,
    danglingCommands: dangling,
  };
}
