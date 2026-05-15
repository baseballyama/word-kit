import type { XmlElement } from "@word-kit/ooxml-xml";
import { WML_NS } from "./namespaces.js";
import type { WmlBlock, WmlDocument, WmlInline, WmlParagraph, WmlRunPiece } from "./types.js";

/**
 * A successful match against the flattened text of a paragraph.
 *
 * The match's `start` and `end` offsets refer to the **paragraph's flat
 * text** — that is, the concatenation of every `<w:t>` value in document
 * order, ignoring tab/break/drawing pieces and unrecognized inlines.
 */
export interface TextMatch {
  readonly paragraph: WmlParagraph;
  readonly text: string;
  readonly start: number;
  readonly end: number;
  readonly captures: readonly string[];
}

/**
 * Find all matches of `query` in every paragraph of the document.
 *
 * - When `query` is a string, performs literal, case-sensitive search.
 * - When `query` is a `RegExp`, the search is global regardless of whether
 *   the source pattern had the `g` flag.
 */
export function findText(doc: WmlDocument, query: string | RegExp): TextMatch[] {
  const matches: TextMatch[] = [];
  for (const block of doc.body.blocks) {
    if (block.kind !== "paragraph") continue;
    matches.push(...findInParagraph(block, query));
  }
  return matches;
}

/**
 * Replace every occurrence of `query` in the document with `replacement`.
 *
 * The replacement may either be a plain string or a function that receives
 * each match and returns the replacement text. Returns the number of
 * matches replaced.
 *
 * Cross-run matches are supported: if `{{name}}` is split across multiple
 * `<w:r>` elements, the replacement is inserted into the first run's text
 * piece (inheriting its `rPr`), and any text consumed from following runs
 * is cleared so the visible output reads the replacement followed by the
 * post-match suffix.
 */
export function replaceText(
  doc: WmlDocument,
  query: string | RegExp,
  replacement: string | ((match: TextMatch) => string),
): number {
  let total = 0;
  for (const block of doc.body.blocks) {
    if (block.kind !== "paragraph") continue;
    total += replaceInParagraph(block, query, replacement);
  }
  return total;
}

/** Same as {@link findText} but limited to a single paragraph. */
export function findInParagraph(p: WmlParagraph, query: string | RegExp): TextMatch[] {
  const { flat } = flattenParagraph(p);
  return runQuery(flat, query, p);
}

/** Same as {@link replaceText} but limited to a single paragraph. */
export function replaceInParagraph(
  p: WmlParagraph,
  query: string | RegExp,
  replacement: string | ((match: TextMatch) => string),
): number {
  const { flat, refs } = flattenParagraph(p);
  const matches = runQuery(flat, query, p);
  if (matches.length === 0) return 0;

  type Edit = { start: number; end: number; replacement: string };
  const edits = new Map<number, Edit[]>();
  const push = (refIdx: number, edit: Edit): void => {
    const arr = edits.get(refIdx);
    if (arr) arr.push(edit);
    else edits.set(refIdx, [edit]);
  };

  for (const match of matches) {
    const startLoc = locate(refs, match.start);
    const endLoc = locate(refs, match.end);
    if (!startLoc || !endLoc) continue;
    const repl = typeof replacement === "function" ? replacement(match) : replacement;

    if (startLoc.refIdx === endLoc.refIdx) {
      push(startLoc.refIdx, {
        start: startLoc.offset,
        end: endLoc.offset,
        replacement: repl,
      });
      continue;
    }
    const startRef = refs[startLoc.refIdx];
    const endRef = refs[endLoc.refIdx];
    if (!startRef || !endRef) continue;
    push(startLoc.refIdx, {
      start: startLoc.offset,
      end: startRef.text.length,
      replacement: repl,
    });
    for (let i = startLoc.refIdx + 1; i < endLoc.refIdx; i++) {
      const ref = refs[i];
      if (!ref) continue;
      push(i, { start: 0, end: ref.text.length, replacement: "" });
    }
    push(endLoc.refIdx, { start: 0, end: endLoc.offset, replacement: "" });
  }

  for (const [refIdx, refEdits] of edits) {
    refEdits.sort((a, b) => b.start - a.start);
    const ref = refs[refIdx];
    if (!ref) continue;
    let value = ref.text;
    for (const e of refEdits) {
      value = value.slice(0, e.start) + e.replacement + value.slice(e.end);
    }
    const piece = getTextPiece(p, ref);
    if (piece) {
      piece.value = value;
      // Preserve leading/trailing whitespace by forcing xml:space="preserve"
      // whenever the resulting text contains leading or trailing whitespace.
      if (/^\s|\s$/.test(value)) piece.preserveSpace = true;
    }
  }
  return matches.length;
}

interface TextRef {
  runChildIdx: number;
  pieceIdx: number;
  /** Original text snapshot at flatten time. */
  text: string;
}

function flattenParagraph(p: WmlParagraph): { flat: string; refs: TextRef[] } {
  const refs: TextRef[] = [];
  let flat = "";
  for (let r = 0; r < p.children.length; r++) {
    const child = p.children[r];
    if (!child || child.kind !== "run") continue;
    for (let pi = 0; pi < child.pieces.length; pi++) {
      const piece = child.pieces[pi];
      if (!piece) continue;
      if (piece.kind === "text") {
        refs.push({ runChildIdx: r, pieceIdx: pi, text: piece.value });
        flat += piece.value;
      }
    }
  }
  return { flat, refs };
}

function runQuery(flat: string, query: string | RegExp, paragraph: WmlParagraph): TextMatch[] {
  if (typeof query === "string") {
    if (query.length === 0) return [];
    const out: TextMatch[] = [];
    let idx = 0;
    while ((idx = flat.indexOf(query, idx)) >= 0) {
      out.push({
        paragraph,
        text: query,
        start: idx,
        end: idx + query.length,
        captures: [],
      });
      idx += query.length;
    }
    return out;
  }
  const re = query.global ? query : new RegExp(query.source, `${query.flags}g`);
  re.lastIndex = 0;
  const out: TextMatch[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(flat)) !== null) {
    if (m[0].length === 0) {
      re.lastIndex++;
      continue;
    }
    out.push({
      paragraph,
      text: m[0],
      start: m.index,
      end: m.index + m[0].length,
      captures: m.slice(1).map((s) => s ?? ""),
    });
  }
  return out;
}

function locate(
  refs: readonly TextRef[],
  absOffset: number,
): { refIdx: number; offset: number } | undefined {
  let acc = 0;
  for (let i = 0; i < refs.length; i++) {
    const ref = refs[i];
    if (!ref) continue;
    if (absOffset <= acc + ref.text.length) {
      return { refIdx: i, offset: absOffset - acc };
    }
    acc += ref.text.length;
  }
  return undefined;
}

function getTextPiece(p: WmlParagraph, ref: TextRef): (WmlRunPiece & { kind: "text" }) | undefined {
  const run = p.children[ref.runChildIdx];
  if (!run || run.kind !== "run") return undefined;
  const piece = run.pieces[ref.pieceIdx];
  if (!piece || piece.kind !== "text") return undefined;
  return piece;
}

/**
 * Return the concatenated visible text content of a paragraph.
 *
 * Includes text inside structured runs (`<w:t>`), nested hyperlinks
 * (`<w:hyperlink>...<w:r><w:t>...`), structured-document tags
 * (`<w:sdt>`), tracked-change insertions (`<w:ins>`) and deletions
 * (`<w:del>`). `<w:tab/>` becomes `\t` and `<w:br/>` becomes `\n` so
 * the result round-trips with `appendParagraph` / `appendTextRun`.
 * Drawing / picture / page-break content is still ignored.
 */
export function paragraphText(p: WmlParagraph): string {
  let acc = "";
  for (const child of p.children) {
    acc += visibleTextOfInline(child);
  }
  return acc;
}

function visibleTextOfInline(inline: WmlInline): string {
  if (inline.kind === "run") {
    let acc = "";
    for (const piece of inline.pieces) {
      if (piece.kind === "text" || piece.kind === "delText") acc += piece.value;
      else if (piece.kind === "tab") acc += "\t";
      else if (piece.kind === "break" && piece.breakType !== "page") acc += "\n";
    }
    return acc;
  }
  return visibleTextOfElement(inline.node);
}

function visibleTextOfElement(el: XmlElement): string {
  let acc = "";
  for (const child of el.children) {
    if (child.kind === "element") {
      if (
        child.name.uri === WML_NS &&
        (child.name.local === "t" || child.name.local === "delText")
      ) {
        for (const tc of child.children) {
          if (tc.kind === "text") acc += tc.value;
          else if (tc.kind === "cdata") acc += tc.value;
        }
      } else if (child.name.uri === WML_NS && child.name.local === "tab") {
        acc += "\t";
      } else if (child.name.uri === WML_NS && child.name.local === "br") {
        const typeAttr = child.attrs.find((a) => a.name.uri === WML_NS && a.name.local === "type");
        if (typeAttr?.value !== "page") acc += "\n";
      } else {
        acc += visibleTextOfElement(child);
      }
    }
  }
  return acc;
}

/** Like {@link paragraphText} but for an entire document. */
export function documentText(doc: WmlDocument, separator = "\n"): string {
  const parts: string[] = [];
  for (const block of doc.body.blocks) {
    parts.push(visibleTextOfBlock(block));
  }
  return parts.join(separator);
}

function visibleTextOfBlock(block: WmlBlock): string {
  if (block.kind === "paragraph") return paragraphText(block);
  if (block.kind === "table") {
    const parts: string[] = [];
    for (const row of block.rows) {
      for (const cell of row.cells) {
        for (const p of cell.paragraphs) parts.push(paragraphText(p));
      }
    }
    return parts.join("\n");
  }
  return visibleTextOfElement(block.node);
}
