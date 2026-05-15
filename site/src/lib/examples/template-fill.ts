// Open an existing .docx as a template, run cross-part placeholder
// substitution, write it back. The interesting bit: replaceTextEverywhere
// also walks headers / footers / footnotes / textboxes — not just the
// body — so {{name}} in a header swaps out too.

import { openDocx, replaceTextEverywhere, toUint8Array } from "@word-kit/core";

declare const templateBytes: Uint8Array;
declare const values: Record<string, string>;

const doc = openDocx(templateBytes);

const replaced: number = replaceTextEverywhere(
  doc,
  /\{\{(\w+)\}\}/g,
  (m) => values[m.captures[0] ?? ""] ?? "",
);

const out: Uint8Array = toUint8Array(doc);
void replaced;
void out;
