// Mail-merge across every story in the package — body, headers,
// footers, footnotes/endnotes, comments, textboxes. replaceTextEverywhere
// follows the same regex/callback contract as replaceText but walks the
// whole document, so a {{name}} hidden in a header is filled the same
// way one in the body is.

import { openDocx, replaceTextEverywhere, toUint8Array } from "@word-kit/core";

declare const templateBytes: Uint8Array;
declare const values: Record<string, string>;

const doc = openDocx(templateBytes);

// Regex captures land in `m.captures`; everything else (`m.match`,
// `m.index`) mirrors RegExp.exec semantics. Run-spanning matches are
// joined before the regex sees them, so `{{name}}` split across
// multiple runs still substitutes cleanly.
const replaced: number = replaceTextEverywhere(
  doc,
  /\{\{(\w+)\}\}/g,
  (m) => values[m.captures[0] ?? ""] ?? "",
);

console.log(`replaced ${replaced} placeholders`);

const filled: Uint8Array = toUint8Array(doc);
void filled;
