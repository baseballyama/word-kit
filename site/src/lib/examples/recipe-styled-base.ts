// PowerPoint-style "designed base": open a hand-designed .docx, lift
// only its style table into your authoring graph, then keep appending
// body content. The visual identity (fonts, colours, custom paragraph
// styles) carries over without any placeholder-substitution dance.

import {
  appendHeading,
  appendParagraph,
  createDocx,
  findStyleIdByName,
  mergeStylesFromTemplate,
  setParagraphStyle,
  toUint8Array,
} from "@word-kit/core";

declare const templateBytes: Uint8Array;

const doc = createDocx({ paragraphs: [] });

// Merge styles from the hand-designed template into the empty doc.
// `overwrite: false` (default) means built-in heading styles in the
// authoring graph win if both define them; pass `overwrite: true` to
// take the template's flavour.
const merged: number = mergeStylesFromTemplate(doc, templateBytes);
console.log(`merged ${merged} styles from template`);

appendHeading(doc, "Quarterly review", 1);

// Apply a custom paragraph style by *name* (as it appears in Word's UI)
// rather than the raw `w:styleId`. Returns the underlying styleId so
// downstream code can re-use it.
const calloutId: string | undefined = findStyleIdByName(doc, "Callout");
const callout = appendParagraph(doc, "Heads-up: this is the executive summary.");
if (calloutId) setParagraphStyle(callout, calloutId);

const bytes: Uint8Array = toUint8Array(doc);
void bytes;
