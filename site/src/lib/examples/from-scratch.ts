// Build a .docx from scratch. This file is imported as ?raw into the
// docs site so the snippet shown to readers is exactly what svelte-check
// type-checked — if an API rename breaks it, the docs build fails.

import {
  addBulletList,
  addTable,
  appendHeading,
  appendParagraph,
  createDocx,
  PAGE_SIZE_A4,
  setPageSize,
  toUint8Array,
} from "@word-kit/core";

const doc = createDocx({ paragraphs: [] });
setPageSize(doc, PAGE_SIZE_A4);

appendHeading(doc, "Quarterly review", 1);
appendParagraph(doc, "Highlights from Q3.");
addBulletList(doc, ["Shipped preview", "Round-trip stable", "512 tests green"]);

addTable(doc, [
  ["Metric", "Q2", "Q3"],
  ["MRR", "$120k", "$148k"],
  ["Churn", "3.1%", "2.4%"],
]);

const bytes: Uint8Array = toUint8Array(doc);
void bytes;
