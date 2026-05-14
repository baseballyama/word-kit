#!/usr/bin/env node
// Tree-shake budget check. Bundles a tiny entry that imports a small slice
// of @word-kit/core and verifies that:
//
//   1. The minified bundle size stays under MIN_BUDGET. If we ever add code
//      that the bundler can't drop, this catches it.
//   2. The minimal bundle is strictly smaller than a bundle that imports the
//      whole package. A regression where tree-shaking is silently broken
//      (eg. by re-introducing a class) would make these two equal.
//   3. No "forbidden" string literals from unused features leak into the
//      minimal bundle. These are stable markers from feature areas the
//      minimal entry doesn't touch (numbering, images, comments, footnotes,
//      etc.) — if they show up after minify, the corresponding code branch
//      survived tree-shaking when it should have been dropped.
//
// Run via `node scripts/check-tree-shake.mjs` after `pnpm build`. CI runs
// this on every push to main and on every PR.

import { build } from "esbuild";
import { fileURLToPath } from "node:url";
import { dirname, join, resolve } from "node:path";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");
// Resolve workspace package imports as a sibling of @word-kit/core would.
const RESOLVE_DIR = join(ROOT, "packages", "core");

async function bundleSize(entry) {
  const out = await build({
    stdin: {
      contents: entry,
      loader: "js",
      resolveDir: RESOLVE_DIR,
      sourcefile: "treeshake-entry.mjs",
    },
    bundle: true,
    minify: true,
    format: "esm",
    platform: "neutral",
    target: "es2022",
    write: false,
    mainFields: ["module", "main"],
    conditions: ["import"],
    treeShaking: true,
    absWorkingDir: ROOT,
  });
  return out.outputFiles[0].text;
}

const minimalEntry = `
import { createDocx, appendParagraph, toUint8Array } from "@word-kit/core";
const d = createDocx();
appendParagraph(d, "hello");
globalThis.__bytes = toUint8Array(d);
`;

const fullEntry = `
import * as everything from "@word-kit/core";
globalThis.__exports = Object.keys(everything);
`;

const minimal = await bundleSize(minimalEntry);
const full = await bundleSize(fullEntry);

const MIN_BUDGET = 50_000; // 50 KB after minify
console.log(`minimal bundle: ${minimal.length} bytes`);
console.log(`full bundle:    ${full.length} bytes`);

let failed = false;

if (minimal.length > MIN_BUDGET) {
  console.error(`FAIL: minimal bundle exceeded budget (${minimal.length} > ${MIN_BUDGET}).`);
  failed = true;
}

if (minimal.length >= full.length) {
  console.error(
    `FAIL: minimal bundle (${minimal.length}) is not smaller than full bundle (${full.length}). Tree-shake regressed.`,
  );
  failed = true;
}

// Markers from feature-specific function bodies that the minimal entry
// should not reach. These are stable string literals; if any of them
// shows up after tree-shake, the corresponding feature survived when it
// shouldn't have.
//
// We intentionally avoid markers that come from top-level `var X = [...].join("")`
// declarations: esbuild keeps those because it can't prove the join() call
// is pure. They aren't a real bloat signal — the unused branches that
// reference them are dropped — but they would trigger false positives here.
const FORBIDDEN_MARKERS = [
  // Validator.
  "validatePackage",
  // Image (addImage / addImageRun / insertImageInto).
  "<w:drawing",
  '"image/png"',
  '"image/jpeg"',
  // Hyperlinks (addHyperlink).
  "<w:hyperlink",
  // Headers / footers (addHeader / addFooter / removeAllHeaders / …).
  "w:headerReference",
  "w:footerReference",
  // Footnote / endnote *body* code (addFootnote / addEndnote / removeAllFootnotes).
  "footnoteReference",
  "endnoteReference",
  // Revisions (acceptAllRevisions / rejectAllRevisions).
  "w:ins",
  // Numbering (addBulletList / addNumberedList / applyListToParagraph).
  "<w:numPr",
  "abstractNumId",
];

for (const marker of FORBIDDEN_MARKERS) {
  if (minimal.includes(marker)) {
    console.error(
      `FAIL: minimal bundle contains forbidden marker ${JSON.stringify(marker)} — tree-shake leak.`,
    );
    failed = true;
  }
}

if (failed) {
  process.exit(1);
}

console.log("tree-shake check passed.");
