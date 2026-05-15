#!/usr/bin/env node
// Verify that site/src/routes/api/+page.svelte lists every named export
// from @word-kit/core (and only those — no ghost names left behind from
// renames). The site's API page is hand-curated so contributors can group
// and order the listing, but that means it silently goes stale when an
// export is added or removed; this script is the safety net.
//
// Run via `pnpm check:api-page`. CI runs it after the build job, since
// it imports the built `packages/core/dist/index.mjs` artefact.
//
// Exit codes: 0 = in sync. 1 = drift found (prints the diff and exits).

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, "..");
const CORE_DIST = resolve(ROOT, "packages/core/dist/index.mjs");
const API_PAGE = resolve(ROOT, "site/src/routes/api/+page.svelte");

// --- Read the live exports off the built core bundle. ----------------

const coreExports = await import(CORE_DIST);
const exportedNames = Object.keys(coreExports)
  .filter((n) => n !== "default")
  .filter((n) => !n.startsWith("_")); // private convention, if any

// --- Read what the API page advertises. ------------------------------

const pageSrc = readFileSync(API_PAGE, "utf-8");

// `name: "createDocx"` entries inside the `groups` array. We're matching
// the literal `name:` field, not arbitrary string occurrences, to avoid
// catching prose mentions inside template descriptions.
const NAME_RE = /\bname:\s*["']([A-Za-z_$][A-Za-z0-9_$]*)["']/g;

const advertised = new Set();
for (const match of pageSrc.matchAll(NAME_RE)) {
  advertised.add(match[1]);
}

// `previewToDOM` lives in `@word-kit/preview`, not `@word-kit/core`. The
// API page lists it under the "Browser preview" group, so we tolerate it
// being on the page even though it isn't in the core bundle.
const TOLERATED_NON_CORE = new Set(["previewToDOM"]);

// --- Diff and report. ------------------------------------------------

const exported = new Set(exportedNames);
const missingFromPage = [...exported].filter((n) => !advertised.has(n)).sort();
const stale = [...advertised].filter((n) => !exported.has(n) && !TOLERATED_NON_CORE.has(n)).sort();

let failed = false;

if (missingFromPage.length) {
  failed = true;
  console.error(
    `[check:api-page] ${missingFromPage.length} export(s) missing from site/src/routes/api/+page.svelte:\n` +
      missingFromPage.map((n) => `  + ${n}`).join("\n") +
      "\n→ add them to the appropriate group in the page (or move them to a new group).",
  );
}

if (stale.length) {
  failed = true;
  console.error(
    `[check:api-page] ${stale.length} name(s) on the API page are no longer exported from @word-kit/core:\n` +
      stale.map((n) => `  - ${n}`).join("\n") +
      "\n→ remove them from site/src/routes/api/+page.svelte (or, if they moved to another package, add the package to TOLERATED_NON_CORE in this script).",
  );
}

if (failed) {
  process.exit(1);
}

console.log(
  `[check:api-page] OK — ${exported.size} core exports, ${advertised.size} entries on the API page (incl. ${TOLERATED_NON_CORE.size} tolerated from sibling packages).`,
);
