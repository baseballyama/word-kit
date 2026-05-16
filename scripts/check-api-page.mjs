#!/usr/bin/env node
// Verify that site/src/lib/api-groups.ts lists every named export from
// @word-kit/core (and only those — no ghost names left behind from
// renames). The API listing is hand-curated so contributors can group
// and order it, but that means it silently goes stale when an export is
// added or removed; this script is the safety net.
//
// Both the `/api` page and `/llms-full.txt` consume `api-groups.ts`, so
// this single check covers both surfaces.
//
// Run via `pnpm check:api-page`. CI runs it after the build job, since
// it imports the built `packages/core/dist/index.mjs` artefact.
//
// Exit codes: 0 = in sync. 1 = drift found (prints the diff and exits).

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");
const CORE_DIST = resolve(ROOT, "packages/core/dist/index.mjs");
const API_GROUPS = resolve(ROOT, "site/src/lib/api-groups.ts");

// --- Read the live exports off the built core bundle. ----------------

const coreExports = await import(CORE_DIST);
const exportedNames = Object.keys(coreExports)
  .filter((n) => n !== "default")
  .filter((n) => !n.startsWith("_")); // private convention, if any

// --- Read what the shared api-groups module advertises. --------------

const groupsSrc = readFileSync(API_GROUPS, "utf-8");

// `name: "createDocx"` entries inside the entries arrays. We match the
// literal `name:` field rather than arbitrary string occurrences, so
// nothing in the file's comments accidentally counts.
const NAME_RE = /\bname:\s*["']([A-Za-z_$][A-Za-z0-9_$]*)["']/g;

const advertised = new Set();
for (const match of groupsSrc.matchAll(NAME_RE)) {
  advertised.add(match[1]);
}

// `previewToDOM` lives in `@word-kit/preview`, not `@word-kit/core`. The
// API page lists it under the "Browser preview" group, so we tolerate it
// being on the page even though it isn't in the core bundle.
const TOLERATED_NON_CORE = new Set(["previewToDOM"]);

// --- Diff and report. ------------------------------------------------

const exported = new Set(exportedNames);
const missingFromPage = [...exported].filter((n) => !advertised.has(n)).toSorted();
const stale = [...advertised]
  .filter((n) => !exported.has(n) && !TOLERATED_NON_CORE.has(n))
  .toSorted();

let failed = false;

if (missingFromPage.length) {
  failed = true;
  console.error(
    `[check:api-page] ${missingFromPage.length} export(s) missing from site/src/lib/api-groups.ts:\n` +
      missingFromPage.map((n) => `  + ${n}`).join("\n") +
      "\n→ add them to the appropriate group (or open a new group).",
  );
}

if (stale.length) {
  failed = true;
  console.error(
    `[check:api-page] ${stale.length} name(s) in api-groups.ts are no longer exported from @word-kit/core:\n` +
      stale.map((n) => `  - ${n}`).join("\n") +
      "\n→ remove them from site/src/lib/api-groups.ts (or, if they moved to another package, add the package to TOLERATED_NON_CORE in this script).",
  );
}

if (failed) {
  process.exit(1);
}

console.log(
  `[check:api-page] OK — ${exported.size} core exports, ${advertised.size} entries in api-groups.ts (incl. ${TOLERATED_NON_CORE.size} tolerated from sibling packages).`,
);
