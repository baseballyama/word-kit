// Extract the element universe from the ECMA-376 XSD schemas so the editor's
// capability ledger can be checked for completeness against it.
//
// Source of truth: references/python-docx/ref/xsd/*.xsd (git submodule, never
// published). We snapshot the extracted list into a committed JSON file so the
// coverage test runs even when the submodule is not checked out (CI).
//
// Scope: docx (WordprocessingML). We take every element declared in wml.xsd,
// plus the wordprocessing-relevant DrawingML / VML / math elements a .docx can
// actually contain. Each element is tagged with the namespace prefix and the
// nearest enclosing named complexType (for grouping in the ledger).
//
// Run: node scripts/extract-ooxml-elements.mjs

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(here, "..");
const xsdDir = join(repoRoot, "references/python-docx/ref/xsd");
const outFile = join(repoRoot, "packages/editor/src/capability/element-universe.json");

// Which schema file maps to which namespace prefix + feature domain. Only the
// parts a WordprocessingML document can embed are included.
const SCHEMAS = [
  { file: "wml.xsd", prefix: "w", domain: "wml" },
  { file: "dml-wordprocessingDrawing.xsd", prefix: "wp", domain: "drawing" },
  { file: "dml-picture.xsd", prefix: "pic", domain: "drawing" },
  { file: "dml-main.xsd", prefix: "a", domain: "drawing" },
  { file: "shared-math.xsd", prefix: "m", domain: "math" },
  { file: "vml-main.xsd", prefix: "v", domain: "vml" },
  { file: "vml-officeDrawing.xsd", prefix: "o", domain: "vml" },
  { file: "vml-wordprocessingDrawing.xsd", prefix: "wvml", domain: "vml" },
  { file: "shared-documentPropertiesCore.xsd", prefix: "cp", domain: "docprops" },
  { file: "shared-documentPropertiesExtended.xsd", prefix: "ep", domain: "docprops" },
  { file: "shared-documentPropertiesCustom.xsd", prefix: "custprops", domain: "docprops" },
];

const ELEMENT_RE = /<xsd:element\s+name="([A-Za-z0-9_]+)"(?:\s+type="([^"]+)")?/g;
const COMPLEXTYPE_OPEN_RE = /<xsd:complexType\s+name="([A-Za-z0-9_]+)"/;
const COMPLEXTYPE_CLOSE_RE = /<\/xsd:complexType>/;

/**
 * Walk a schema line by line, tracking the current named complexType so each
 * element can be attributed to the type that declares it. Elements declared at
 * the top level (outside any complexType) get definedIn = null.
 */
function extractFromSchema(text, prefix, domain) {
  const lines = text.split("\n");
  const typeStack = [];
  const found = new Map(); // key = `${prefix}:${name}` -> record

  for (const line of lines) {
    const open = line.match(COMPLEXTYPE_OPEN_RE);
    if (open) typeStack.push(open[1]);

    let m;
    ELEMENT_RE.lastIndex = 0;
    while ((m = ELEMENT_RE.exec(line)) !== null) {
      const name = m[1];
      const type = m[2] ?? null;
      const key = `${prefix}:${name}`;
      const definedIn = typeStack.length > 0 ? typeStack[typeStack.length - 1] : null;
      const existing = found.get(key);
      if (existing) {
        if (definedIn && !existing.definedIn.includes(definedIn)) {
          existing.definedIn.push(definedIn);
        }
        if (type && !existing.types.includes(type)) existing.types.push(type);
      } else {
        found.set(key, {
          element: key,
          name,
          prefix,
          domain,
          definedIn: definedIn ? [definedIn] : [],
          types: type ? [type] : [],
        });
      }
    }

    if (COMPLEXTYPE_CLOSE_RE.test(line)) typeStack.pop();
  }
  return [...found.values()];
}

const all = new Map();
let missingFiles = 0;
for (const { file, prefix, domain } of SCHEMAS) {
  const path = join(xsdDir, file);
  let text;
  try {
    text = readFileSync(path, "utf8");
  } catch {
    // shared-documentPropertiesCore.xsd is named differently in some mirrors;
    // skip missing optional schemas rather than fail the whole extraction.
    missingFiles++;
    console.warn(`[extract] skip missing schema: ${file}`);
    continue;
  }
  for (const rec of extractFromSchema(text, prefix, domain)) {
    // wml re-declares a handful of names; keep the first (wml wins for w:).
    if (!all.has(rec.element)) all.set(rec.element, rec);
  }
}

const elements = [...all.values()].toSorted((a, b) => a.element.localeCompare(b.element));
const byDomain = {};
for (const e of elements) byDomain[e.domain] = (byDomain[e.domain] ?? 0) + 1;

const snapshot = {
  // Regenerate with: node scripts/extract-ooxml-elements.mjs
  generatedFrom: "ECMA-376 XSD (references/python-docx/ref/xsd)",
  schemaFiles: SCHEMAS.map((s) => s.file),
  totalElements: elements.length,
  byDomain,
  elements,
};

writeFileSync(outFile, JSON.stringify(snapshot, null, 2) + "\n");
console.log(
  `[extract] ${elements.length} elements from ${SCHEMAS.length - missingFiles} schemas -> ${outFile}`,
);
console.log("[extract] by domain:", byDomain);
