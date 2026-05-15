import { existsSync, readdirSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { openDocx, paragraphs, text, toUint8Array } from "./docx.js";

const REPO_ROOT = resolve(import.meta.dirname, "../../..");
const FIXTURE_DIR = resolve(REPO_ROOT, "references/mammoth.js/test/test-data");

// strict-format.docx uses the ISO/IEC 29500 Strict namespaces which we do
// not yet support; the rest are transitional and should all round-trip.
const SKIP = new Set(["strict-format.docx"]);

describe("mammoth.js fixture corpus: open + round-trip parity", () => {
  if (!existsSync(FIXTURE_DIR)) {
    it.skip("references/mammoth.js submodule not initialised — run `git submodule update --init --recursive`", () => {
      // intentionally empty
    });
    return;
  }

  const all = readdirSync(FIXTURE_DIR)
    .filter((n) => n.endsWith(".docx"))
    .filter((n) => !SKIP.has(n));

  for (const name of all) {
    it(`opens and re-saves ${name} losslessly at the paragraph level`, () => {
      const bytes = readFileSync(resolve(FIXTURE_DIR, name));
      const doc = openDocx(bytes);
      const paragraphsBefore = paragraphs(doc).length;
      const textBefore = text(doc);
      const reopened = openDocx(toUint8Array(doc));
      expect(paragraphs(reopened).length).toBe(paragraphsBefore);
      expect(text(reopened)).toBe(textBefore);
    });
  }
});
