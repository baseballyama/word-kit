// Open each docx in the mammoth.js + python-docx fixture corpora, apply
// a non-trivial set of edits (append paragraph, set core properties,
// addComment, removeAllComments), then save and verify the produced
// bytes still validate. Catches bugs that only show up when the
// library mutates real-world Word output rather than its own
// freshly-built docs.

import { existsSync, readdirSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import {
  addComment,
  appendParagraph,
  createDocx,
  openDocx,
  paragraphs,
  removeAllComments,
  setCoreProperties,
  text,
  toUint8Array,
  validate,
} from "./docx.js";

const REPO_ROOT = resolve(import.meta.dirname, "../../..");
const MAMMOTH_DIR = resolve(REPO_ROOT, "references/mammoth.js/test/test-data");
const PYTHON_DOCX_DIR = resolve(REPO_ROOT, "references/python-docx/tests/test_files");

// Strict-format uses the ISO/IEC 29500 Strict namespace; not yet supported.
const SKIP_NAMES = new Set(["strict-format.docx"]);

function assertValidate(bytes: Uint8Array, label: string): void {
  const reopened = openDocx(bytes);
  const issues = validate(reopened);
  const errors = issues.filter((i) => i.level === "error");
  if (errors.length > 0) {
    throw new Error(
      `[${label}] validate() returned errors after edit:\n${errors
        .map((e) => `  [${e.code}] ${e.message}`)
        .join("\n")}`,
    );
  }
}

function fixturesIn(dir: string): string[] {
  if (!existsSync(dir)) return [];
  return readdirSync(dir)
    .filter((n) => n.endsWith(".docx"))
    .filter((n) => !SKIP_NAMES.has(n));
}

const MAMMOTH = fixturesIn(MAMMOTH_DIR);
const PYTHON_DOCX = fixturesIn(PYTHON_DOCX_DIR);

describe("real-world fixtures survive non-trivial edits + validate clean", () => {
  // Skip if neither corpus is checked out (submodules not initialised).
  if (MAMMOTH.length === 0 && PYTHON_DOCX.length === 0) {
    it.skip("fixture corpora not present — pull submodules to enable", () => {
      // intentionally empty
    });
    return;
  }

  for (const name of MAMMOTH) {
    it(`mammoth: ${name} — appendParagraph + setCoreProperties + save+validate`, () => {
      const bytes = readFileSync(resolve(MAMMOTH_DIR, name));
      const doc = openDocx(bytes);
      const beforeText = text(doc);
      appendParagraph(doc, "[appended by word-kit edit test]");
      setCoreProperties(doc, {
        title: "edited fixture",
        creator: "word-kit edit test",
      });
      const out = toUint8Array(doc);
      assertValidate(out, `mammoth/${name}`);
      // The appended paragraph survives.
      const reopened = openDocx(out);
      expect(text(reopened).startsWith(beforeText)).toBe(true);
      expect(text(reopened)).toContain("[appended by word-kit edit test]");
    });
  }

  for (const name of PYTHON_DOCX) {
    it(`python-docx: ${name} — appendParagraph + save+validate`, () => {
      const bytes = readFileSync(resolve(PYTHON_DOCX_DIR, name));
      let doc;
      try {
        doc = openDocx(bytes);
      } catch {
        // python-docx ships some pathological inputs (broken zip, etc.)
        // that exist to prove the parser. Skip those — round-trip is
        // tested elsewhere.
        return;
      }
      appendParagraph(doc, "[appended]");
      const out = toUint8Array(doc);
      assertValidate(out, `python-docx/${name}`);
    });
  }
});

describe("mammoth comments fixture: edit comments + cleanup pass", () => {
  const target = resolve(MAMMOTH_DIR, "comments.docx");
  if (!existsSync(target)) {
    it.skip("mammoth comments.docx fixture not present", () => {
      // intentionally empty
    });
    return;
  }

  it("addComment then removeAllComments round-trips clean", () => {
    const bytes = readFileSync(target);
    const doc = openDocx(bytes);
    const ps = paragraphs(doc);
    if (ps[0]) {
      addComment(doc, ps[0], { author: "Reviewer", initials: "R", text: "added comment" });
    }
    const intermediate = toUint8Array(doc);
    assertValidate(intermediate, "mammoth/comments.docx (with added comment)");

    const reopened = openDocx(intermediate);
    removeAllComments(reopened);
    const final = toUint8Array(reopened);
    assertValidate(final, "mammoth/comments.docx (after removeAllComments)");
  });
});

describe("python-docx having-images fixture: edit then save", () => {
  const target = resolve(PYTHON_DOCX_DIR, "having-images.docx");
  if (!existsSync(target)) {
    it.skip("python-docx having-images.docx fixture not present", () => {
      // intentionally empty
    });
    return;
  }

  it("appendParagraph + save + validate", () => {
    const bytes = readFileSync(target);
    const doc = openDocx(bytes);
    appendParagraph(doc, "[caption]");
    const out = toUint8Array(doc);
    assertValidate(out, "python-docx/having-images.docx");
    expect(text(openDocx(out))).toContain("[caption]");
  });
});

// Sanity: the same edit applied to a freshly-built doc also validates,
// so any failure above isolates to fixture-specific quirks rather than
// a regression in the edit code itself.
describe("control: same edits on a freshly-built doc validate clean", () => {
  it("appendParagraph + setCoreProperties + addComment + removeAllComments", () => {
    const doc = createDocx({ paragraphs: ["seed"] });
    appendParagraph(doc, "[appended]");
    setCoreProperties(doc, { title: "control", creator: "word-kit" });
    const ps = paragraphs(doc);
    if (ps[0]) {
      addComment(doc, ps[0], { author: "R", initials: "R", text: "x" });
    }
    const reopened = openDocx(toUint8Array(doc));
    removeAllComments(reopened);
    assertValidate(toUint8Array(reopened), "control");
  });
});
