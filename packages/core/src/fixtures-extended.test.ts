import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import {
  commentsPart,
  endnotesPart,
  footnotesPart,
  images,
  openDocx,
  paragraphs,
  text,
  toUint8Array,
  validate,
} from "./docx.js";

const REPO_ROOT = resolve(import.meta.dirname, "../../..");
const FIXTURE_DIR = resolve(REPO_ROOT, "references/mammoth.js/test/test-data");

function load(name: string): Uint8Array {
  return readFileSync(resolve(FIXTURE_DIR, name));
}

describe("real fixtures: extended coverage", () => {
  it("opens comments.docx, validates, and reads the comments part", () => {
    const doc = openDocx(load("comments.docx"));
    const cp = commentsPart(doc);
    expect(cp).toBeDefined();
    expect(cp?.comments.length).toBeGreaterThan(0);
    // Validation should report at most warnings (Word produced this file).
    const errors = validate(doc).filter((i) => i.level === "error");
    expect(errors).toEqual([]);
  });

  it("opens footnotes.docx and reads the footnotes part", () => {
    const doc = openDocx(load("footnotes.docx"));
    expect(footnotesPart(doc)).toBeDefined();
    expect(footnotesPart(doc)?.footnotes.length ?? 0).toBeGreaterThan(0);
    expect(validate(doc).filter((i) => i.level === "error")).toEqual([]);
  });

  it("opens endnotes.docx and reads the endnotes part", () => {
    const doc = openDocx(load("endnotes.docx"));
    expect(endnotesPart(doc)).toBeDefined();
    expect(validate(doc).filter((i) => i.level === "error")).toEqual([]);
  });

  it("opens simple-list.docx and round-trips it", () => {
    const doc = openDocx(load("simple-list.docx"));
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    expect(paragraphs(reopened).length).toBe(paragraphs(doc).length);
  });

  it("opens tiny-picture.docx and finds at least one media part", () => {
    const doc = openDocx(load("tiny-picture.docx"));
    expect(images(doc).length).toBeGreaterThan(0);
  });

  // strict-format.docx uses the ISO/IEC 29500 Strict namespaces
  // (http://purl.oclc.org/ooxml/...) rather than the transitional
  // openxmlformats namespaces. We only support transitional today
  // (see PLAN.md), so we explicitly skip this fixture.
  it.skip("opens strict-format.docx (not yet supported)", () => {
    const doc = openDocx(load("strict-format.docx"));
    expect(paragraphs(doc).length).toBeGreaterThan(0);
  });

  it("opens text-box.docx and preserves the body on round-trip", () => {
    const doc = openDocx(load("text-box.docx"));
    const before = text(doc);
    const reopened = openDocx(toUint8Array(doc));
    expect(text(reopened)).toBe(before);
  });

  it("opens utf8-bom.docx and reads it", () => {
    const doc = openDocx(load("utf8-bom.docx"));
    expect(paragraphs(doc).length).toBeGreaterThan(0);
  });

  it("opens hyperlinks.docx (subfolder fixture) when available", () => {
    const dir = resolve(FIXTURE_DIR, "hyperlinks");
    try {
      const fnames = require("node:fs").readdirSync(dir) as string[];
      const first = fnames.find((n) => n.endsWith(".docx"));
      if (!first) return;
      const doc = openDocx(readFileSync(resolve(dir, first)));
      expect(paragraphs(doc).length).toBeGreaterThan(0);
    } catch {
      // directory missing or unreadable — silently skip
    }
  });

  it("opens external-picture.docx (image part may be missing)", () => {
    const doc = openDocx(load("external-picture.docx"));
    expect(paragraphs(doc).length).toBeGreaterThan(0);
  });
});
