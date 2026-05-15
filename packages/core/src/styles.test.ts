import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { findStyle, isDefaultStyle, styleBasedOn, styleId, styleType } from "@word-kit/wml";
import { describe, expect, it } from "vitest";
import { addStyle, createDocx, openDocx, stylesPart, toUint8Array } from "./docx.js";

const REPO_ROOT = resolve(import.meta.dirname, "../../..");

describe("Docx styles", () => {
  it("Docx.create seeds a minimal styles.xml", () => {
    const doc = createDocx();
    const part = stylesPart(doc);
    expect(part).toBeDefined();
    if (!part) return;
    expect(part.styles.length).toBeGreaterThan(0);
    expect(part.styles.map(styleId)).toContain("Normal");
    const normal = findStyle(part, "Normal");
    expect(normal).toBeDefined();
    expect(isDefaultStyle(normal!)).toBe(true);
    expect(styleType(normal!)).toBe("paragraph");
  });

  it("addStyle inserts a new <w:style> and persists through save+reopen", () => {
    const doc = createDocx();
    addStyle(doc, {
      type: "paragraph",
      styleId: "MyHeading",
      name: "My Heading",
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      bold: true,
      fontSizeHalfPoints: 32,
      color: "1F497D",
    });
    const bytes = toUint8Array(doc);
    const reopened = openDocx(bytes);
    const part = stylesPart(reopened);
    expect(part).toBeDefined();
    if (!part) return;
    const mine = findStyle(part, "MyHeading");
    expect(mine).toBeDefined();
    expect(styleBasedOn(mine!)).toBe("Normal");
    expect(styleType(mine!)).toBe("paragraph");
  });

  it("addStyle replaces an existing style with the same id", () => {
    const doc = createDocx();
    addStyle(doc, { type: "paragraph", styleId: "MyHeading", name: "First", bold: true });
    addStyle(doc, { type: "paragraph", styleId: "MyHeading", name: "Second", italic: true });
    const part = stylesPart(doc);
    expect(part).toBeDefined();
    if (!part) return;
    const all = part.styles.filter((s) => styleId(s) === "MyHeading");
    expect(all).toHaveLength(1);
  });

  it("opens a real Word docx and reads its style list", (ctx) => {
    const fixture = resolve(
      REPO_ROOT,
      "references/mammoth.js/test/test-data/single-paragraph.docx",
    );
    if (!existsSync(fixture)) {
      ctx.skip();
      return;
    }
    const bytes = readFileSync(fixture);
    const doc = openDocx(bytes);
    const part = stylesPart(doc);
    expect(part).toBeDefined();
    if (!part) return;
    expect(part.styles.length).toBeGreaterThan(0);
  });
});
