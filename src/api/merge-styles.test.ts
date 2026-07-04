// `mergeStylesFromTemplate` is the natural mate to the styled-base
// template workflow: when the design system lives in one .docx but
// the program is producing another (or wants the design grafted onto
// a fresh doc), this helper imports just the `word/styles.xml`
// definitions without dragging the body / header / footer along.

import { describe, expect, it } from "vitest";
import {
  addStyle,
  appendParagraph,
  createDocx,
  listStyles,
  mergeStylesFromTemplate,
  openDocx,
  paragraphs,
  toUint8Array,
  validate,
} from "./index.js";

function buildTemplate(): Uint8Array {
  const tpl = createDocx({ paragraphs: [] });
  addStyle(tpl, {
    type: "paragraph",
    styleId: "AcmeTitle",
    name: "Acme Title",
    qFormat: true,
    bold: true,
    fontSizeHalfPoints: 44,
    color: "1F497D",
  });
  addStyle(tpl, {
    type: "paragraph",
    styleId: "AcmeBody",
    name: "Acme Body",
    qFormat: true,
    fontSizeHalfPoints: 22,
  });
  // Body content the merge MUST NOT pull in.
  appendParagraph(tpl, "TEMPLATE_BODY_SHOULD_NOT_LEAK", { style: "AcmeTitle" });
  return toUint8Array(tpl);
}

describe("mergeStylesFromTemplate", () => {
  it("copies style definitions but not the template's body", () => {
    const tplBytes = buildTemplate();
    const target = createDocx({ paragraphs: ["fresh body"] });
    const n = mergeStylesFromTemplate(target, tplBytes);
    expect(n).toBeGreaterThanOrEqual(2);
    const ids = listStyles(target).map((s) => s.styleId);
    expect(ids).toContain("AcmeTitle");
    expect(ids).toContain("AcmeBody");
    // Body should be untouched.
    expect(paragraphs(target).map((p) => p.children).length).toBe(1);
  });

  it("does not overwrite existing same-id styles by default", () => {
    const tplBytes = buildTemplate();
    const target = createDocx({ paragraphs: [] });
    addStyle(target, {
      type: "paragraph",
      styleId: "AcmeTitle",
      name: "Local Override",
      qFormat: true,
      italic: true,
      color: "00FF00",
    });
    const n = mergeStylesFromTemplate(target, tplBytes);
    // AcmeTitle skipped (already present), AcmeBody copied.
    expect(n).toBe(1);
    // Verify the local definition won.
    const reopened = openDocx(toUint8Array(target));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/styles.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain('w:val="Local Override"');
    expect(xml).not.toContain('w:val="Acme Title"');
  });

  it("honours overwrite: true to replace existing same-id styles", () => {
    const tplBytes = buildTemplate();
    const target = createDocx({ paragraphs: [] });
    addStyle(target, {
      type: "paragraph",
      styleId: "AcmeTitle",
      name: "Local Override",
      bold: false,
      italic: true,
    });
    const n = mergeStylesFromTemplate(target, tplBytes, { overwrite: true });
    // overwrite: true copies every template style; both the user-added
    // AcmeTitle / AcmeBody and the template's seed defaults (Normal,
    // Heading1, …) are written through. The exact total is just "all
    // of the template's styles".
    expect(n).toBeGreaterThanOrEqual(2);
    const reopened = openDocx(toUint8Array(target));
    const xml = new TextDecoder().decode(
      reopened.opc.parts.get("/word/styles.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain('w:val="Acme Title"');
    expect(xml).not.toContain('w:val="Local Override"');
  });

  it("returns 0 if the template has no styles part", () => {
    // Build a minimal package-like docx without styles.xml is hard to do
    // through the public API; emulate the "nothing to merge" case by
    // passing the *target's own* bytes to itself — every style in the
    // target's own styles.xml is already present, so the count is 0.
    const target = createDocx({ paragraphs: [] });
    const n = mergeStylesFromTemplate(target, toUint8Array(target));
    expect(n).toBe(0);
  });

  it("merged styles survive save+reopen and the doc validates", () => {
    const tplBytes = buildTemplate();
    const target = createDocx({ paragraphs: ["body"] });
    mergeStylesFromTemplate(target, tplBytes);
    appendParagraph(target, "Title via merged style", { style: "AcmeTitle" });
    const reopened = openDocx(toUint8Array(target));
    expect(validate(reopened).filter((i) => i.level === "error")).toHaveLength(0);
    expect(listStyles(reopened).map((s) => s.styleId)).toContain("AcmeTitle");
  });
});
