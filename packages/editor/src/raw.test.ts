/**
 * Tests for the universal raw-XML editor. These back the ledger's decision to
 * classify DrawingML / OMML / VML as `edit`: they prove the raw commands
 * genuinely mutate a raw XML subtree inside `document.xml` and that the change
 * survives serialize → reopen (flushed through the document AST).
 */

import {
  addImage,
  createDocx,
  ensureHeadingStyles,
  getElementAttr,
  getRawPartRoot,
  openDocx,
  toUint8Array,
  validate,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { allRawElements, editorFor, partRawTree, rawTrees, xmlParts } from "./index.js";
import { runCommand } from "./commands/types.js";
import {
  setAttributeCommand,
  setChildValCommand,
  setPartAttributeCommand,
  setPartChildValCommand,
} from "./commands/raw.js";

const PNG_1x1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1,
  0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78,
  0x9c, 0x63, 0, 1, 0, 0, 5, 0, 1, 0x0d, 0x0a, 0x2d, 0xb4, 0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe("raw-XML inspector", () => {
  it("collects the DrawingML subtree of an inline image", () => {
    const model = editorFor(createDocx({ paragraphs: [] }));
    addImage(model.doc, PNG_1x1, { widthEmu: 914400, heightEmu: 914400 });
    const trees = rawTrees(model.doc);
    expect(trees.length).toBe(1);
    expect(trees[0]!.local).toBe("drawing");
    const locals = allRawElements(model.doc).map((e) => e.name.local);
    // Reaches deep DrawingML elements (blip, ext, prstGeom, …).
    expect(locals).toContain("blip");
    expect(locals).toContain("prstGeom");
  });
});

describe("raw editing commands", () => {
  it("sets an attribute on a DrawingML element and round-trips", () => {
    const model = editorFor(createDocx({ paragraphs: [] }));
    addImage(model.doc, PNG_1x1, { widthEmu: 914400, heightEmu: 914400 });
    const prstGeom = allRawElements(model.doc).find((e) => e.name.local === "prstGeom")!;
    runCommand(model, setAttributeCommand, { target: prstGeom, local: "prst", value: "ellipse" });

    const re = openDocx(toUint8Array(model.doc));
    expect(validate(re).length).toBe(0);
    const reGeom = allRawElements(re).find((e) => e.name.local === "prstGeom")!;
    expect(getElementAttr(reGeom, "prst")).toBe("ellipse");
  });

  it("adds a child element on a DrawingML element and round-trips", () => {
    const model = editorFor(createDocx({ paragraphs: [] }));
    addImage(model.doc, PNG_1x1, { widthEmu: 914400, heightEmu: 914400 });
    const spPr = allRawElements(model.doc).find((e) => e.name.local === "spPr")!;
    // Add a rotation/attr-bearing child value; here a simple <a:xfrm>-sibling flag.
    runCommand(model, setChildValCommand, { target: spPr, local: "rot", val: "5400000" });

    const re = openDocx(toUint8Array(model.doc));
    expect(validate(re).length).toBe(0);
    const reSpPr = allRawElements(re).find((e) => e.name.local === "spPr")!;
    expect(reSpPr.children.some((c) => c.kind === "element" && c.name.local === "rot")).toBe(true);
  });
});

describe("part-level raw editor", () => {
  it("lists the XML parts in the package", () => {
    const model = editorFor(createDocx({ paragraphs: ["hi"] }));
    ensureHeadingStyles(model.doc);
    const parts = xmlParts(model.doc);
    expect(parts).toContain("/word/document.xml");
    expect(parts).toContain("/word/styles.xml");
  });

  it("edits an element in styles.xml (a non-document part) and round-trips", () => {
    const model = editorFor(createDocx({ paragraphs: ["hi"] }));
    ensureHeadingStyles(model.doc);
    const tree = partRawTree(model.doc, "/word/styles.xml")!;
    expect(tree.local).toBe("styles");
    const style = tree.children.find((n) => n.local === "style")!;

    // Set an attribute and add a child on a <w:style> via the part commands.
    runCommand(model, setPartAttributeCommand, {
      partName: "/word/styles.xml",
      target: style.element,
      local: "customMarker",
      value: "abc",
    });
    runCommand(model, setPartChildValCommand, {
      partName: "/word/styles.xml",
      target: style.element,
      local: "uiPriority",
      val: "42",
    });

    const re = openDocx(toUint8Array(model.doc));
    expect(validate(re).length).toBe(0);
    const reRoot = getRawPartRoot(re, "/word/styles.xml")!;
    const reStyle = reRoot.children.find((c) => c.kind === "element" && c.name.local === "style");
    if (!reStyle || reStyle.kind !== "element") throw new Error("style not found");
    expect(getElementAttr(reStyle, "customMarker")).toBe("abc");
    expect(
      reStyle.children.some((c) => c.kind === "element" && c.name.local === "uiPriority"),
    ).toBe(true);
  });
});
