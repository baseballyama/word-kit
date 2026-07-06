/**
 * Round-trip tests for the settings / style-definition / numbering-level
 * property commands added in the coverage-expansion phase.
 */

import {
  addBulletList,
  addImage,
  createDocx,
  ensureHeadingStyles,
  getDocumentSetting,
  getImageInfo,
  getNumberingLevelProp,
  getParagraphStyle,
  getStyleProp,
  openDocx,
  paragraphs,
  setParagraphStyle,
  toUint8Array,
  validate,
} from "@office-kit/docx";
import { describe, expect, it } from "vitest";
import { editorFor } from "../index.js";
import type { EditorModel } from "../model.js";
import { caretAt } from "../selection.js";
import { getCommand } from "./registry.js";
import { runCommand } from "./types.js";

function reopen(model: EditorModel) {
  const re = openDocx(toUint8Array(model.doc));
  expect(validate(re).length).toBe(0);
  return re;
}

describe("settings commands", () => {
  it("toggles evenAndOddHeaders and sets defaultTabStop", () => {
    const model = editorFor(createDocx({ paragraphs: ["x"] }));
    model.setSelection(caretAt({ block: 0 }));
    runCommand(model, getCommand("settings.evenAndOddHeaders")!, undefined);
    runCommand(model, getCommand("settings.defaultTabStop")!, { val: "720" });
    const re = reopen(model);
    expect(getDocumentSetting(re, "evenAndOddHeaders").present).toBe(true);
    expect(getDocumentSetting(re, "defaultTabStop").val).toBe("720");
  });
});

describe("style-definition commands", () => {
  it("toggles a flag on the caret paragraph's style", () => {
    const model = editorFor(createDocx({ paragraphs: ["Heading text"] }));
    ensureHeadingStyles(model.doc);
    setParagraphStyle(paragraphs(model.doc)[0]!, "Heading1");
    model.setSelection(caretAt({ block: 0 }));
    expect(getParagraphStyle(paragraphs(model.doc)[0]!)).toBe("Heading1");
    // `hidden` is not set on the built-in Heading1, so toggling adds it.
    expect(getStyleProp(model.doc, "Heading1", "hidden").present).toBe(false);
    runCommand(model, getCommand("style.hidden")!, undefined);
    const re = reopen(model);
    expect(getStyleProp(re, "Heading1", "hidden").present).toBe(true);
  });
});

const PNG_1x1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1,
  0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78,
  0x9c, 0x63, 0, 1, 0, 0, 5, 0, 1, 0x0d, 0x0a, 0x2d, 0xb4, 0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe("image commands", () => {
  it("resizes and alt-texts an existing image", () => {
    const model = editorFor(createDocx({ paragraphs: [] }));
    addImage(model.doc, PNG_1x1, { widthEmu: 914400, heightEmu: 914400, altText: "orig" });
    model.setSelection(caretAt({ block: 0 }));
    runCommand(model, getCommand("image.resize")!, { index: 0, cxEmu: 457200, cyEmu: 228600 });
    runCommand(model, getCommand("image.altText")!, { index: 0, descr: "a cat", title: "Cat" });
    const re = reopen(model);
    const info = getImageInfo(re, 0)!;
    expect(info.cx).toBe("457200");
    expect(info.cy).toBe("228600");
    expect(info.descr).toBe("a cat");
    expect(info.title).toBe("Cat");
  });
});

describe("numbering-level commands", () => {
  it("sets numFmt on the list level", () => {
    const model = editorFor(createDocx({ paragraphs: [] }));
    addBulletList(model.doc, ["item"]);
    model.setSelection(caretAt({ block: 0 }));
    runCommand(model, getCommand("numbering.numFmt")!, { val: "decimal" });
    const re = reopen(model);
    expect(getNumberingLevelProp(re, 0, 0, "numFmt").val).toBe("decimal");
  });
});
