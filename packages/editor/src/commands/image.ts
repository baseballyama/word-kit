/**
 * Image commands: insert an inline image at the caret, and replace an existing
 * image's bytes. Sizing is in EMU (914400 per inch), matching `AddImageOptions`.
 */

import {
  addImage,
  type AddImageOptions,
  imageDrawings,
  images,
  replaceImage,
  setImageAltText,
  setImageSizeEmu,
} from "@office-kit/docx";
import { caretBlockIndex, moveLastBlockAfter } from "./insert-util.js";
import type { Command } from "./types.js";

export const insertImageCommand: Command<{ bytes: Uint8Array; options: AddImageOptions }> = {
  id: "image.insert",
  group: "image",
  label: "Insert picture",
  run(model, { bytes, options }) {
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    addImage(model.doc, bytes, options);
    moveLastBlockAfter(model.doc, at);
  },
};

export const replaceImageCommand: Command<{ index: number; bytes: Uint8Array }> = {
  id: "image.replace",
  group: "image",
  label: "Replace picture",
  run(model, { index, bytes }) {
    const refs = images(model.doc);
    const ref = refs[index];
    if (!ref) return;
    replaceImage(model.doc, ref.partName, bytes);
  },
  isEnabled: (model) => images(model.doc).length > 0,
};

/** Resize an existing image (EMU: 914400 per inch). */
export const resizeImageCommand: Command<{ index: number; cxEmu: number; cyEmu: number }> = {
  id: "image.resize",
  group: "image",
  label: "Resize picture",
  run(model, { index, cxEmu, cyEmu }) {
    setImageSizeEmu(model.doc, index, cxEmu, cyEmu);
  },
  isEnabled: (model) => imageDrawings(model.doc).length > 0,
};

/** Set alt text (description / title) on an existing image. */
export const setImageAltCommand: Command<{ index: number; descr?: string; title?: string }> = {
  id: "image.altText",
  group: "image",
  label: "Picture alt text",
  run(model, { index, descr, title }) {
    setImageAltText(model.doc, index, {
      ...(descr !== undefined ? { descr } : {}),
      ...(title !== undefined ? { title } : {}),
    });
  },
  isEnabled: (model) => imageDrawings(model.doc).length > 0,
};

export const imageCommands = [
  insertImageCommand,
  replaceImageCommand,
  resizeImageCommand,
  setImageAltCommand,
];
