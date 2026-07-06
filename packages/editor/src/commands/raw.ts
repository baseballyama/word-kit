/**
 * Universal raw-XML editing commands. They operate on any `XmlElement` in the
 * document AST (typically a DrawingML / OMML / VML node reached via the
 * raw-XML inspector), setting attributes or child elements. This is the escape
 * hatch that makes every OOXML element editable — the mutation flushes through
 * the document AST on save, so DrawingML shape geometry, math, and legacy VML
 * are all reachable even without a bespoke command per element.
 */

import {
  appendChildElement,
  getElementAttr,
  getElementProp,
  markRawPartDirty,
  setElementAttr,
  setElementOnOff,
  setElementValProp,
  type XmlElement,
} from "@office-kit/docx";
import type { Command } from "./types.js";

/** Set (or remove, with `undefined`) an attribute on a raw element. */
export const setAttributeCommand: Command<{
  target: XmlElement;
  local: string;
  value: string | undefined;
}> = {
  id: "raw.setAttribute",
  group: "advanced",
  label: "Set XML attribute",
  run(_model, { target, local, value }) {
    setElementAttr(target, local, value);
  },
};

/** Toggle an on-off child element (`<x/>`) on a raw element. */
export const setChildOnOffCommand: Command<{
  target: XmlElement;
  local: string;
  on: boolean;
}> = {
  id: "raw.setChildOnOff",
  group: "advanced",
  label: "Toggle XML child",
  run(_model, { target, local, on }) {
    setElementOnOff(target, local, on);
  },
};

/** Set a single-value child element (`<x val="…"/>`) on a raw element. */
export const setChildValCommand: Command<{
  target: XmlElement;
  local: string;
  val: string | undefined;
}> = {
  id: "raw.setChildVal",
  group: "advanced",
  label: "Set XML child value",
  run(_model, { target, local, val }) {
    setElementValProp(target, local, val);
  },
};

/** Append a new child element to a raw element and (optionally) read it back. */
export const addChildCommand: Command<{ target: XmlElement; local: string }> = {
  id: "raw.addChild",
  group: "advanced",
  label: "Add XML child",
  run(_model, { target, local }) {
    appendChildElement(target, local);
  },
};

// --- Part-level raw editing ---------------------------------------------------
//
// Same operations, but on an element that belongs to a non-document XML part
// (fontTable / settings / styles / numbering / comments / notes / headers /
// footers / webSettings / docProps). The part must be marked dirty so its bytes
// are re-serialized on save.

/** Set (or remove) an attribute on an element inside an arbitrary XML part. */
export const setPartAttributeCommand: Command<{
  partName: string;
  target: XmlElement;
  local: string;
  value: string | undefined;
}> = {
  id: "rawpart.setAttribute",
  group: "advanced",
  label: "Set part XML attribute",
  run(model, { partName, target, local, value }) {
    setElementAttr(target, local, value);
    markRawPartDirty(model.doc, partName);
  },
};

/** Toggle an on-off child element on an element inside an arbitrary XML part. */
export const setPartChildOnOffCommand: Command<{
  partName: string;
  target: XmlElement;
  local: string;
  on: boolean;
}> = {
  id: "rawpart.setChildOnOff",
  group: "advanced",
  label: "Toggle part XML child",
  run(model, { partName, target, local, on }) {
    setElementOnOff(target, local, on);
    markRawPartDirty(model.doc, partName);
  },
};

/** Set a single-value child element on an element inside an arbitrary XML part. */
export const setPartChildValCommand: Command<{
  partName: string;
  target: XmlElement;
  local: string;
  val: string | undefined;
}> = {
  id: "rawpart.setChildVal",
  group: "advanced",
  label: "Set part XML child value",
  run(model, { partName, target, local, val }) {
    setElementValProp(target, local, val);
    markRawPartDirty(model.doc, partName);
  },
};

/** Append a new child element to an element inside an arbitrary XML part. */
export const addPartChildCommand: Command<{
  partName: string;
  target: XmlElement;
  local: string;
}> = {
  id: "rawpart.addChild",
  group: "advanced",
  label: "Add part XML child",
  run(model, { partName, target, local }) {
    appendChildElement(target, local);
    markRawPartDirty(model.doc, partName);
  },
};

/** Read helpers re-exported so an inspector UI can render current values. */
export { getElementAttr, getElementProp };

export const rawCommands = [
  setAttributeCommand,
  setChildOnOffCommand,
  setChildValCommand,
  addChildCommand,
  setPartAttributeCommand,
  setPartChildOnOffCommand,
  setPartChildValCommand,
  addPartChildCommand,
];
