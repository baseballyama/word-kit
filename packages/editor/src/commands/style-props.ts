/**
 * Style-definition property commands. They edit the `<w:style>` that the caret
 * paragraph currently uses (its `pStyle`), covering the on-off and single-value
 * children of `CT_Style` (qFormat, hidden, semiHidden, uiPriority, basedOn, …).
 */

import { getParagraphStyle, getStyleProp, setStyleOnOff, setStyleValProp } from "@office-kit/docx";
import { paragraphAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import type { Command } from "./types.js";

/** The style id of the caret paragraph, if it has one. */
function targetStyleId(model: EditorModel): string | undefined {
  const pos = model.selection?.focus;
  const para = pos ? paragraphAt(model.doc, pos) : undefined;
  return para ? getParagraphStyle(para) : undefined;
}

const STYLE_ONOFF = [
  "autoRedefine",
  "hidden",
  "semiHidden",
  "unhideWhenUsed",
  "qFormat",
  "locked",
  "personal",
  "personalCompose",
  "personalReply",
];
const STYLE_VAL = ["name", "aliases", "basedOn", "next", "link", "uiPriority", "rsid"];

function styleToggle(local: string): Command<void> {
  return {
    id: `style.${local}`,
    group: "style",
    label: `Style: ${local}`,
    run(model) {
      const id = targetStyleId(model);
      if (id) setStyleOnOff(model.doc, id, local, !getStyleProp(model.doc, id, local).present);
    },
    isEnabled: (model) => !!targetStyleId(model),
    isActive: (model) => {
      const id = targetStyleId(model);
      return !!id && getStyleProp(model.doc, id, local).present;
    },
  };
}

function styleVal(local: string): Command<{ val: string | undefined }> {
  return {
    id: `style.${local}`,
    group: "style",
    label: `Style: ${local}`,
    run(model, { val }) {
      const id = targetStyleId(model);
      if (id) setStyleValProp(model.doc, id, local, val);
    },
    isEnabled: (model) => !!targetStyleId(model),
  };
}

export const stylePropertyCommands = [...STYLE_ONOFF.map(styleToggle), ...STYLE_VAL.map(styleVal)];

/** `w:<local>` element names these commands make editable. */
export const STYLE_PROPERTY_ELEMENTS = [...STYLE_ONOFF, ...STYLE_VAL];
