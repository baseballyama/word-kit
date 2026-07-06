/**
 * Numbering-level property commands. They edit the `<w:lvl>` of the first
 * abstract numbering definition at the caret paragraph's list level, covering
 * `CT_Lvl` children: numFmt, lvlText, start, lvlRestart, suff, lvlJc, pStyle,
 * isLgl. The list format (decimal/bullet/roman…) is `numFmt`; the label
 * template (`%1.`) is `lvlText`.
 */

import {
  getNumberingLevelProp,
  getParagraphNumbering,
  setNumberingLevelOnOff,
  setNumberingLevelVal,
} from "@office-kit/docx";
import { paragraphAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import type { Command } from "./types.js";

/** The list level (ilvl) of the caret paragraph, defaulting to 0. */
function targetLevel(model: EditorModel): number {
  const pos = model.selection?.focus;
  const para = pos ? paragraphAt(model.doc, pos) : undefined;
  return (para ? getParagraphNumbering(para)?.ilvl : undefined) ?? 0;
}

const LVL_ONOFF = ["isLgl"];
const LVL_VAL = ["start", "numFmt", "lvlRestart", "pStyle", "suff", "lvlText", "lvlJc"];

function lvlToggle(local: string): Command<void> {
  return {
    id: `numbering.${local}`,
    group: "list",
    label: `List level: ${local}`,
    run(model) {
      const ilvl = targetLevel(model);
      setNumberingLevelOnOff(
        model.doc,
        0,
        ilvl,
        local,
        !getNumberingLevelProp(model.doc, 0, ilvl, local).present,
      );
    },
    isActive: (model) => getNumberingLevelProp(model.doc, 0, targetLevel(model), local).present,
  };
}

function lvlVal(local: string): Command<{ val: string | undefined }> {
  return {
    id: `numbering.${local}`,
    group: "list",
    label: `List level: ${local}`,
    run(model, { val }) {
      setNumberingLevelVal(model.doc, 0, targetLevel(model), local, val);
    },
  };
}

export const numberingPropertyCommands = [...LVL_ONOFF.map(lvlToggle), ...LVL_VAL.map(lvlVal)];

/** `w:<local>` element names these commands make editable. */
export const NUMBERING_PROPERTY_ELEMENTS = [...LVL_ONOFF, ...LVL_VAL];
