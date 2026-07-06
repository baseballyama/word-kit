/**
 * Generic run/paragraph property commands.
 *
 * Rather than a bespoke command per WordprocessingML formatting element, these
 * factories drive the library's generic property setters (`setRunOnOff` /
 * `setRunValProp` / `setParagraphOnOff` / `setParagraphValProp`) so the whole
 * long tail of `<w:rPr>` / `<w:pPr>` on-off and single-value formatting becomes
 * editable. Each generated command id is `text.<local>` or `paragraph.<local>`,
 * which the capability ledger maps `w:<local>` onto.
 */

import {
  getParagraphProp,
  getRunProp,
  setParagraphOnOff,
  setParagraphValProp,
  setRunOnOff,
  setRunValProp,
  type WmlParagraph,
} from "@office-kit/docx";
import { paragraphsInRange } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { orderSelection } from "../selection.js";
import { applyToSelectionRuns, overlappingSelectionRuns } from "../selection-runs.js";
import type { Command } from "./types.js";

function selectedParagraphs(model: EditorModel): WmlParagraph[] {
  const sel = model.selection;
  return sel ? paragraphsInRange(model.doc, orderSelection(sel)) : [];
}

/** On-off `<w:rPr>` elements that toggle a character effect. */
const RUN_ONOFF: { local: string; label: string }[] = [
  { local: "caps", label: "All caps" },
  { local: "smallCaps", label: "Small caps" },
  { local: "dstrike", label: "Double strikethrough" },
  { local: "outline", label: "Outline" },
  { local: "shadow", label: "Shadow" },
  { local: "emboss", label: "Emboss" },
  { local: "imprint", label: "Engrave" },
  { local: "vanish", label: "Hidden" },
  { local: "specVanish", label: "Always hidden" },
  { local: "webHidden", label: "Web hidden" },
  { local: "noProof", label: "Do not check spelling" },
  { local: "snapToGrid", label: "Snap to grid" },
  { local: "rtl", label: "Right-to-left" },
  { local: "cs", label: "Complex script" },
  { local: "oMath", label: "Office math run" },
];

/** Single-value `<w:rPr>` elements. */
const RUN_VAL: { local: string; label: string }[] = [
  { local: "vertAlign", label: "Superscript/subscript" }, // baseline | superscript | subscript
  { local: "em", label: "Emphasis mark" }, // none | dot | comma | circle | underDot
  { local: "position", label: "Character position" }, // signed half-points
  { local: "kern", label: "Kerning" }, // half-points
  { local: "spacing", label: "Character spacing" }, // signed twips
  { local: "w", label: "Character scale" }, // percent
  { local: "effect", label: "Text effect" }, // blinkBackground | lights | …
  { local: "rStyle", label: "Character style" }, // styleId
];

/** On-off `<w:pPr>` elements. */
const PARA_ONOFF: { local: string; label: string }[] = [
  { local: "keepNext", label: "Keep with next" },
  { local: "keepLines", label: "Keep lines together" },
  { local: "pageBreakBefore", label: "Page break before" },
  { local: "widowControl", label: "Widow/orphan control" },
  { local: "suppressLineNumbers", label: "Suppress line numbers" },
  { local: "suppressAutoHyphens", label: "Suppress hyphenation" },
  { local: "kinsoku", label: "Kinsoku (line breaking)" },
  { local: "wordWrap", label: "Allow latin word-break" },
  { local: "overflowPunct", label: "Overflow punctuation" },
  { local: "topLinePunct", label: "Compress top-line punctuation" },
  { local: "autoSpaceDE", label: "Auto-space CJK/Latin" },
  { local: "autoSpaceDN", label: "Auto-space CJK/number" },
  { local: "bidi", label: "Right-to-left paragraph" },
  { local: "adjustRightInd", label: "Auto-adjust right indent" },
  { local: "snapToGrid", label: "Snap to grid" },
  { local: "contextualSpacing", label: "No space between same style" },
  { local: "mirrorIndents", label: "Mirror indents" },
  { local: "suppressOverlap", label: "Suppress frame overlap" },
];

/** Single-value `<w:pPr>` elements. */
const PARA_VAL: { local: string; label: string }[] = [
  { local: "textDirection", label: "Text direction" }, // lrTb | tbRl | …
  { local: "textAlignment", label: "Vertical text alignment" }, // top | center | baseline | bottom | auto
  { local: "outlineLvl", label: "Outline level" }, // 0-9
  { local: "divId", label: "HTML div id" },
];

function runToggle({ local, label }: { local: string; label: string }): Command<void> {
  return {
    id: `text.${local}`,
    group: "text",
    label,
    run(model) {
      const runs = overlappingSelectionRuns(model);
      const next = !runs.every((r) => getRunProp(r, local).present) || runs.length === 0;
      applyToSelectionRuns(model, (r) => setRunOnOff(r, local, next));
    },
    isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
    isActive: (model) => {
      const runs = overlappingSelectionRuns(model);
      return runs.length > 0 && runs.every((r) => getRunProp(r, local).present);
    },
  };
}

function runVal({
  local,
  label,
}: {
  local: string;
  label: string;
}): Command<{ val: string | undefined }> {
  return {
    id: `text.${local}`,
    group: "text",
    label,
    run(model, { val }) {
      applyToSelectionRuns(model, (r) => setRunValProp(r, local, val));
    },
    isEnabled: (model) => overlappingSelectionRuns(model).length > 0,
  };
}

function paraToggle({ local, label }: { local: string; label: string }): Command<void> {
  return {
    id: `paragraph.${local}`,
    group: "paragraph",
    label,
    run(model) {
      const paras = selectedParagraphs(model);
      const next = !paras.every((p) => getParagraphProp(p, local).present) || paras.length === 0;
      for (const p of paras) setParagraphOnOff(p, local, next);
    },
    isEnabled: (model) => selectedParagraphs(model).length > 0,
    isActive: (model) => {
      const paras = selectedParagraphs(model);
      return paras.length > 0 && paras.every((p) => getParagraphProp(p, local).present);
    },
  };
}

function paraVal({
  local,
  label,
}: {
  local: string;
  label: string;
}): Command<{ val: string | undefined }> {
  return {
    id: `paragraph.${local}`,
    group: "paragraph",
    label,
    run(model, { val }) {
      for (const p of selectedParagraphs(model)) setParagraphValProp(p, local, val);
    },
    isEnabled: (model) => selectedParagraphs(model).length > 0,
  };
}

export const runPropertyCommands = [...RUN_ONOFF.map(runToggle), ...RUN_VAL.map(runVal)];
export const paragraphPropertyCommands = [...PARA_ONOFF.map(paraToggle), ...PARA_VAL.map(paraVal)];

/** The set of `w:<local>` element names these commands make editable. */
export const RUN_PROPERTY_ELEMENTS = [...RUN_ONOFF, ...RUN_VAL].map((e) => e.local);
export const PARA_PROPERTY_ELEMENTS = [...PARA_ONOFF, ...PARA_VAL].map((e) => e.local);

export const propertyCommands = [...runPropertyCommands, ...paragraphPropertyCommands];
