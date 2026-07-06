/**
 * Document-settings commands (`word/settings.xml`). Each `<w:settings>` on-off
 * or single-value child gets a `settings.<local>` command backed by the
 * library's `setDocumentSettingOnOff` / `setDocumentSettingVal`. This makes the
 * bulk of `CT_Settings` (spelling/grid/print/track-changes/hyphenation flags,
 * default tab stop, hyphenation zone, …) editable.
 */

import {
  getDocumentSetting,
  setDocumentSettingOnOff,
  setDocumentSettingVal,
} from "@office-kit/docx";
import type { Command } from "./types.js";

/** On-off (`CT_OnOff` / `CT_Empty`) settings. */
const SETTINGS_ONOFF: string[] = [
  "removePersonalInformation",
  "removeDateAndTime",
  "doNotDisplayPageBoundaries",
  "displayBackgroundShape",
  "printPostScriptOverText",
  "printFractionalCharacterWidth",
  "printFormsData",
  "embedTrueTypeFonts",
  "embedSystemFonts",
  "saveSubsetFonts",
  "saveFormsData",
  "mirrorMargins",
  "alignBordersAndEdges",
  "bordersDoNotSurroundHeader",
  "bordersDoNotSurroundFooter",
  "gutterAtTop",
  "hideSpellingErrors",
  "hideGrammaticalErrors",
  "formsDesign",
  "linkStyles",
  "trackRevisions",
  "doNotTrackMoves",
  "doNotTrackFormatting",
  "autoFormatOverride",
  "styleLockTheme",
  "styleLockQFSet",
  "autoHyphenation",
  "doNotHyphenateCaps",
  "showEnvelope",
  "evenAndOddHeaders",
  "bookFoldRevPrinting",
  "bookFoldPrinting",
  "doNotUseMarginsForDrawingGridOrigin",
  "doNotShadeFormData",
  "noPunctuationKerning",
  "printTwoOnOne",
  "strictFirstAndLastChars",
  "savePreviewPicture",
  "doNotValidateAgainstSchema",
  "saveInvalidXml",
  "ignoreMixedContent",
  "alwaysShowPlaceholderText",
  "doNotDemarcateInvalidXml",
  "saveXmlDataOnly",
  "useXSLTWhenSaving",
  "showXMLTags",
  "alwaysMergeEmptyNamespace",
  "updateFields",
  "doNotIncludeSubdocsInStats",
  "doNotAutoCompressPictures",
  "doNotEmbedSmartTags",
  "forceUpgrade",
];

/** Single-value settings (`CT_TwipsMeasure` / `CT_DecimalNumber` / `CT_String` / …). */
const SETTINGS_VAL: string[] = [
  "defaultTabStop",
  "consecutiveHyphenLimit",
  "hyphenationZone",
  "summaryLength",
  "clickAndTypeStyle",
  "defaultTableStyle",
  "bookFoldPrintingSheets",
  "drawingGridHorizontalSpacing",
  "drawingGridVerticalSpacing",
  "displayHorizontalDrawingGridEvery",
  "displayVerticalDrawingGridEvery",
  "drawingGridHorizontalOrigin",
  "drawingGridVerticalOrigin",
  "decimalSymbol",
  "listSeparator",
  "characterSpacingControl",
  "documentType",
];

function settingToggle(local: string): Command<void> {
  return {
    id: `settings.${local}`,
    group: "advanced",
    label: `Setting: ${local}`,
    run(model) {
      setDocumentSettingOnOff(model.doc, local, !getDocumentSetting(model.doc, local).present);
    },
    isActive: (model) => getDocumentSetting(model.doc, local).present,
  };
}

function settingVal(local: string): Command<{ val: string | undefined }> {
  return {
    id: `settings.${local}`,
    group: "advanced",
    label: `Setting: ${local}`,
    run(model, { val }) {
      setDocumentSettingVal(model.doc, local, val);
    },
  };
}

export const settingsCommands = [
  ...SETTINGS_ONOFF.map(settingToggle),
  ...SETTINGS_VAL.map(settingVal),
];

/** `w:<local>` element names these commands make editable (all → `settings.<local>`). */
export const SETTINGS_ELEMENTS = [...SETTINGS_ONOFF, ...SETTINGS_VAL];
