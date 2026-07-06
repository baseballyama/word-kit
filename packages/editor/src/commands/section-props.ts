/**
 * Section property commands for the long tail of `<w:sectPr>`: title page,
 * vertical alignment, text direction, RTL, form protection, endnote
 * suppression. These ensure the body's trailing `<w:sectPr>` exists, then set
 * the property via the library's container-level setters.
 */

import {
  getElementProp,
  makePropsElement,
  setElementOnOff,
  setElementValProp,
  type XmlElement,
} from "@office-kit/docx";
import type { EditorModel } from "../model.js";
import type { Command } from "./types.js";

/** Ensure the body has a trailing `<w:sectPr>` and return it. */
function ensureSectPr(model: EditorModel): XmlElement {
  const body = model.doc.document.body;
  if (!body.sectPr) body.sectPr = makePropsElement("sectPr");
  return body.sectPr;
}

const SECT_ONOFF = [
  { local: "titlePg", label: "Different first page" },
  { local: "bidi", label: "Right-to-left section" },
  { local: "rtlGutter", label: "RTL gutter" },
  { local: "noEndnote", label: "Suppress endnotes" },
  { local: "formProt", label: "Protect form" },
];
const SECT_VAL = [
  { local: "textDirection", label: "Section text direction" },
  { local: "vAlign", label: "Vertical page alignment" }, // top | center | both | bottom
];

export const sectionPropertyCommands: Command<never>[] = [
  ...SECT_ONOFF.map(
    ({ local, label }): Command<void> => ({
      id: `section.${local}`,
      group: "section",
      label,
      run(model) {
        setElementOnOff(
          ensureSectPr(model),
          local,
          !getElementProp(model.doc.document.body.sectPr, local).present,
        );
      },
      isActive: (model) => getElementProp(model.doc.document.body.sectPr, local).present,
    }),
  ),
  ...SECT_VAL.map(
    ({ local, label }): Command<{ val: string | undefined }> => ({
      id: `section.${local}`,
      group: "section",
      label,
      run(model, { val }) {
        setElementValProp(ensureSectPr(model), local, val);
      },
    }),
  ),
] as Command<never>[];

/** `w:<local>` element names these commands make editable. */
export const SECTION_PROPERTY_ELEMENTS = [...SECT_ONOFF, ...SECT_VAL].map((e) => e.local);
