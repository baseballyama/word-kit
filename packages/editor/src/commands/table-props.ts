/**
 * Generic table / row / cell property commands. These reach the long tail of
 * `<w:tblPr>` / `<w:trPr>` / `<w:tcPr>` on-off and single-value formatting via
 * the library's container-level property setters, ensuring the container
 * element exists on the current table/row/cell first.
 */

import {
  getElementProp,
  makePropsElement,
  setElementOnOff,
  setElementValProp,
  type WmlTable,
  type WmlTableCell,
  type WmlTableRow,
  type XmlElement,
} from "@office-kit/docx";
import { asTable, blockAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import type { Command } from "./types.js";

function currentTable(model: EditorModel): WmlTable | undefined {
  const block = model.selection?.focus.block;
  return block === undefined ? undefined : asTable(blockAt(model.doc, block));
}

function currentCell(model: EditorModel): WmlTableCell | undefined {
  const table = currentTable(model);
  const cell = model.selection?.focus.cell;
  return table && cell ? table.rows[cell.row]?.cells[cell.col] : undefined;
}

function currentRow(model: EditorModel): WmlTableRow | undefined {
  const table = currentTable(model);
  const cell = model.selection?.focus.cell;
  return table && cell ? table.rows[cell.row] : undefined;
}

/** Ensure `<w:tblPr>` exists on a table and return it. */
function ensureTblPr(table: WmlTable): XmlElement {
  if (!table.tblPr) table.tblPr = makePropsElement("tblPr");
  return table.tblPr;
}
function ensureTrPr(row: WmlTableRow): XmlElement {
  if (!row.trPr) row.trPr = makePropsElement("trPr");
  return row.trPr;
}
function ensureTcPr(cell: WmlTableCell): XmlElement {
  if (!cell.tcPr) cell.tcPr = makePropsElement("tcPr");
  return cell.tcPr;
}

const TBL_ONOFF = [{ local: "bidiVisual", label: "Right-to-left table" }];
const TBL_VAL = [
  { local: "tblStyle", label: "Table style" },
  { local: "tblLayout", label: "Table layout" }, // fixed | autofit
  { local: "tblCaption", label: "Table caption" },
  { local: "tblDescription", label: "Table description" },
  { local: "jc", label: "Table alignment" },
];
const TR_ONOFF = [
  { local: "cantSplit", label: "Can't split row" },
  { local: "hidden", label: "Hidden row" },
];
const TR_VAL = [
  { local: "jc", label: "Row alignment" },
  { local: "gridBefore", label: "Cells before" },
  { local: "gridAfter", label: "Cells after" },
];
const TC_ONOFF = [
  { local: "noWrap", label: "No wrap" },
  { local: "hideMark", label: "Hide end-of-cell mark" },
];
const TC_VAL = [
  { local: "gridSpan", label: "Column span" },
  { local: "vMerge", label: "Vertical merge" }, // restart | continue
  { local: "hMerge", label: "Horizontal merge" },
  { local: "textDirection", label: "Cell text direction" },
];

type Ent = { local: string; label: string };

function containerToggle(
  id: string,
  { local, label }: Ent,
  group: "table",
  ensure: (m: EditorModel) => XmlElement | undefined,
  read: (m: EditorModel) => XmlElement | undefined,
): Command<void> {
  return {
    id,
    group,
    label,
    run(model) {
      const c = ensure(model);
      if (c) setElementOnOff(c, local, !getElementProp(read(model), local).present);
    },
    isEnabled: (model) => !!read(model),
    isActive: (model) => getElementProp(read(model), local).present,
  };
}

function containerVal(
  id: string,
  { local, label }: Ent,
  ensure: (m: EditorModel) => XmlElement | undefined,
  scope: (m: EditorModel) => boolean,
): Command<{ val: string | undefined }> {
  return {
    id,
    group: "table",
    label,
    run(model, { val }) {
      const c = ensure(model);
      if (c) setElementValProp(c, local, val);
    },
    isEnabled: scope,
  };
}

const tblPrOf = (m: EditorModel) => {
  const t = currentTable(m);
  return t ? ensureTblPr(t) : undefined;
};
const trPrOf = (m: EditorModel) => {
  const r = currentRow(m);
  return r ? ensureTrPr(r) : undefined;
};
const tcPrOf = (m: EditorModel) => {
  const c = currentCell(m);
  return c ? ensureTcPr(c) : undefined;
};

const hasTable = (m: EditorModel) => !!currentTable(m);
const hasRow = (m: EditorModel) => !!currentRow(m);
const hasCell = (m: EditorModel) => !!currentCell(m);

export const tablePropertyCommands: Command<never>[] = [
  ...TBL_ONOFF.map((e) =>
    containerToggle(`table.${e.local}`, e, "table", tblPrOf, (m) => currentTable(m)?.tblPr),
  ),
  ...TBL_VAL.map((e) => containerVal(`table.tbl_${e.local}`, e, tblPrOf, hasTable)),
  ...TR_ONOFF.map((e) =>
    containerToggle(`table.row_${e.local}`, e, "table", trPrOf, (m) => currentRow(m)?.trPr),
  ),
  ...TR_VAL.map((e) => containerVal(`table.row_${e.local}`, e, trPrOf, hasRow)),
  ...TC_ONOFF.map((e) =>
    containerToggle(`table.cell_${e.local}`, e, "table", tcPrOf, (m) => currentCell(m)?.tcPr),
  ),
  ...TC_VAL.map((e) => containerVal(`table.cell_${e.local}`, e, tcPrOf, hasCell)),
] as Command<never>[];

/** Element local name → the command id that edits it (for the ledger). */
export const TABLE_PROPERTY_BINDINGS: { local: string; id: string }[] = [
  ...TBL_ONOFF.map((e) => ({ local: e.local, id: `table.${e.local}` })),
  ...TBL_VAL.map((e) => ({ local: e.local, id: `table.tbl_${e.local}` })),
  ...TR_ONOFF.map((e) => ({ local: e.local, id: `table.row_${e.local}` })),
  ...TR_VAL.map((e) => ({ local: e.local, id: `table.row_${e.local}` })),
  ...TC_ONOFF.map((e) => ({ local: e.local, id: `table.cell_${e.local}` })),
  ...TC_VAL.map((e) => ({ local: e.local, id: `table.cell_${e.local}` })),
];
