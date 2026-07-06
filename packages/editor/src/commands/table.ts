/**
 * Table commands: insert a table, add/remove rows, set cell text and cell
 * vertical alignment, mark header rows, set row height, and apply table borders
 * / cell shading. Structural table ops resolve the table the caret sits in.
 */

import {
  addTable,
  appendTableRow,
  removeTableRow,
  setTableBorders,
  type TableBordersOptions,
  type TableCellShadingOptions,
  setTableCellShading,
  setTableCellText,
  setTableCellVerticalAlign,
  type TableCellVerticalAlign,
  type TableRowHeightRule,
  setTableRowAsHeader,
  setTableRowHeight,
  type WmlTable,
} from "@office-kit/docx";
import { asTable, blockAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import { caretAt } from "../selection.js";
import { caretBlockIndex, moveLastBlockAfter } from "./insert-util.js";
import type { Command } from "./types.js";

/** The table the caret is inside, or undefined. */
function currentTable(model: EditorModel): WmlTable | undefined {
  const block = model.selection?.focus.block;
  if (block === undefined) return undefined;
  return asTable(blockAt(model.doc, block));
}

export const insertTableCommand: Command<{ rows: number; cols: number }> = {
  id: "table.insert",
  group: "table",
  label: "Insert table",
  run(model, { rows, cols }) {
    const at = caretBlockIndex(model.doc, model.selection?.focus.block);
    const grid: string[][] = Array.from({ length: Math.max(rows, 1) }, () =>
      Array.from({ length: Math.max(cols, 1) }, () => ""),
    );
    const table = addTable(model.doc, grid);
    setTableBorders(table, {});
    moveLastBlockAfter(model.doc, at);
    model.setSelection(caretAt({ block: at + 1, cell: { row: 0, col: 0 } }));
  },
};

export const addRowCommand: Command<{ texts?: readonly string[] }> = {
  id: "table.addRow",
  group: "table",
  label: "Add row",
  run(model, { texts = [] }) {
    const table = currentTable(model);
    if (!table) return;
    appendTableRow(table, texts);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const deleteRowCommand: Command<{ row?: number }> = {
  id: "table.deleteRow",
  group: "table",
  label: "Delete row",
  run(model, { row }) {
    const table = currentTable(model);
    if (!table) return;
    const index = row ?? model.selection?.focus.cell?.row ?? table.rows.length - 1;
    removeTableRow(table, index);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const setCellTextCommand: Command<{ row: number; col: number; text: string }> = {
  id: "table.setCellText",
  group: "table",
  label: "Set cell text",
  run(model, { row, col, text }) {
    const table = currentTable(model);
    if (!table) return;
    setTableCellText(table, row, col, text);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const setCellVerticalAlignCommand: Command<{
  row: number;
  col: number;
  align: TableCellVerticalAlign;
}> = {
  id: "table.cellVAlign",
  group: "table",
  label: "Cell vertical alignment",
  run(model, { row, col, align }) {
    const cell = currentTable(model)?.rows[row]?.cells[col];
    if (cell) setTableCellVerticalAlign(cell, align);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const setRowHeaderCommand: Command<{ row: number; isHeader?: boolean }> = {
  id: "table.rowHeader",
  group: "table",
  label: "Header row",
  run(model, { row, isHeader = true }) {
    const r = currentTable(model)?.rows[row];
    if (r) setTableRowAsHeader(r, isHeader);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const setRowHeightCommand: Command<{
  row: number;
  heightTwips: number;
  rule?: TableRowHeightRule;
}> = {
  id: "table.rowHeight",
  group: "table",
  label: "Row height",
  run(model, { row, heightTwips, rule = "atLeast" }) {
    const r = currentTable(model)?.rows[row];
    if (r) setTableRowHeight(r, heightTwips, rule);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const setTableBordersCommand: Command<TableBordersOptions> = {
  id: "table.borders",
  group: "table",
  label: "Table borders",
  run(model, options) {
    const table = currentTable(model);
    if (table) setTableBorders(table, options);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const setCellShadingCommand: Command<{
  row: number;
  col: number;
  shading: TableCellShadingOptions;
}> = {
  id: "table.cellShading",
  group: "table",
  label: "Cell shading",
  run(model, { row, col, shading }) {
    const cell = currentTable(model)?.rows[row]?.cells[col];
    if (cell) setTableCellShading(cell, shading);
  },
  isEnabled: (model) => !!currentTable(model),
};

export const tableCommands = [
  insertTableCommand,
  addRowCommand,
  deleteRowCommand,
  setCellTextCommand,
  setCellVerticalAlignCommand,
  setRowHeaderCommand,
  setRowHeightCommand,
  setTableBordersCommand,
  setCellShadingCommand,
];
