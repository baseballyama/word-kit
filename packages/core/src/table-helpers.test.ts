import { getPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import { appendTableRow, getTableCellText, removeTableRow, setTableCellText } from "./index.js";

import { addTable, createDocx, openDocx, tables, toUint8Array } from "./docx.js";
describe("table cell helpers", () => {
  it("setTableCellText replaces a cell's text", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [
      ["A", "B"],
      ["1", "2"],
    ]);
    const table = tables(doc)[0];
    expect(table).toBeDefined();
    if (!table) return;
    setTableCellText(table, 1, 1, "rewritten");
    expect(getTableCellText(table, 1, 1)).toBe("rewritten");
  });

  it("setTableCellText supports run formatting", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["x"]]);
    const table = tables(doc)[0];
    if (!table) return;
    setTableCellText(table, 0, 0, "bold", { bold: true, color: "FF0000" });
    const part = toUint8Array(doc);
    const reopened = openDocx(part);
    const xml = new TextDecoder().decode(
      getPart(reopened.opc, "/word/document.xml")?.data ?? new Uint8Array(),
    );
    expect(xml).toContain("<w:b/>");
    expect(xml).toContain('w:val="FF0000"');
  });

  it("setTableCellText throws when out of range", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["x"]]);
    const table = tables(doc)[0];
    if (!table) return;
    expect(() => setTableCellText(table, 5, 0, "y")).toThrow();
    expect(() => setTableCellText(table, 0, 5, "y")).toThrow();
  });

  it("appendTableRow adds a row matching the existing column count", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [
      ["A", "B", "C"],
      ["1", "2", "3"],
    ]);
    const table = tables(doc)[0];
    if (!table) return;
    appendTableRow(table, ["x", "y", "z"]);
    expect(table.rows).toHaveLength(3);
    expect(getTableCellText(table, 2, 1)).toBe("y");
  });

  it("appendTableRow pads missing values with empty cells", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["A", "B"]]);
    const table = tables(doc)[0];
    if (!table) return;
    appendTableRow(table, ["only-one"]);
    expect(table.rows[1]?.cells).toHaveLength(2);
    expect(getTableCellText(table, 1, 0)).toBe("only-one");
    expect(getTableCellText(table, 1, 1)).toBe("");
  });

  it("removeTableRow drops the row at the given index", () => {
    const doc = createDocx({ paragraphs: [] });
    addTable(doc, [["A"], ["B"], ["C"]]);
    const table = tables(doc)[0];
    if (!table) return;
    expect(removeTableRow(table, 1)).toBe(true);
    expect(table.rows).toHaveLength(2);
    expect(getTableCellText(table, 1, 0)).toBe("C");
    expect(removeTableRow(table, 99)).toBe(false);
  });
});
