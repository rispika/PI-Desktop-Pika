import type { ComponentProps } from "react";
import type { Workbook } from "@fortune-sheet/react";
import type { OfficePreview, SpreadsheetCellPatch } from "@pi-desktop/shared";

export type Spreadsheet = Extract<OfficePreview, { format: "spreadsheet" }>;
export type GridSheets = ComponentProps<typeof Workbook>["data"];

export function toGrid(preview: Spreadsheet): GridSheets {
  return preview.sheets.map((sheet, index) => {
    const cells = sheet.cells ?? sheet.rows.flatMap((row, r) => row.map((value, c) => ({ row: r, column: c, value })));
    const merge: NonNullable<GridSheets[number]["config"]>["merge"] = {};
    const celldata: NonNullable<GridSheets[number]["celldata"]> = cells.map(cell => ({ r: cell.row, c: cell.column, v: { v: cell.value, m: sheet.rows[cell.row]?.[cell.column] ?? String(cell.value), ct: { fa: "General", t: typeof cell.value === "string" ? "s" : "n" } } }));
    for (const item of sheet.merges ?? []) {
      merge[`${item.row}_${item.column}`] = { r: item.row, c: item.column, rs: item.rows, cs: item.columns };
      for (let r = item.row; r < item.row + item.rows; r++) for (let c = item.column; c < item.column + item.columns; c++) {
        let cell = celldata.find(cell => cell.r === r && cell.c === c);
        if (!cell) { cell = { r, c, v: {} }; celldata.push(cell); }
        cell.v = { ...cell.v, mc: r === item.row && c === item.column ? merge[`${item.row}_${item.column}`] : { r: item.row, c: item.column } };
      }
    }
    return { id: String(index), name: sheet.name, status: index === preview.sheets.findIndex(s => !s.hidden) ? 1 : 0, hide: sheet.hidden ? 1 : 0,
      row: 250, column: 60, addRows: 0, celldata, config: { columnlen: sheet.columnWidths, rowlen: sheet.rowHeights, merge } };
  });
}

export function canEditCell(preview: Spreadsheet, sheet: number, row: number, column: number): boolean {
  const data = preview.sheets[sheet];
  if (!data || data.protected || data.hidden || row < 0 || row >= 250 || column < 0 || column >= 60) return false;
  if (data.cells?.some(cell => cell.row === row && cell.column === column && cell.formula !== undefined)) return false;
  return !(data.merges ?? []).some(m => row >= m.row && row < m.row + m.rows && column >= m.column && column < m.column + m.columns && (row !== m.row || column !== m.column));
}

export function gridPatches(preview: Spreadsheet, sheets: GridSheets): SpreadsheetCellPatch[] {
  if (sheets.length !== preview.sheets.length) throw new Error("valuesOnly");
  const patches: SpreadsheetCellPatch[] = [];
  for (let index = 0; index < sheets.length; index++) {
    const sheet = sheets[index], source = preview.sheets[index];
    if (sheet.id !== String(index) || sheet.name !== source.name) throw new Error("valuesOnly");
    const originals = new Map((source.cells ?? source.rows.flatMap((row, r) => row.map((value, c) => ({ row: r, column: c, value })))).map(cell => [`${cell.row}:${cell.column}`, cell.value]));
    const data = sheet.data ? sheet.data.flatMap((row, r) => row.map((v, c) => ({ r, c, v }))) : sheet.celldata ?? [];
    const current = new Map(data.map(cell => [`${cell.r}:${cell.c}`, cell]));
    for (const key of new Set([...originals.keys(), ...current.keys()])) {
      const cell = current.get(key);
      const [row, column] = key.split(":").map(Number);
      const value = cell?.v?.v ?? null;
      if (cell?.v?.f || Object.keys(cell?.v ?? {}).some(k => !["v", "m", "ct", "mc"].includes(k))) throw new Error("valuesOnly");
      const before = originals.get(key) ?? null;
      if (value === before || ((value === "" || value === null) && (before === "" || before === null))) continue;
      if (!canEditCell(preview, index, row, column) || (typeof value === "string" && (value.startsWith("=") || value.length > 4000))) throw new Error("valuesOnly");
      patches.push({ sheet: index, row, column, value });
    }
  }
  return patches;
}
