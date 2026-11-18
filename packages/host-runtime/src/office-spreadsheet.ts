import type JSZip from "jszip";
import * as SSF from "ssf";
import type { OfficePreview, OfficeSheet } from "@pi-desktop/shared";
import { newBudget, takeBudget, nodeChildren, nodeAttr, directChildren, descendants, firstDescendant, textContent, parseXml, parseRelationships, resolvePartPath, relationshipPartPath, readArchiveText } from "./office-xml.ts";
const MAX_SHEET_COUNT = 20;
const MAX_SHEET_ROWS = 250;
const MAX_SHEET_COLUMNS = 60;
function columnIndex(reference: string | undefined, fallback: number): number {
  const letters = reference?.match(/^[A-Z]+/i)?.[0]?.toUpperCase();
  if (!letters) return fallback;
  let index = 0;
  for (const letter of letters) index = index * 26 + letter.charCodeAt(0) - 64;
  return index - 1;
}

export function parseSpreadsheet(zip: JSZip, workbookXml: string, workbookPath: string): Promise<OfficePreview> {
  return (async () => {
    const budget = newBudget();
    const workbookRoot = parseXml(workbookXml);
    const relationships = parseRelationships(
      await readArchiveText(zip, relationshipPartPath(workbookPath)),
    );
    const sharedRoot = zip.file("xl/sharedStrings.xml")
      ? parseXml((await readArchiveText(zip, "xl/sharedStrings.xml")) ?? "")
      : [];
    const sharedStrings = descendants(sharedRoot, "si").map((item) =>
      textContent(nodeChildren(item)),
    );
    const stylesXml = await readArchiveText(zip, "xl/styles.xml");
    const styles = stylesXml ? parseXml(stylesXml) : [];
    const customFormats = new Map(descendants(styles, "numFmt").map(node => [Number(nodeAttr(node, "numFmtId")), nodeAttr(node, "formatCode") ?? "General"]));
    const cellFormats = directChildren(nodeChildren(firstDescendant(styles, "cellXfs")), "xf").map(node => Number(nodeAttr(node, "numFmtId")) || 0);
    const date1904 = ["1", "true"].includes(nodeAttr(firstDescendant(workbookRoot, "workbookPr"), "date1904") ?? "");
    const sheets: OfficeSheet[] = [];
    let truncated = false;

    for (const sheetNode of descendants(workbookRoot, "sheet")) {
      if (sheets.length >= MAX_SHEET_COUNT) {
        truncated = true;
        break;
      }
      const name = nodeAttr(sheetNode, "name") ?? String(sheets.length + 1);
      const relationshipId = nodeAttr(sheetNode, "r:id");
      const target = relationshipId ? relationships.get(relationshipId) : undefined;
      if (!target) { truncated = true; continue; }
      const sheetPath = resolvePartPath(workbookPath, target);
      const sheetXml = await readArchiveText(zip, sheetPath);
      if (!sheetXml) {
        truncated = true;
        continue;
      }
      const sheetRoot = parseXml(sheetXml);
      const outputRows: string[][] = [];
      const cells: NonNullable<OfficeSheet["cells"]> = [];
      const columnWidths: Record<number, number> = {};
      const rowHeights: Record<number, number> = {};
      for (const col of descendants(sheetRoot, "col")) {
        const start = Number(nodeAttr(col, "min")), end = Math.min(MAX_SHEET_COLUMNS, Number(nodeAttr(col, "max")));
        const width = Number(nodeAttr(col, "width"));
        if (Number.isInteger(start) && start > 0 && Number.isFinite(width) && width > 0) {
          for (let i = start; i <= end; i++) columnWidths[i - 1] = Math.min(1000, Math.max(24, width * 7 + 5));
        }
      }
      let sheetTruncated = false;
      for (const rowNode of descendants(sheetRoot, "row")) {
        if (outputRows.length >= MAX_SHEET_ROWS) {
          sheetTruncated = true;
          break;
        }
        const declaredRow = Number(nodeAttr(rowNode, "r"));
        const rowIndex = Number.isInteger(declaredRow) && declaredRow > 0 ? declaredRow - 1 : outputRows.length;
        if (rowIndex >= MAX_SHEET_ROWS) { sheetTruncated = true; break; }
        while (outputRows.length < rowIndex) outputRows.push([]);
        const height = Number(nodeAttr(rowNode, "ht"));
        if (Number.isFinite(height) && height > 0) rowHeights[rowIndex] = Math.min(600, height * 4 / 3);
        const values: string[] = [];
        let fallbackColumn = 0;
        for (const cellNode of directChildren(nodeChildren(rowNode), "c")) {
          const index = columnIndex(nodeAttr(cellNode, "r"), fallbackColumn);
          fallbackColumn = index + 1;
          if (index >= MAX_SHEET_COLUMNS) {
            sheetTruncated = true;
            continue;
          }
          const type = nodeAttr(cellNode, "t");
          const rawValue = textContent(nodeChildren(firstDescendant(nodeChildren(cellNode), "v")));
          let value = rawValue;
          if (type === "s") {
            const sharedIndex = Number.parseInt(rawValue, 10);
            value = Number.isInteger(sharedIndex) ? sharedStrings[sharedIndex] ?? "" : "";
          } else if (type === "inlineStr") {
            value = textContent(nodeChildren(firstDescendant(nodeChildren(cellNode), "is")));
          } else if (type === "b") {
            value = rawValue === "1" ? "TRUE" : "FALSE";
          } else if (!value) {
            const formula = firstDescendant(nodeChildren(cellNode), "f");
            value = formula ? `=${textContent(nodeChildren(formula))}` : "";
          }
          const formula = firstDescendant(nodeChildren(cellNode), "f");
          const raw: string | number | boolean = type === "b" ? rawValue === "1" : ((!type || type === "n") && rawValue && Number.isFinite(Number(rawValue))) ? Number(rawValue) : value;
          cells.push({ row: rowIndex, column: index, value: typeof raw === "string" ? raw.slice(0, 4000) : raw, ...(formula ? { formula: textContent(nodeChildren(formula)) || " " } : {}) });
          if ((!type || type === "n") && rawValue && Number.isFinite(Number(rawValue))) {
            const formatId = cellFormats[Number(nodeAttr(cellNode, "s")) || 0] ?? 0;
            const code = customFormats.get(formatId) ?? SSF.get_table()[formatId] ?? "General";
            try { value = SSF.format(code, Number(rawValue), { date1904 }); }
            catch { value = rawValue; sheetTruncated = true; }
          }
          if (value.length > 4000) sheetTruncated = true;
          while (values.length < index) values.push("");
          values[index] = takeBudget(value.slice(0, 4000), budget);
        }
        while (values.length > 0 && values[values.length - 1] === "") values.pop();
        outputRows.push(values);
      }
      const merges: NonNullable<OfficeSheet["merges"]> = [];
      for (const merge of descendants(sheetRoot, "mergeCell")) {
        const ref = nodeAttr(merge, "ref")?.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
        if (!ref) { sheetTruncated = true; continue; }
        const row = Number(ref[2]) - 1, column = columnIndex(ref[1], 0);
        const endRow = Number(ref[4]), endColumn = columnIndex(ref[3], 0) + 1;
        if (row >= 0 && endRow > row && endRow <= MAX_SHEET_ROWS && column >= 0 && endColumn > column && endColumn <= MAX_SHEET_COLUMNS) merges.push({ row, column, rows: endRow - row, columns: endColumn - column });
        else sheetTruncated = true;
      }
      const protectedSheet = !!firstDescendant(sheetRoot, "sheetProtection") || descendants(sheetRoot, "f").some(formula => !!nodeAttr(formula, "ref") || ![undefined, "normal"].includes(nodeAttr(formula, "t")));
      sheets.push({ name, rows: outputRows, cells, columnWidths, rowHeights, merges, protected: protectedSheet, hidden: ["hidden", "veryHidden"].includes(nodeAttr(sheetNode, "state") ?? ""), truncated: sheetTruncated });
      truncated ||= sheetTruncated;
    }

    return {
      format: "spreadsheet",
      sheets,
      truncated: truncated || budget.truncated,
    };
  })();
}
