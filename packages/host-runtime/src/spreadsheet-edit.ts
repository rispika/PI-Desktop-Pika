import { createHash, randomUUID } from "node:crypto";
import { readFile, lstat, open, rename, unlink } from "node:fs/promises";
import { dirname, join } from "node:path";
import JSZip from "jszip";
import { XMLBuilder } from "fast-xml-parser";
import type { SpreadsheetCellPatch, OfficePreview } from "@pi-desktop/shared";
import { MAX_OFFICE_BYTES, previewOfficeFile } from "./office-preview.ts";
import { resolveRealOpenablePath } from "./workspace-files.ts";
import { parseXml, nodeChildren, nodeAttr, nodeName, directChildren, descendants, firstDescendant, parseRelationships, relationshipPartPath, resolvePartPath, readArchiveText } from "./office-xml.ts";

const builder = new XMLBuilder({ preserveOrder: true, ignoreAttributes: false, format: false });
const digest = (bytes: Buffer) => createHash("sha256").update(bytes).digest("hex");
const pending = new Set<string>();
function columnName(column: number): string {
  return column < 26 ? String.fromCharCode(65 + column) : columnName(Math.floor(column / 26) - 1) + columnName(column % 26);
}

function validatePatches(value: unknown): asserts value is SpreadsheetCellPatch[] {
  if (!Array.isArray(value) || value.length === 0 || value.length > 15000) throw new Error("Invalid spreadsheet patches");
  const seen = new Set<string>();
  for (const item of value) {
    if (!item || typeof item !== "object") throw new Error("Invalid cell patch");
    const p = item as Record<string, unknown>;
    if (![p.sheet, p.row, p.column].every(n => typeof n === "number" && Number.isInteger(n) && n >= 0) || Number(p.sheet) >= 20 || Number(p.row) >= 250 || Number(p.column) >= 60) throw new Error("Invalid cell address");
    if (p.value !== null && typeof p.value !== "boolean" && !(typeof p.value === "number" && Number.isFinite(p.value)) && !(typeof p.value === "string" && p.value.length <= 4000 && !p.value.startsWith("=") && !/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(p.value))) throw new Error("Invalid cell value");
    const key = `${p.sheet}:${p.row}:${p.column}`;
    if (seen.has(key)) throw new Error("Duplicate cell patch");
    seen.add(key);
  }
}

/** Patch original OOXML parts; never round-trip the workbook through a lossy exporter. */
export async function saveSpreadsheet(root: string, path: string, revision: string, patches: unknown, extraRoots: readonly string[] = []): Promise<Extract<OfficePreview, { format: "spreadsheet" }>> {
  validatePatches(patches);
  if (typeof revision !== "string" || !/^[a-f0-9]{64}$/.test(revision) || typeof path !== "string" || !/\.xlsx$/i.test(path)) throw new Error("Invalid spreadsheet save request");
  const target = await resolveRealOpenablePath(path, root, extraRoots);
  if (!target) throw new Error("Spreadsheet path escapes editable roots");
  if (pending.has(target)) throw new Error("SPREADSHEET_BUSY");
  pending.add(target);
  let temp: string | undefined;
  try {
    const info = await lstat(target);
    if (!info.isFile() || info.nlink !== 1 || info.size > MAX_OFFICE_BYTES) throw new Error("Spreadsheet is not safely editable");
    const original = await readFile(target);
    if (digest(original) !== revision) throw new Error("SPREADSHEET_CONFLICT");
    const preview = await previewOfficeFile(original, path);
    if (preview?.format !== "spreadsheet" || !preview.editable) throw new Error("Spreadsheet is read-only or not fully loaded");
    const zip = await JSZip.loadAsync(original);
    if (Object.keys(zip.files).some(name => /(?:^_xmlsignatures\/|vbaProject\.bin$)/i.test(name))) throw new Error("Signed or macro-enabled spreadsheet is read-only");
    const packageRels = parseRelationships(await readArchiveText(zip, "_rels/.rels"));
    const workbookPath = resolvePartPath("package.xml", [...packageRels.values()].find(p => /workbook\.xml$/i.test(p)) ?? "xl/workbook.xml");
    const workbookXml = await readArchiveText(zip, workbookPath);
    if (!workbookXml) throw new Error("Missing workbook");
    const workbook = parseXml(workbookXml);
    if (firstDescendant(workbook, "workbookProtection")) throw new Error("Protected workbook is read-only");
    const sheets = descendants(workbook, "sheet");
    const rels = parseRelationships(await readArchiveText(zip, relationshipPartPath(workbookPath)));
    for (const sheetIndex of new Set(patches.map(p => p.sheet))) {
      const sheet = sheets[sheetIndex];
      const targetPart = rels.get(nodeAttr(sheet, "r:id") ?? "");
      if (!targetPart || preview.sheets[sheetIndex]?.protected || preview.sheets[sheetIndex]?.hidden) throw new Error("Sheet is read-only");
      const sheetPath = resolvePartPath(workbookPath, targetPart);
      const xml = await readArchiveText(zip, sheetPath);
      if (!xml) throw new Error("Missing worksheet");
      const tree = parseXml(xml);
      const sheetData = firstDescendant(tree, "sheetData");
      if (!sheetData) throw new Error("Missing sheet data");
      const rows = nodeChildren(sheetData);
      for (const patch of patches.filter(p => p.sheet === sheetIndex)) {
        for (const merge of preview.sheets[sheetIndex]?.merges ?? []) {
          if (patch.row >= merge.row && patch.row < merge.row + merge.rows && patch.column >= merge.column && patch.column < merge.column + merge.columns && (patch.row !== merge.row || patch.column !== merge.column)) throw new Error("Merged cell is read-only");
        }
        let row = directChildren(rows, "row").find(r => Number(nodeAttr(r, "r")) === patch.row + 1);
        if (!row) {
          row = { row: [], ":@": { "@_r": String(patch.row + 1) } };
          const at = rows.findIndex(r => nodeName(r) === "row" && Number(nodeAttr(r, "r")) > patch.row + 1);
          rows.splice(at < 0 ? rows.length : at, 0, row);
        }
        const cells = nodeChildren(row);
        const address = columnName(patch.column) + (patch.row + 1);
        let cell = directChildren(cells, "c").find(c => nodeAttr(c, "r") === address);
        if (cell && firstDescendant(nodeChildren(cell), "f")) throw new Error("Formula cells are read-only");
        if (!cell) {
          cell = { c: [], ":@": { "@_r": address } };
          cells.push(cell);
          cells.sort((a, b) => {
            const col = (node: unknown) => (nodeAttr(node, "r")?.match(/^[A-Z]+/)?.[0] ?? "");
            const aCol = col(a), bCol = col(b);
            return aCol.length - bCol.length || aCol.localeCompare(bCol);
          });
        }
        const attributes = cell[":@"] as Record<string, unknown>;
        delete attributes["@_t"];
        const children = nodeChildren(cell).filter(n => !["v", "is"].includes(nodeName(n) ?? ""));
        if (typeof patch.value === "string") {
          attributes["@_t"] = "inlineStr";
          children.unshift({ is: [{ t: [{ "#text": patch.value }], ":@": { "@_xml:space": "preserve" } }] });
        } else if (patch.value !== null) {
          if (typeof patch.value === "boolean") attributes["@_t"] = "b";
          children.unshift({ v: [{ "#text": typeof patch.value === "boolean" ? (patch.value ? "1" : "0") : String(patch.value) }] });
        }
        cell.c = children;
      }
      // A dimension hint may otherwise exclude newly populated cells.
      const worksheet = firstDescendant(tree, "worksheet");
      if (worksheet) worksheet.worksheet = nodeChildren(worksheet).filter(n => nodeName(n) !== "dimension");
      zip.file(sheetPath, builder.build(tree));
    }
    const workbookNode = firstDescendant(workbook, "workbook");
    if (!workbookNode) throw new Error("Invalid workbook");
    let calc = firstDescendant(workbook, "calcPr");
    if (!calc) {
      calc = { calcPr: [], ":@": {} };
      const children = nodeChildren(workbookNode);
      const afterCalculation = new Set(["oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"]);
      const index = children.findIndex(node => afterCalculation.has(nodeName(node) ?? ""));
      children.splice(index < 0 ? children.length : index, 0, calc);
    }
    calc[":@"] = { ...(calc[":@"] as Record<string, unknown>), "@_fullCalcOnLoad": "1", "@_forceFullCalc": "1" };
    zip.file(workbookPath, builder.build(workbook));
    const output = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
    if (output.length > MAX_OFFICE_BYTES) throw new Error("Saved spreadsheet exceeds size limit");
    const result = await previewOfficeFile(output, path);
    if (result?.format !== "spreadsheet") throw new Error("Saved spreadsheet failed validation");
    temp = join(dirname(target), `.pi-sheet-${randomUUID()}.tmp`);
    const handle = await open(temp, "wx", info.mode);
    try { await handle.writeFile(output); await handle.sync(); } finally { await handle.close(); }
    if (await resolveRealOpenablePath(path, root, extraRoots) !== target || digest(await readFile(target)) !== revision) throw new Error("SPREADSHEET_CONFLICT");
    await rename(temp, target);
    temp = undefined;
    return { ...result, editable: true };
  } finally {
    pending.delete(target);
    if (temp) await unlink(temp);
  }
}
