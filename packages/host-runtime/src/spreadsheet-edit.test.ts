import { mkdtemp, readFile, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";
import { describe, it, expect } from "vitest";
import { previewOfficeFile } from "./office-preview.ts";
import { saveSpreadsheet } from "./spreadsheet-edit.ts";

async function fixture(protectedSheet = false) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file("xl/workbook.xml", '<workbook xmlns:r="urn:r"><sheets><sheet name="Data" r:id="one"/></sheets></workbook>');
  zip.file("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="one" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file("xl/worksheets/sheet1.xml", `<worksheet><sheetData><row r="1"><c r="A1" s="2" t="inlineStr"><is><t>Original</t></is></c><c r="B1"><v>12</v></c><c r="C1"><f>B1*2</f><v>24</v></c></row></sheetData>${protectedSheet ? '<sheetProtection sheet="1"/>' : ""}<mergeCells><mergeCell ref="A2:B2"/></mergeCells></worksheet>`);
  zip.file("xl/charts/chart1.xml", "<chart>keep me</chart>");
  zip.file("xl/media/image1.png", new Uint8Array([1, 2, 3, 4]));
  return zip.generateAsync({ type: "nodebuffer" });
}

describe("spreadsheet editing", () => {
  it("opens, patches and reopens without replacing unrelated package parts", async () => {
    const dir = await mkdtemp(join(tmpdir(), "sheet-edit-"));
    try {
      const bytes = await fixture();
      await writeFile(join(dir, "book.xlsx"), bytes);
      const initial = await previewOfficeFile(bytes, "book.xlsx");
      if (initial?.format !== "spreadsheet") throw new Error("Missing spreadsheet");
      const saved = await saveSpreadsheet(dir, "book.xlsx", initial.revision!, [{ sheet: 0, row: 0, column: 0, value: "New <value> & text" }, { sheet: 0, row: 2, column: 1, value: 42 }]);
      expect(saved.format).toBe("spreadsheet");
      const result = await readFile(join(dir, "book.xlsx"));
      const zip = await JSZip.loadAsync(result);
      expect(await zip.file("xl/charts/chart1.xml")!.async("string")).toBe("<chart>keep me</chart>");
      expect(await zip.file("xl/media/image1.png")!.async("uint8array")).toEqual(new Uint8Array([1, 2, 3, 4]));
      const xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
      expect(xml).toContain('s="2"');
      expect(xml).toContain('<f>B1*2</f>');
      expect(xml).toContain('ref="A2:B2"');
      const reopened = await previewOfficeFile(result, "book.xlsx");
      if (reopened?.format !== "spreadsheet") throw new Error("Missing spreadsheet");
      expect(reopened.sheets[0].rows[0][0]).toBe("New <value> & text");
      expect(reopened.sheets[0].rows[2][1]).toBe("42");
      expect(reopened.revision).not.toBe(initial.revision);
      await expect(saveSpreadsheet(dir, "book.xlsx", initial.revision!, [{ sheet: 0, row: 0, column: 0, value: "stale" }])).rejects.toThrow("SPREADSHEET_CONFLICT");
      expect(await readFile(join(dir, "book.xlsx"))).toEqual(result);
    } finally { await rm(dir, { recursive: true, force: true }); }
  });

  it("rejects formula changes, protected sheets, invalid input and escaped paths without writing", async () => {
    const dir = await mkdtemp(join(tmpdir(), "sheet-edit-"));
    try {
      for (const protectedSheet of [false, true]) {
        const bytes = await fixture(protectedSheet);
        await writeFile(join(dir, "book.xlsx"), bytes);
        const preview = await previewOfficeFile(bytes, "book.xlsx");
        if (preview?.format !== "spreadsheet") throw new Error("Missing spreadsheet");
        for (const patch of [
          { sheet: 0, row: 0, column: protectedSheet ? 0 : 2, value: "replacement" },
          { sheet: 0, row: -1, column: 0, value: "bad" },
          { sheet: 0, row: 0, column: 0, value: "=1+1" },
        ]) await expect(saveSpreadsheet(dir, "book.xlsx", preview.revision!, [patch])).rejects.toThrow();
        await expect(saveSpreadsheet(dir, "../outside.xlsx", preview.revision!, [])).rejects.toThrow();
        expect(await readFile(join(dir, "book.xlsx"))).toEqual(bytes);
      }
    } finally { await rm(dir, { recursive: true, force: true }); }
  });
});
