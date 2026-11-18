import { describe, expect, it } from "vitest";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";
import { isOfficePreviewPath, previewOfficeFile } from "./office-preview.js";
import { previewFile } from "./workspace-files.js";

const CONTENT_TYPES =
  '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';

async function archive(files: Record<string, string | Buffer>): Promise<Buffer> {
  const zip = new JSZip();
  for (const [path, content] of Object.entries(files)) zip.file(path, content);
  return zip.generateAsync({ type: "nodebuffer" });
}

function packageRelationship(target: string): string {
  return `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="${target}"/></Relationships>`;
}

describe("Office previews", () => {
  it("detects modern Office extensions and declared attachment MIME types", () => {
    expect(isOfficePreviewPath("report.docx")).toBe(true);
    expect(isOfficePreviewPath("book.xlsx")).toBe(true);
    expect(isOfficePreviewPath("deck.pptx")).toBe(true);
    expect(
      isOfficePreviewPath(
        "attachments/" + "a".repeat(64),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      ),
    ).toBe(true);
    expect(isOfficePreviewPath("archive.zip")).toBe(false);
  });

  it("renders Word headings, formatted runs, and tables", async () => {
    const buffer = await archive({
      "[Content_Types].xml": CONTENT_TYPES,
      "_rels/.rels": packageRelationship("word/document.xml"),
      "word/document.xml": `<?xml version="1.0"?><w:document xmlns:w="urn:w"><w:body>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Quarterly report</w:t></w:r></w:p>
        <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Revenue: </w:t></w:r><w:r><w:t>growing</w:t></w:r></w:p>
        <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Total</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
      </w:body></w:document>`,
    });

    const preview = await previewOfficeFile(buffer, "report.docx");
    expect(preview?.format).toBe("word");
    if (preview?.format !== "word") return;
    expect(preview.blocks[0]).toMatchObject({
      kind: "paragraph",
      paragraph: { text: "Quarterly report", style: "heading1" },
    });
    expect(preview.blocks[1]).toMatchObject({
      kind: "paragraph",
      paragraph: { text: "Revenue: growing" },
    });
    expect(preview.blocks[2]).toEqual({
      kind: "table",
      rows: [["Region", "Total"]],
    });
  });

  it("renders workbook sheets with shared strings, booleans, and formulas", async () => {
    const buffer = await archive({
      "[Content_Types].xml": CONTENT_TYPES,
      "_rels/.rels": packageRelationship("xl/workbook.xml"),
      "xl/workbook.xml": `<?xml version="1.0"?><workbook xmlns:r="urn:r"><sheets><sheet name="Summary" r:id="rId1"/></sheets></workbook>`,
      "xl/_rels/workbook.xml.rels": `<?xml version="1.0"?><Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`,
      "xl/sharedStrings.xml": `<?xml version="1.0"?><sst><si><t>Revenue</t></si><si><t>Ready</t></si></sst>`,
      "xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet><sheetData>
        <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
        <row r="2"><c r="A2"><v>42</v></c><c r="B2" t="b"><v>1</v></c></row>
        <row r="3"><c r="A3"><f>SUM(A2:A2)</f><v>42</v></c></row>
      </sheetData></worksheet>`,
    });

    const preview = await previewOfficeFile(buffer, "figures.xlsx");
    expect(preview?.format).toBe("spreadsheet");
    if (preview?.format !== "spreadsheet") return;
    expect(preview.sheets).toEqual([
      {
        name: "Summary",
        rows: [
          ["Revenue", "Ready"],
          ["42", "TRUE"],
          ["42"],
        ],
        truncated: false,
      },
    ]);
  });

  it("renders slide order, text, dimensions, and embedded images", async () => {
    const png = Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
      "base64",
    );
    const buffer = await archive({
      "[Content_Types].xml": CONTENT_TYPES,
      "_rels/.rels": packageRelationship("ppt/presentation.xml"),
      "ppt/presentation.xml": `<?xml version="1.0"?><p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldSz cx="12192000" cy="6858000"/><p:sldIdLst><p:sldId r:id="rId2"/></p:sldIdLst></p:presentation>`,
      "ppt/_rels/presentation.xml.rels": `<?xml version="1.0"?><Relationships><Relationship Id="rId2" Target="slides/slide1.xml"/></Relationships>`,
      "ppt/slides/slide1.xml": `<?xml version="1.0"?><p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Welcome</a:t></a:r></a:p><a:p><a:r><a:t>Agenda</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:blipFill><a:blip r:embed="rId3"/></p:blipFill></p:pic></p:spTree></p:cSld></p:sld>`,
      "ppt/slides/_rels/slide1.xml.rels": `<?xml version="1.0"?><Relationships><Relationship Id="rId3" Target="../media/image1.png"/></Relationships>`,
      "ppt/media/image1.png": png,
    });

    const preview = await previewOfficeFile(buffer, "deck.pptx");
    expect(preview?.format).toBe("presentation");
    if (preview?.format !== "presentation") return;
    expect(preview.slideWidth).toBe(12_192_000);
    expect(preview.slideHeight).toBe(6_858_000);
    expect(preview.slides).toHaveLength(1);
    expect(preview.slides[0].paragraphs.map((paragraph) => paragraph.text)).toEqual([
      "Welcome",
      "Agenda",
    ]);
    expect(preview.slides[0].imageDataUrls[0]).toMatch(/^data:image\/png;base64,/);
  });

  it("decodes numeric entities without double decoding and formats numeric cells", async () => {
    const buffer = await archive({
      "[Content_Types].xml": CONTENT_TYPES,
      "xl/workbook.xml": '<workbook xmlns:r="urn:r"><sheets><sheet name="Summary" r:id="rId1"/></sheets></workbook>',
      "xl/_rels/workbook.xml.rels": '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
      "xl/styles.xml": '<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="10"/><xf numFmtId="164"/></cellXfs></styleSheet>',
      "xl/worksheets/sheet1.xml": '<worksheet><sheetData><row r="3"><c r="A3" t="inlineStr"><is><t>&#27979;&#x8BD5; &amp;#65;</t></is></c><c r="B3" s="1"><v>0.125</v></c><c r="C3" s="2"><v>45292</v></c><c r="D3" t="inlineStr"><is><t>00123</t></is></c></row></sheetData></worksheet>',
    });
    const preview = await previewOfficeFile(buffer, "formatted.xlsx");
    expect(preview?.format).toBe("spreadsheet");
    if (preview?.format !== "spreadsheet") return;
    expect(preview.sheets[0].rows).toEqual([[], [], ["\u6d4b\u8bd5 &#65;", "12.50%", "2024-01-01", "00123"]]);
  });

  it("preserves positioned slide objects and run styles", async () => {
    const buffer = await archive({
      "[Content_Types].xml": CONTENT_TYPES,
      "ppt/presentation.xml": '<p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldSz cx="1000" cy="500"/><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>',
      "ppt/_rels/presentation.xml.rels": '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>',
      "ppt/slides/slide1.xml": '<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:sp><p:spPr><a:xfrm><a:off x="100" y="50"/><a:ext cx="400" cy="200"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:rPr b="1" sz="2400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Title</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
    });
    const preview = await previewOfficeFile(buffer, "layout.pptx");
    expect(preview?.format).toBe("presentation");
    if (preview?.format !== "presentation") return;
    expect(preview.slides[0]).toMatchObject({ background: "#112233", elements: [{ x: 100, y: 50, width: 400, height: 200, paragraphs: [{ runs: [{ text: "Title", bold: true, fontSize: 24, color: "#FF0000" }] }] }] });
  });

  it("filters active Word parts and external relationships from layout packages", async () => {
    const buffer = await archive({
      "[Content_Types].xml": CONTENT_TYPES,
      "word/document.xml": '<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Safe</w:t></w:r></w:p></w:body></w:document>',
      "word/_rels/document.xml.rels": '<Relationships><Relationship Id="external" TargetMode="External" Target="https://example.invalid/pixel.png"/><Relationship Id="scheme" Target="file:///private"/><Relationship Id="safe" Target="media/local.png"/></Relationships>',
      "word/media/local.png": Buffer.from([1, 2, 3]),
      "word/vbaProject.bin": Buffer.from([1]),
      "word/embeddings/object.xml": '<object/>',
      "word/afchunk/part.html": '<script>alert(1)</script>',
    });
    const result = await previewOfficeFile(buffer, "safe.docm");
    expect(result?.format).toBe("word");
    if (result?.format !== "word") return;
    expect(result.packageBase64).toBeTypeOf("string");
    const filtered = await JSZip.loadAsync(Buffer.from(result.packageBase64!, "base64"));
    expect(Object.keys(filtered.files).filter(name => !filtered.files[name].dir)).toEqual(expect.arrayContaining(["word/document.xml", "word/media/local.png"]));
    expect(filtered.file("word/vbaProject.bin")).toBeNull();
    expect(filtered.file("word/embeddings/object.xml")).toBeNull();
    expect(filtered.file("word/afchunk/part.html")).toBeNull();
    const rels = await filtered.file("word/_rels/document.xml.rels")!.async("string");
    expect(rels).toContain('Id="safe"');
    expect(rels).not.toContain('Id="external"');
    expect(rels).not.toContain('Id="scheme"');
  });

  it("keeps text fallback but omits layout when Word block limits are exceeded", async () => {
    const buffer = await archive({
      "[Content_Types].xml": CONTENT_TYPES,
      "word/document.xml": '<w:document xmlns:w="urn:w"><w:body>' + '<w:p><w:r><w:t>Line</w:t></w:r></w:p>'.repeat(601) + '</w:body></w:document>',
    });
    const result = await previewOfficeFile(buffer, "large.docx");
    expect(result?.format).toBe("word");
    if (result?.format !== "word") return;
    expect(result.truncated).toBe(true);
    expect(result.blocks).toHaveLength(600);
    expect(result.packageBase64).toBeUndefined();
  });

  it("returns undefined for corrupt packages instead of trusting the extension", async () => {
    expect(await previewOfficeFile(Buffer.from("not a zip"), "broken.docx")).toBeUndefined();
    expect(await previewOfficeFile(Buffer.from("plain text"), "notes.txt")).toBeUndefined();
  });

  it("returns the parsed Office payload through the normal file preview path", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pi-office-preview-"));
    try {
      const path = join(directory, "report.docx");
      await writeFile(
        path,
        await archive({
          "[Content_Types].xml": CONTENT_TYPES,
          "_rels/.rels": packageRelationship("word/document.xml"),
          "word/document.xml":
            '<?xml version="1.0"?><w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Preview me</w:t></w:r></w:p></w:body></w:document>',
        }),
      );
      const result = await previewFile(path, "report.docx");
      expect(result.kind).toBe("office");
      expect(result.office).toMatchObject({
        format: "word",
        blocks: [{ kind: "paragraph", paragraph: { text: "Preview me" } }],
      });
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });
});
