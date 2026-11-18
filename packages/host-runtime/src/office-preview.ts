import { createHash } from "node:crypto";
import { XMLBuilder } from "fast-xml-parser";
import { parseSpreadsheet } from "./office-spreadsheet.ts";
import { newBudget, takeBudget, nodeName, nodeChildren, nodeAttr, directChildren, descendants, firstDescendant, textContent, parseXml, parseRelationships, resolvePartPath, relationshipPartPath, readArchiveBytes, readArchiveText, type XmlNode, type PreviewBudget } from "./office-xml.ts";
import JSZip, { type JSZipObject } from "jszip";
import type {
  OfficeParagraph,
  OfficePreview,
  OfficeSlide,
  OfficeTextRun,
  OfficeWordBlock,
} from "@pi-desktop/shared";

export const MAX_OFFICE_BYTES = 25 * 1024 * 1024;
const MAX_ARCHIVE_ENTRIES = 10_000;
const MAX_UNCOMPRESSED_BYTES = 80 * 1024 * 1024;
const MAX_WORD_BLOCKS = 600;
const MAX_SLIDE_COUNT = 150;
const MAX_SLIDE_IMAGES = 12;
const MAX_SLIDE_IMAGE_BYTES = 2 * 1024 * 1024;
const MAX_SLIDE_IMAGES_TOTAL_BYTES = 10 * 1024 * 1024;

type OfficeFormat = OfficePreview["format"];
const OFFICE_MIME: Record<string, OfficeFormat> = {
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "word",
  "application/vnd.ms-word.document.macroenabled.12": "word",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "spreadsheet",
  "application/vnd.ms-excel.sheet.macroenabled.12": "spreadsheet",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation": "presentation",
  "application/vnd.ms-powerpoint.presentation.macroenabled.12": "presentation",
};

const IMAGE_MIME: Record<string, string> = {
  bmp: "image/bmp",
  gif: "image/gif",
  jpeg: "image/jpeg",
  jpg: "image/jpeg",
  png: "image/png",
  webp: "image/webp",
};

function officeFormatFor(displayPath: string, mimeType?: string): OfficeFormat | undefined {
  const base = displayPath.replace(/\\/g, "/").split("/").pop() ?? "";
  const dot = base.lastIndexOf(".");
  const extension = dot > 0 ? base.slice(dot + 1).toLowerCase() : "";
  if (extension === "docx" || extension === "docm") return "word";
  if (extension === "xlsx" || extension === "xlsm") return "spreadsheet";
  if (extension === "pptx" || extension === "pptm") return "presentation";
  if (extension) return undefined;
  return OFFICE_MIME[String(mimeType ?? "").trim().toLowerCase()];
}

export function isOfficePreviewPath(displayPath: string, mimeType?: string): boolean {
  return officeFormatFor(displayPath, mimeType) !== undefined;
}

function archiveEntrySize(entry: JSZipObject): number {
  const compressed = entry as unknown as { _data?: { uncompressedSize?: number } };
  const size = compressed._data?.uncompressedSize;
  return typeof size === "number" && Number.isFinite(size) ? size : 0;
}

async function openArchive(buffer: Buffer): Promise<JSZip | undefined> {
  try {
    const zip = await JSZip.loadAsync(buffer, { checkCRC32: false });
    const entries = Object.values(zip.files);
    if (entries.length > MAX_ARCHIVE_ENTRIES) return undefined;
    if (entries.reduce((total, entry) => total + archiveEntrySize(entry), 0) > MAX_UNCOMPRESSED_BYTES) {
      return undefined;
    }
    if (!zip.file("[Content_Types].xml")) return undefined;
    return zip;
  } catch {
    return undefined;
  }
}

function paragraphStyle(paragraph: XmlNode): OfficeParagraph["style"] {
  const properties = directChildren(nodeChildren(paragraph), "w:pPr")[0];
  const style = firstDescendant(nodeChildren(properties), "w:pStyle");
  const value = nodeAttr(style, "w:val")?.toLowerCase().replace(/[\s_-]/g, "") ?? "";
  if (value === "title") return "title";
  if (/^heading1$/.test(value) || value === "1") return "heading1";
  if (/^heading2$/.test(value) || value === "2") return "heading2";
  if (/^heading3$/.test(value) || value === "3") return "heading3";
  return undefined;
}

function wordParagraph(paragraph: XmlNode, budget: PreviewBudget): OfficeParagraph {
  const runs: OfficeTextRun[] = [];
  const collect = (nodes: readonly unknown[]) => {
    for (const node of nodes) {
      const name = nodeName(node);
      if (name === "w:t") {
        const text = textContent([node]);
        if (text) runs.push({ text });
        continue;
      }
      if (name === "w:tab") {
        runs.push({ text: "\t" });
        continue;
      }
      if (name === "w:br" || name === "w:cr") {
        runs.push({ text: "\n" });
        continue;
      }
      if (name === "w:r") {
        const properties = directChildren(nodeChildren(node), "w:rPr")[0];
        const run: OfficeTextRun = {
          text: "",
          bold: Boolean(firstDescendant(nodeChildren(properties), "w:b")),
          italic: Boolean(firstDescendant(nodeChildren(properties), "w:i")),
          underline: Boolean(firstDescendant(nodeChildren(properties), "w:u")),
        };
        const before = runs.length;
        collect(nodeChildren(node));
        for (let index = before; index < runs.length; index += 1) {
          runs[index] = { ...runs[index], ...run, text: runs[index].text };
        }
        continue;
      }
      collect(nodeChildren(node));
    }
  };
  collect(nodeChildren(paragraph));
  const fullText = runs.map((run) => run.text).join("");
  const text = takeBudget(fullText, budget);
  return {
    text,
    runs: text === fullText ? runs : [{ text }],
    style: paragraphStyle(paragraph),
  };
}

function wordTable(table: XmlNode, budget: PreviewBudget): string[][] {
  const rows: string[][] = [];
  for (const row of directChildren(nodeChildren(table), "w:tr")) {
    const values: string[] = [];
    for (const cell of directChildren(nodeChildren(row), "w:tc")) {
      const paragraphs = descendants(nodeChildren(cell), "w:p").map((paragraph) => {
        const text = textContent(nodeChildren(paragraph));
        return takeBudget(text, budget);
      });
      values.push(paragraphs.join("\n"));
    }
    rows.push(values);
  }
  return rows;
}

async function parseWord(zip: JSZip, xml: string): Promise<OfficePreview> {
  const budget = newBudget();
  const root = parseXml(xml);
  const body = firstDescendant(root, "w:body");
  const blocks: OfficeWordBlock[] = [];
  let structurallyTruncated = false;

  const visit = (nodes: readonly unknown[]) => {
    for (const node of nodes) {
      if (blocks.length >= MAX_WORD_BLOCKS || budget.remaining <= 0) {
        structurallyTruncated = true;
        return;
      }
      const name = nodeName(node);
      if (name === "w:p") {
        blocks.push({ kind: "paragraph", paragraph: wordParagraph(node as XmlNode, budget) });
      } else if (name === "w:tbl") {
        blocks.push({ kind: "table", rows: wordTable(node as XmlNode, budget) });
      } else {
        visit(nodeChildren(node));
      }
    }
  };
  visit(nodeChildren(body));

  return {
    format: "word",
    packageBase64: budget.truncated || structurallyTruncated ? undefined : await wordRenderPackage(zip),
    blocks,
    truncated: budget.truncated || structurallyTruncated,
  };
}

function rgbColor(nodes: readonly unknown[]): string | undefined {
  const color = nodeAttr(firstDescendant(nodes, "a:srgbClr"), "val");
  return color && /^[0-9a-f]{6}$/i.test(color) ? `#${color}` : undefined;
}

function presentationParagraph(paragraph: XmlNode, budget: PreviewBudget): OfficeParagraph {
  const runs: OfficeTextRun[] = [];
  for (const run of nodeChildren(paragraph)) {
    if (nodeName(run) === "a:br") { runs.push({ text: "\n" }); continue; }
    if (!["a:r", "a:fld"].includes(nodeName(run) ?? "")) continue;
    const properties = firstDescendant(nodeChildren(run), "a:rPr");
    const fontSize = Number(nodeAttr(properties, "sz")) / 100;
    const text = takeBudget(descendants(nodeChildren(run), "a:t").map(node => textContent(nodeChildren(node))).join(""), budget);
    runs.push({ text, bold: nodeAttr(properties, "b") === "1", italic: nodeAttr(properties, "i") === "1", underline: Boolean(nodeAttr(properties, "u") && nodeAttr(properties, "u") !== "none"), fontSize: fontSize > 0 && fontSize <= 400 ? fontSize : undefined, color: rgbColor(nodeChildren(properties)) });
  }
  const align = nodeAttr(firstDescendant(nodeChildren(paragraph), "a:pPr"), "algn");
  return { text: runs.map(run => run.text).join(""), runs, align: align === "ctr" ? "center" : align === "r" ? "right" : "left" };
}

async function wordRenderPackage(zip: JSZip): Promise<string | undefined> {
  // Only pass bounded passive Word parts. Macros, OLE, HTML and external relationships never reach the renderer.
  const output = new JSZip();
  let total = 0;
  for (const entry of Object.values(zip.files)) {
    if (entry.dir || (!entry.name.startsWith("word/") && entry.name !== "_rels/.rels" && entry.name !== "[Content_Types].xml")) continue;
    if (!/\.(xml|rels|png|jpe?g|gif|webp|bmp)$/i.test(entry.name)) continue;
    if (/^word\/(embeddings|activeX|afchunk)\//i.test(entry.name)) continue;
    const bytes = await readArchiveBytes(zip, entry.name, 8 * 1024 * 1024);
    if (!bytes) return undefined;
    total += bytes.byteLength;
    if (total > 16 * 1024 * 1024) return undefined;
    if (entry.name.endsWith(".rels")) {
      const nodes = parseXml(new TextDecoder().decode(bytes));
      for (const node of descendants(nodes, "Relationships")) {
        node.Relationships = nodeChildren(node).filter(rel => nodeName(rel) !== "Relationship" || (!nodeAttr(rel, "TargetMode") && !/^[a-z][a-z0-9+.-]*:|^\/\//i.test(nodeAttr(rel, "Target") ?? "")));
      }
      output.file(entry.name, new XMLBuilder({ preserveOrder: true, ignoreAttributes: false }).build(nodes));
    } else output.file(entry.name, bytes);
  }
  const bytes = await output.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  return bytes.byteLength <= 8 * 1024 * 1024 ? bytes.toString("base64") : undefined;
}

function imageMime(path: string): string | undefined {
  const extension = path.split(".").pop()?.toLowerCase() ?? "";
  return IMAGE_MIME[extension];
}

async function slideImages(
  zip: JSZip,
  slidePath: string,
  slideRoot: readonly unknown[],
  state: { count: number; bytes: number },
  cache: Map<string, string>,
): Promise<string[]> {
  const relationships = parseRelationships(
    await readArchiveText(zip, relationshipPartPath(slidePath)),
  );
  const images: string[] = [];
  const seen = new Set<string>();
  for (const blip of descendants(slideRoot, "a:blip")) {
    const relationshipId = nodeAttr(blip, "r:embed");
    const target = relationshipId ? relationships.get(relationshipId) : undefined;
    if (!target) continue;
    const imagePath = resolvePartPath(slidePath, target);
    if (seen.has(imagePath)) continue;
    seen.add(imagePath);
    const cached = cache.get(imagePath);
    if (cached) { images.push(cached); continue; }
    const mime = imageMime(imagePath);
    if (!mime || state.count >= MAX_SLIDE_IMAGES) continue;
    const bytes = await readArchiveBytes(zip, imagePath, MAX_SLIDE_IMAGE_BYTES);
    if (!bytes || state.bytes + bytes.byteLength > MAX_SLIDE_IMAGES_TOTAL_BYTES) continue;
    state.count += 1;
    state.bytes += bytes.byteLength;
    const dataUrl = `data:${mime};base64,${Buffer.from(bytes).toString("base64")}`;
    cache.set(imagePath, dataUrl);
    images.push(dataUrl);
  }
  return images;
}

function parsePresentation(
  zip: JSZip,
  presentationXml: string,
  presentationPath: string,
): Promise<OfficePreview> {
  return (async () => {
    const budget = newBudget();
    const root = parseXml(presentationXml);
    const relationships = parseRelationships(
      await readArchiveText(zip, relationshipPartPath(presentationPath)),
    );
    const slideIds = descendants(root, "p:sldId");
    const slides: OfficeSlide[] = [];
    const imageState = { count: 0, bytes: 0 };
    const imageCache = new Map<string, string>();
    let truncated = slideIds.length > MAX_SLIDE_COUNT;

    for (const slideId of slideIds.slice(0, MAX_SLIDE_COUNT)) {
      const relationshipId = nodeAttr(slideId, "r:id");
      const target = relationshipId ? relationships.get(relationshipId) : undefined;
      if (!target) {
        truncated = true;
        continue;
      }
      const slidePath = resolvePartPath(presentationPath, target);
      const slideXml = await readArchiveText(zip, slidePath);
      if (!slideXml) {
        truncated = true;
        continue;
      }
      const slideRoot = parseXml(slideXml);
      const paragraphNodes = descendants(slideRoot, "a:p");
      const parsedParagraphs = new Map(paragraphNodes.map(paragraph => [paragraph, presentationParagraph(paragraph, budget)]));
      const paragraphs = [...parsedParagraphs.values()].filter(paragraph => paragraph.text.trim().length > 0);
      const imageDataUrls = await slideImages(zip, slidePath, slideRoot, imageState, imageCache);
      const layoutBudget = budget;
      const elements: NonNullable<OfficeSlide["elements"]> = [];
      const tree = firstDescendant(slideRoot, "p:spTree");
      for (const shape of nodeChildren(tree)) {
        const kind = nodeName(shape);
        if (kind !== "p:sp" && kind !== "p:pic") continue;
        const transform = firstDescendant(nodeChildren(shape), "a:xfrm");
        const offset = firstDescendant(nodeChildren(transform), "a:off");
        const extent = firstDescendant(nodeChildren(transform), "a:ext");
        const x = Number(nodeAttr(offset, "x")), y = Number(nodeAttr(offset, "y"));
        const width = Number(nodeAttr(extent, "cx")), height = Number(nodeAttr(extent, "cy"));
        if (![x, y, width, height].every(Number.isFinite) || width <= 0 || height <= 0) continue;
        const shapeParagraphs = descendants(nodeChildren(shape), "a:p").map(paragraph => parsedParagraphs.get(paragraph) ?? presentationParagraph(paragraph, layoutBudget));
        const shapeImages = kind === "p:pic" ? await slideImages(zip, slidePath, [shape], imageState, imageCache) : [];
        elements.push({ x, y, width, height, rotation: (Number(nodeAttr(transform, "rot")) || 0) / 60000, paragraphs: shapeParagraphs, imageDataUrl: shapeImages[0], background: rgbColor(nodeChildren(directChildren(nodeChildren(directChildren(nodeChildren(shape), "p:spPr")[0]), "a:solidFill")[0])) });
      }
      slides.push({ number: slides.length + 1, paragraphs, imageDataUrls, elements, background: rgbColor(nodeChildren(firstDescendant(slideRoot, "p:bg"))) });
      if (budget.remaining <= 0) {
        truncated = true;
        break;
      }
    }

    const size = firstDescendant(root, "p:sldSz");
    return {
      format: "presentation",
      slides,
      slideWidth: Number(nodeAttr(size, "cx")) || 12_192_000,
      slideHeight: Number(nodeAttr(size, "cy")) || 6_858_000,
      truncated: truncated || budget.truncated || imageState.count >= MAX_SLIDE_IMAGES,
    };
  })();
}

export async function previewOfficeFile(
  buffer: Buffer,
  displayPath: string,
  mimeType?: string,
): Promise<OfficePreview | undefined> {
  const format = officeFormatFor(displayPath, mimeType);
  if (!format || buffer.byteLength > MAX_OFFICE_BYTES) return undefined;
  const zip = await openArchive(buffer);
  if (!zip) return undefined;

  const packageRelationships = parseRelationships(
    await readArchiveText(zip, "_rels/.rels"),
  );
  const officeTarget =
    [...packageRelationships.values()].find((target) =>
      /(?:document|workbook|presentation)\.xml$/i.test(target),
    ) ??
    (format === "word"
      ? "word/document.xml"
      : format === "spreadsheet"
        ? "xl/workbook.xml"
        : "ppt/presentation.xml");
  const officePath = resolvePartPath("package.xml", officeTarget);
  const officeXml = await readArchiveText(zip, officePath);
  if (!officeXml) return undefined;

  if (format === "word") return parseWord(zip, officeXml);
  if (format === "spreadsheet") {
    const result = await parseSpreadsheet(zip, officeXml, officePath);
    if (result.format === "spreadsheet") {
      result.revision = createHash("sha256").update(buffer).digest("hex");
      result.editable = /\.xlsx$/i.test(displayPath) && !result.truncated && !result.sheets.some(sheet => sheet.protected) &&
        !firstDescendant(parseXml(officeXml), "workbookProtection") &&
        !Object.keys(zip.files).some(name => /(?:^_xmlsignatures\/|vbaProject\.bin$)/i.test(name));
    }
    return result;
  }
  return parsePresentation(zip, officeXml, officePath);
}
