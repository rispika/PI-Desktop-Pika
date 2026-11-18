import { posix } from "node:path";
import { XMLParser } from "fast-xml-parser";
import type JSZip from "jszip";
export type XmlNode = Record<string, unknown>;
const MAX_PREVIEW_TEXT_CHARS = 200_000;
const MAX_XML_BYTES = 16 * 1024 * 1024;
const XML = new XMLParser({ ignoreAttributes: false, preserveOrder: true, removeNSPrefix: false, trimValues: false, parseTagValue: false, htmlEntities: true });
function archiveEntrySize(entry: { name: string }): number {
  const size = (entry as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize;
  return typeof size === "number" && Number.isFinite(size) ? size : 0;
}
export type PreviewBudget = {
  remaining: number;
  truncated: boolean;
};

export function newBudget(): PreviewBudget {
  return { remaining: MAX_PREVIEW_TEXT_CHARS, truncated: false };
}

export function takeBudget(text: string, budget: PreviewBudget): string {
  if (budget.remaining <= 0) {
    if (text) budget.truncated = true;
    return "";
  }
  if (text.length <= budget.remaining) {
    budget.remaining -= text.length;
    return text;
  }
  const clipped = text.slice(0, budget.remaining);
  budget.remaining = 0;
  budget.truncated = true;
  return clipped;
}

export function isRecord(value: unknown): value is XmlNode {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function nodeName(node: unknown): string | undefined {
  if (!isRecord(node)) return undefined;
  return Object.keys(node).find((key) => key !== ":@" && key !== "#text");
}

export function nodeChildren(node: unknown): unknown[] {
  const name = nodeName(node);
  if (!name || !isRecord(node)) return [];
  const value = node[name];
  return Array.isArray(value) ? value : [];
}

export function nodeAttr(node: unknown, name: string): string | undefined {
  if (!isRecord(node) || !isRecord(node[":@"])) return undefined;
  const value = node[":@"][`@_${name}`];
  return typeof value === "string" ? value : undefined;
}

export function directChildren(nodes: readonly unknown[], name: string): XmlNode[] {
  return nodes.filter(
    (node): node is XmlNode => nodeName(node) === name && isRecord(node),
  );
}

export function descendants(nodes: readonly unknown[], name: string): XmlNode[] {
  const found: XmlNode[] = [];
  const visit = (items: readonly unknown[]) => {
    for (const item of items) {
      if (nodeName(item) === name && isRecord(item)) found.push(item);
      visit(nodeChildren(item));
    }
  };
  visit(nodes);
  return found;
}

export function firstDescendant(nodes: readonly unknown[], name: string): XmlNode | undefined {
  return descendants(nodes, name)[0];
}

export function textContent(nodes: readonly unknown[]): string {
  let text = "";
  for (const node of nodes) {
    if (
      isRecord(node) &&
      (typeof node["#text"] === "string" || typeof node["#text"] === "number")
    ) {
      text += String(node["#text"]);
    }
    text += textContent(nodeChildren(node));
  }
  return text;
}

export function parseXml(xml: string): unknown[] {
  const parsed: unknown = XML.parse(xml);
  return Array.isArray(parsed) ? parsed : [];
}

export function parseRelationships(xml: string | undefined): Map<string, string> {
  const relationships = new Map<string, string>();
  if (!xml) return relationships;
  for (const relationship of descendants(parseXml(xml), "Relationship")) {
    const id = nodeAttr(relationship, "Id");
    const target = nodeAttr(relationship, "Target");
    if (id && target) relationships.set(id, target);
  }
  return relationships;
}

export function resolvePartPath(source: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  return posix.normalize(posix.join(posix.dirname(source), target));
}

export function relationshipPartPath(source: string): string {
  return posix.join(posix.dirname(source), "_rels", `${posix.basename(source)}.rels`);
}

export async function readArchiveBytes(
  zip: JSZip,
  path: string,
  limit = MAX_XML_BYTES,
): Promise<Uint8Array | undefined> {
  const entry = zip.file(path);
  if (!entry || archiveEntrySize(entry) > limit) return undefined;
  const bytes = await entry.async("uint8array");
  return bytes.byteLength <= limit ? bytes : undefined;
}

export async function readArchiveText(zip: JSZip, path: string): Promise<string | undefined> {
  const bytes = await readArchiveBytes(zip, path);
  return bytes ? new TextDecoder("utf-8").decode(bytes) : undefined;
}
