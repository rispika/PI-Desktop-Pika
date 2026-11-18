import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const read = (path) => readFile(new URL(path, import.meta.url), "utf8");

test("the Files panel routes parsed Office documents to the Office renderer", async () => {
  const panel = await read("../src/components/workpanel/FilesTab.tsx");
  const remote = await read("../src/components/remote/RemoteWorkspace.tsx");
  assert.match(panel, /file\.kind === "office" && file\.office/);
  assert.match(panel, /<OfficePreview preview=\{file\.office\}/);
  assert.match(remote, /file\.read\.kind === "office" && file\.read\.office/);
  assert.match(remote, /<OfficePreview preview=\{file\.read\.office\}/);
});

test("the Office renderer exposes document, spreadsheet, and slide previews", async () => {
  const component = await read("../src/components/workpanel/OfficePreview.tsx");
  const styles = await read("../src/styles/work-panel.css");
  const spreadsheet = await read("../src/components/workpanel/SpreadsheetPreview.tsx");
  assert.match(component, /className="office-word-table"/);
  assert.match(component, /className="office-slide-canvas"/);
  assert.match(spreadsheet, /<Workbook/);
  assert.match(spreadsheet, /<SegmentedControl/);
  assert.match(styles, /\.office-word-preview/);
  assert.match(styles, /\.office-spreadsheet-grid/);
  assert.match(styles, /\.office-slide-canvas/);
});
