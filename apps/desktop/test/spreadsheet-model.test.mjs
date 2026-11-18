import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import test from "node:test";

const require = createRequire(new URL("../../../packages/agent-runtime/package.json", import.meta.url));
const { build } = require("esbuild");
const bundle = await build({ entryPoints: [fileURLToPath(new URL("../src/components/workpanel/spreadsheet-model.ts", import.meta.url))], bundle: true, platform: "node", format: "cjs", write: false });
const compiled = { exports: {} };
new Function("module", "exports", bundle.outputFiles[0].text)(compiled, compiled.exports);
const { toGrid, gridPatches, canEditCell } = compiled.exports;

function fixture() {
  return { format: "spreadsheet", truncated: false, revision: "test", editable: true, sheets: [
    { name: "Data", rows: [["00123", "12", "24"]], cells: [
      { row: 0, column: 0, value: "00123" }, { row: 0, column: 1, value: 12 }, { row: 0, column: 2, value: 24, formula: "B1*2" },
    ], merges: [{row:1,column:0,rows:1,columns:2}], columnWidths: {0:144}, rowHeights: {0:32}, truncated: false },
  ] };
}

test("grid preserves cell types, dimensions and merges and emits only changed values", () => {
  const preview = fixture();
  const sheets = toGrid(preview);
  assert.deepEqual(gridPatches(preview, sheets), []);
  assert.equal(sheets[0].config.columnlen[0], 144);
  assert.equal(sheets[0].config.rowlen[0], 32);
  assert.equal(sheets[0].celldata[0].v.v, "00123");
  sheets[0].celldata[1].v.v = 42;
  sheets[0].celldata.push({r:2,c:0,v:{v:true}}, {r:3,c:0,v:{v:"A < B & C\nline"}});
  assert.deepEqual(gridPatches(preview, sheets), [
    {sheet:0,row:0,column:1,value:42}, {sheet:0,row:2,column:0,value:true}, {sheet:0,row:3,column:0,value:"A < B & C\nline"},
  ]);
});

test("grid rejects formula edits, merged followers, structure and formatting changes", () => {
  const preview = fixture();
  assert.equal(canEditCell(preview, 0, 0, 2), false);
  assert.equal(canEditCell(preview, 0, 1, 1), false);
  assert.equal(canEditCell(preview, 0, 1, 0), true);
  assert.equal(canEditCell(preview, 0, 250, 0), false);
  for (const edit of [
    sheets => { sheets[0].celldata[2].v.v = 99; },
    sheets => { sheets[0].celldata[0].v.f = "1+1"; },
    sheets => { sheets[0].celldata[0].v.v = "=1+1"; },
    sheets => { sheets[0].celldata[0].v.bl = 1; },
    sheets => { sheets[0].name = "Renamed"; },
  ]) {
    const sheets = toGrid(preview);
    edit(sheets);
    assert.throws(() => gridPatches(preview, sheets));
  }
});
