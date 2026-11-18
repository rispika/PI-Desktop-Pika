import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createRequire } from "node:module";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import test from "node:test";
import { repositoryRoot, resolveElectronBinary } from "../../../scripts/e2e/boot.mjs";

const root = repositoryRoot();
const require = createRequire(join(root, "packages/agent-runtime/package.json"));
const { build } = require("esbuild");
const JSZip = createRequire(join(root, "packages/host-runtime/package.json"))("jszip");

const probe = String.raw`
import React from "react";
import { createRoot } from "react-dom/client";
import { flushSync } from "react-dom";
import i18n from "i18next";
import { initReactI18next } from "react-i18next";
import { OfficePreview } from "./apps/desktop/src/components/workpanel/OfficePreview";
await i18n.use(initReactI18next).init({ lng: "en", resources: { en: { translation: {} } }, initImmediate: false });
const root = createRoot(document.getElementById("root"));
const check = (value, message) => { if (!value) throw new Error(message); };
const waitFor = async (predicate, label) => {
  const deadline = performance.now() + 8000;
  while (!predicate()) {
    if (performance.now() > deadline) throw new Error(label);
    await new Promise(resolve => requestAnimationFrame(resolve));
  }
};
const show = (preview) => flushSync(() => root.render(<OfficePreview preview={preview} />));
window.editProbe = async (data) => {
  let saved = null;
  let fail = true;
  let release;
  const editor = { identity: "test-workbook", save: async (revision, patches) => {
    if (fail) throw new Error("SPREADSHEET_CONFLICT");
    await new Promise(resolve => { release = resolve; });
    saved = { revision, patches };
    return window.spreadsheetTest.save(revision, patches);
  } };
  const render = () => flushSync(() => root.render(<OfficePreview preview={data} spreadsheetEditor={editor} />));
  const button = key => document.querySelector('[aria-label="panel.files.sheet.' + key + '"]');
  render();
  await waitFor(() => document.querySelector(".fortune-sheet-canvas"), "Editable grid missing");
  check(button("save").disabled, "Clean workbook can save");
  flushSync(() => [...document.querySelectorAll('button')].find(b => b.textContent === 'panel.files.sheet.edit').click());
  await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
  const overlay = document.querySelector('.fortune-cell-area');
  check(overlay, "Cell overlay missing");
  const rect = overlay.getBoundingClientRect();
  window.expectCellInput = () => waitFor(() => document.activeElement?.id === 'luckysheet-rich-text-editor', "Cell editor not focused");
  window.finishEditProbe = async () => {
    await waitFor(() => document.querySelector('.office-spreadsheet-status')?.textContent === 'panel.files.sheet.unsaved' || document.querySelector('.office-spreadsheet-error'), "Input did not create draft");
    check(!document.querySelector('.office-spreadsheet-error'), "Editing rejected: " + document.querySelector('.office-spreadsheet-error')?.textContent);
    check(!button('save').disabled, "Read mode lost draft");
    show({format:'presentation',slides:[],truncated:false});
    render();
    check(!button('save').disabled, "File switch lost draft");
    button('save').click();
    await waitFor(() => document.querySelector('.office-spreadsheet-error'), "Save failure not shown");
    check(document.querySelector('.office-spreadsheet-error').textContent === 'panel.files.sheet.conflict', "Conflict not explained");
    check(!button('save').disabled, "Failed save lost draft");
    fail = false;
    button('save').click();
    await waitFor(() => release, "Save not started");
    show({format:'presentation',slides:[],truncated:false});
    render();
    check(button('save').disabled, "Pending save allowed duplicate writes after remount");
    release();
    await waitFor(() => saved && button('save').disabled && !document.querySelector('.office-spreadsheet-status').textContent, "Save did not finish");
    check(saved.patches.some(p => p.value === 'Edited value'), "Wrong value saved: " + JSON.stringify(saved));
  };
  return { x:rect.left+40,y:rect.top+10 };
};
window.officeProbe = async (data) => {
  show(data.sheet);
  await waitFor(() => document.querySelector(".fortune-sheet-canvas"), "Spreadsheet canvas missing");
  const tabs = [...document.querySelectorAll('.luckysheet-sheets-item')];
  flushSync(() => tabs[1].click());
  check(document.querySelector('.luckysheet-sheets-item-active').textContent.includes("Detail"), "Sheet switch failed");
  check(document.querySelector(".fortune-sheet-canvas").width > 100, "Spreadsheet canvas is blank");
  show(data.word);
  await waitFor(() => !document.querySelector(".office-word-layout").hidden, "Word layout never became visible");
  const frame = document.querySelector("iframe");
  check(frame.getAttribute("sandbox") === "allow-same-origin", "Word scripts are not sandboxed");
  const doc = frame.contentDocument;
  check(doc.querySelectorAll("section.docx").length >= 2, "Word page breaks were lost");
  check(doc.body.textContent.includes("Quarterly report"), "Word text missing");
  check(doc.querySelectorAll("a[href]").length === 0, "Word links are active");
  check(doc.querySelector('meta[http-equiv="Content-Security-Policy"]')?.content.includes("default-src 'none'"), "Word network policy missing: " + doc.head.innerHTML);
  let requestRejected = false;
  try { await doc.defaultView.fetch("http://127.0.0.1:9/"); } catch { requestRejected = true; }
  check(requestRejected, "Word iframe allowed network access");
  show({ ...data.word, packageBase64: "invalid-package" });
  await waitFor(() => !document.querySelector("iframe"), "Invalid Word layout did not recover");
  check(document.querySelector("article").textContent.includes("Quarterly report"), "Word text fallback missing");
  show(data.word);
  show(data.slide);
  await new Promise(resolve => requestAnimationFrame(resolve));
  check(!document.querySelector("iframe"), "Stale Word frame survived a file switch");
  const canvas = document.querySelector(".office-slide-canvas");
  const element = document.querySelector(".office-slide-element");
  check(element, "Positioned slide element missing");
  const a = canvas.getBoundingClientRect(), b = element.getBoundingClientRect();
  check(Math.abs((b.x - a.x) / a.width - 0.1) < 0.015, "Slide x position changed");
  check(element.textContent.includes("Slide title"), "Slide text missing");
  check(getComputedStyle(element.querySelector("span")).color === "rgb(255, 0, 0)", "Slide color missing");
  check(a.width > 200 && a.height > 100, "Slide canvas is blank or collapsed");
  return { pages: doc.querySelectorAll("section.docx").length, width: a.width, height: a.height };
};
`;

test("Office previews render pages, switch sheets, preserve slide positions, and recover safely", {
  timeout: 60000,
  skip: process.platform === "linux" && !process.env.DISPLAY ? "Electron test requires a display" : false,
}, async () => {
  const temp = await mkdtemp(join(process.env.PI_SCRATCH_DIR || tmpdir(), "office-renderer-"));
  try {
    const parserBundle = await build({ entryPoints: [join(root, "packages/host-runtime/src/office-preview.ts")], bundle: true, platform: "node", format: "cjs", write: false });
    await writeFile(join(temp, "parser.cjs"), parserBundle.outputFiles[0].contents);
    const { previewOfficeFile } = createRequire(import.meta.url)(join(temp, "parser.cjs"));
    const archive = async (files) => {
      const zip = new JSZip();
      zip.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>');
      for (const [name, content] of Object.entries(files)) zip.file(name, content);
      return zip.generateAsync({ type: "nodebuffer" });
    };
    const word = await previewOfficeFile(await archive({
      "_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
      "word/document.xml": '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Quarterly report</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/><w:t>Second page</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>',
      "word/_rels/document.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
      "word/styles.xml": '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>',
    }), "report.docx");
    const spreadsheetBytes = await archive({
      "xl/workbook.xml": '<workbook xmlns:r="urn:r"><sheets><sheet name="Summary" r:id="one"/><sheet name="Detail" r:id="two"/></sheets></workbook>',
      "xl/_rels/workbook.xml.rels": '<Relationships><Relationship Id="one" Target="worksheets/one.xml"/><Relationship Id="two" Target="worksheets/two.xml"/></Relationships>',
      "xl/styles.xml": '<styleSheet><cellXfs><xf numFmtId="10"/></cellXfs></styleSheet>',
      "xl/worksheets/one.xml": '<worksheet><sheetData><row r="3"><c r="A3" t="inlineStr"><is><t>00123</t></is></c><c r="B3"><v>0.125</v></c></row></sheetData></worksheet>',
      "xl/worksheets/two.xml": '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Second sheet</t></is></c></row></sheetData></worksheet>',
    });
    await writeFile(join(temp, "book.xlsx"), spreadsheetBytes);
    const sheet = await previewOfficeFile(spreadsheetBytes, "book.xlsx");
    const slide = await previewOfficeFile(await archive({
      "ppt/presentation.xml": '<p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldSz cx="9144000" cy="5143500"/><p:sldIdLst><p:sldId r:id="one"/></p:sldIdLst></p:presentation>',
      "ppt/_rels/presentation.xml.rels": '<Relationships><Relationship Id="one" Target="slides/one.xml"/></Relationships>',
      "ppt/slides/one.xml": '<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:spPr><a:xfrm><a:off x="914400" y="514350"/><a:ext cx="7000000" cy="1500000"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Slide title</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
    }), "deck.pptx");
    const rendererBundle = await build({
      stdin: { contents: probe, resolveDir: root, loader: "tsx" }, outfile: join(temp, "renderer.js"), bundle: true, platform: "browser", format: "esm", jsx: "automatic",
      define: { "process.env.NODE_ENV": '"production"' },
      nodePaths: [join(root, "apps/desktop/node_modules")],
      alias: Object.fromEntries(["react", "react-dom", "react-i18next", "i18next"].map(name => [name, join(root, "apps/desktop/node_modules", name)])),
      write: false,
    });
    for (const output of rendererBundle.outputFiles) await writeFile(output.path, output.contents);
    const css = await readFile(join(root, "apps/desktop/src/styles/work-panel.css"), "utf8");
    await writeFile(join(temp, "index.html"), '<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="renderer.css"><style>body{margin:0;--ds-bg-dock:#fff;--ds-text-primary:#222;--ds-border-subtle:#ccc}#root{height:750px}' + css + '</style></head><body><div id="root"></div><script type="module" src="renderer.js"></script></body></html>');
    await writeFile(join(temp, "data.json"), JSON.stringify({ word, sheet, slide }));
    const saveBundle = await build({ entryPoints: [join(root, "packages/host-runtime/src/spreadsheet-edit.ts")], bundle: true, platform: "node", format: "cjs", write: false });
    await writeFile(join(temp, "save.cjs"), saveBundle.outputFiles[0].contents);
    await writeFile(join(temp, "preload.cjs"), `
const { contextBridge, ipcRenderer } = require("electron");
contextBridge.exposeInMainWorld("spreadsheetTest", { save: (revision, patches) => ipcRenderer.invoke("test:saveSpreadsheet", revision, patches) });
`);
    await writeFile(join(temp, "main.cjs"), `
const { app, BrowserWindow, ipcMain } = require("electron");
const fs = require("node:fs"), path = require("node:path");
const { saveSpreadsheet } = require("./save.cjs");
ipcMain.handle("test:saveSpreadsheet", (_, revision, patches) => saveSpreadsheet(__dirname, "book.xlsx", revision, patches));
app.setPath("userData", path.join(__dirname, "profile"));
app.whenReady().then(async () => {
  const win = new BrowserWindow({show:false,width:1100,height:800,webPreferences:{sandbox:true,contextIsolation:true,nodeIntegration:false,backgroundThrottling:false,preload:path.join(__dirname,"preload.cjs")}});
  win.webContents.on("console-message", event => console.error(event.message));
  try {
    await win.loadFile(path.join(__dirname,"index.html"));
    const data = fs.readFileSync(path.join(__dirname,"data.json"),"utf8");
    const desktop = await win.webContents.executeJavaScript("window.officeProbe("+data+")");
    fs.writeFileSync(path.join(__dirname,"desktop.png"),(await win.webContents.capturePage()).toPNG());
    const point = await win.webContents.executeJavaScript("window.editProbe(JSON.parse("+JSON.stringify(data)+").sheet)");
    win.webContents.focus();
    win.webContents.sendInputEvent({type:"mouseDown",x:Math.round(point.x),y:Math.round(point.y),button:"left",clickCount:1});
    win.webContents.sendInputEvent({type:"mouseUp",x:Math.round(point.x),y:Math.round(point.y),button:"left",clickCount:1});
    await win.webContents.executeJavaScript("new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))");
    win.webContents.sendInputEvent({type:"mouseDown",x:Math.round(point.x),y:Math.round(point.y),button:"left",clickCount:2});
    win.webContents.sendInputEvent({type:"mouseUp",x:Math.round(point.x),y:Math.round(point.y),button:"left",clickCount:2});
    await win.webContents.executeJavaScript("window.expectCellInput()");
    win.webContents.sendInputEvent({type:"keyDown",keyCode:"A",modifiers:["control"]});
    win.webContents.sendInputEvent({type:"keyUp",keyCode:"A",modifiers:["control"]});
    await win.webContents.insertText("Edited value");
    const readPoint = await win.webContents.executeJavaScript("(()=>{const b=[...document.querySelectorAll('button')].find(b=>b.textContent==='panel.files.sheet.read').getBoundingClientRect();return {x:Math.round(b.x+b.width/2),y:Math.round(b.y+b.height/2)}})()");
    win.webContents.sendInputEvent({type:"mouseDown",...readPoint,button:"left",clickCount:1});
    win.webContents.sendInputEvent({type:"mouseUp",...readPoint,button:"left",clickCount:1});
    await win.webContents.executeJavaScript("window.finishEditProbe()");
    win.setSize(420,800);
    const narrow = await win.webContents.executeJavaScript("window.officeProbe("+data+")");
    fs.writeFileSync(path.join(__dirname,"narrow.png"),(await win.webContents.capturePage()).toPNG());
    console.log("OFFICE_RENDERER " + JSON.stringify({desktop,narrow}));
    app.exit(0);
  } catch(error) { console.error(error.stack || error); app.exit(1); }
});
`);
    const env = { ...process.env };
    delete env.ELECTRON_RUN_AS_NODE;
    const child = spawn(resolveElectronBinary(root).electronBinary, [join(temp, "main.cjs")], { env, stdio: ["ignore", "pipe", "pipe"] });
    let output = "";
    for (const stream of [child.stdout, child.stderr]) stream.on("data", chunk => { output += chunk; });
    const timer = setTimeout(() => child.kill(), 45000);
    let code;
    try { code = await new Promise((resolve, reject) => { child.once("error", reject); child.once("close", resolve); }); }
    finally { clearTimeout(timer); }
    assert.equal(code, 0, output);
    const reopened = await previewOfficeFile(await readFile(join(temp, "book.xlsx")), "book.xlsx");
    assert.equal(reopened.sheets[0].rows[0][0], "Edited value");
    assert.match(output, /OFFICE_RENDERER /);
    assert.match(output, /violates the following Content Security Policy directive: "default-src 'none'"/);
  } finally {
    if (process.env.PI_KEEP_OFFICE_ARTIFACTS) console.log(`Office artifacts: ${temp}`);
    else await rm(temp, { recursive: true, force: true });
  }
});
