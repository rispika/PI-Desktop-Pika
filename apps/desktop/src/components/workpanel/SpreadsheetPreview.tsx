import { useLayoutEffect, useRef, useState } from "react";
import { flushSync } from "react-dom";
import { Workbook, type WorkbookInstance } from "@fortune-sheet/react";
import "@fortune-sheet/react/dist/index.css";
import { Save, Undo2 } from "lucide-react";
import { useTranslation } from "react-i18next";
import type { SpreadsheetCellPatch } from "@pi-desktop/shared";
import { SegmentedControl, TooltipButton } from "../ui";
import { canEditCell, gridPatches, toGrid, type GridSheets, type Spreadsheet } from "./spreadsheet-model";

export type SpreadsheetEditor = {
  identity: string;
  save: (revision: string, patches: SpreadsheetCellPatch[]) => Promise<Spreadsheet>;
};
type Draft = { preview: Spreadsheet; sheets: GridSheets; patches: SpreadsheetCellPatch[] };
// Drafts survive work-panel/session switches; only save or explicit discard removes them.
const drafts = new Map<string, Draft>();
const pendingSaves = new Map<string, Promise<Spreadsheet>>();
const preventClose = (event: BeforeUnloadEvent) => { event.preventDefault(); event.returnValue = ""; };
function keepDraft(identity: string, draft?: Draft) {
  if (draft?.patches.length) drafts.set(identity, draft); else drafts.delete(identity);
  window.removeEventListener("beforeunload", preventClose);
  if (drafts.size) window.addEventListener("beforeunload", preventClose);
}

export function SpreadsheetPreview({ data, editor }: { data: Spreadsheet; editor?: SpreadsheetEditor }) {
  const { t, i18n } = useTranslation();
  const identity = editor?.identity ?? "";
  const [draft, setDraft] = useState<Draft>(() => drafts.get(identity) ?? { preview: data, sheets: toGrid(data), patches: [] });
  const [mode, setMode] = useState<"read" | "edit">("read");
  const [generation, setGeneration] = useState(0);
  const [busy, setBusy] = useState(() => pendingSaves.has(identity));
  const [error, setError] = useState<string | null>(null);
  const book = useRef<WorkbookInstance>(null);
  const baseline = useRef(data);
  const active = useRef(true);
  const saving = useRef(pendingSaves.has(identity));
  const draftRef = useRef(draft);
  draftRef.current = draft;
  const dirty = draft.patches.length > 0;
  const editable = !!editor && !!data.editable && !!draft.preview.revision && !data.truncated;
  useLayoutEffect(() => {
    active.current = true;
    const pending = pendingSaves.get(identity);
    if (pending) void pending.then(result => {
      if (!active.current) return;
      baseline.current = result;
      const next = { preview: result, sheets: toGrid(result), patches: [] };
      draftRef.current = next;
      setDraft(next);
      setGeneration(n => n + 1);
    }, cause => {
      if (active.current) setError(saveErrorKey(cause));
    }).finally(() => { if (active.current) { saving.current = false; setBusy(false); } });
    return () => { active.current = false; };
  }, [identity]);

  const commitCell = () => {
    const input = document.activeElement;
    if (mode !== "edit" || !(input instanceof HTMLElement) || !input.closest(".office-spreadsheet-grid") || !input.isContentEditable) return;
    // FortuneSheet commits its active cell through the normal Enter handler.
    flushSync(() => input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", code: "Enter", keyCode: 13, which: 13, bubbles: true })));
  };

  const change = (sheets: GridSheets) => {
    if (mode !== "edit" || saving.current) return;
    try {
      const patches = gridPatches(draftRef.current.preview, sheets);
      const next = { preview: draftRef.current.preview, sheets, patches };
      draftRef.current = next;
      setDraft(next);
      keepDraft(identity, next);
    } catch {
      setError("panel.files.sheet.valuesOnly");
      setGeneration(n => n + 1);
    }
  };
  const save = async () => {
    const current = draftRef.current;
    if (!editor || saving.current || !current.patches.length || !current.preview.revision) return;
    saving.current = true;
    setBusy(true);
    setError(null);
    try {
      const pending = editor.save(current.preview.revision, current.patches);
      pendingSaves.set(identity, pending);
      const result = await pending;
      if (drafts.get(identity) === current) keepDraft(identity);
      if (!active.current) return;
      baseline.current = result;
      const next = { preview: result, sheets: toGrid(result), patches: [] };
      draftRef.current = next;
      setDraft(next);
      setGeneration(n => n + 1);
    } catch (cause) {
      if (active.current) setError(saveErrorKey(cause));
    } finally { pendingSaves.delete(identity); saving.current = false; if (active.current) setBusy(false); }
  };
  const discard = () => {
    if (!window.confirm(t("panel.files.sheet.discardConfirm"))) return;
    keepDraft(identity);
    const next = { preview: baseline.current, sheets: toGrid(baseline.current), patches: [] };
    draftRef.current = next;
    setDraft(next);
    setError(null);
    setGeneration(n => n + 1);
  };
  return <div className="office-spreadsheet">
    <div className="office-spreadsheet-toolbar" onPointerDownCapture={commitCell}>
      <SegmentedControl value={mode} onChange={value => { setMode(value); setGeneration(n => n + 1); }} label={t("panel.files.sheet.mode")}
        options={[{ value: "read", label: t("panel.files.sheet.read") }, ...(editable ? [{ value: "edit" as const, label: t("panel.files.sheet.edit") }] : [])]} disabled={busy} />
      <span className="office-spreadsheet-status" role="status">{busy ? t("panel.files.sheet.saving") : dirty ? t("panel.files.sheet.unsaved") : ""}</span>
      {editable && <>
        <TooltipButton tooltip={t("panel.files.sheet.discard")} ariaLabel={t("panel.files.sheet.discard")} disabled={!dirty || busy} onClick={discard}><Undo2 size={15} /></TooltipButton>
        <TooltipButton tooltip={t("panel.files.sheet.save")} ariaLabel={t("panel.files.sheet.save")} disabled={(!dirty && mode !== "edit") || busy} onClick={() => void save()}><Save size={15} /></TooltipButton>
      </>}
    </div>
    {error && <div className="office-spreadsheet-error" role="alert">{t(error)}</div>}
    {mode === "edit" && data.sheets.some(sheet => sheet.cells?.some(cell => cell.formula !== undefined)) && <div className="office-preview-cap">{t("panel.files.sheet.cachedFormulas")}</div>}
    {data.truncated && <div className="office-preview-cap">{t("panel.files.previewTruncated")}</div>}
    <div className="office-spreadsheet-grid" onKeyDownCapture={event => {
      if ((event.ctrlKey || event.metaKey) && ["b", "i", "u"].includes(event.key.toLowerCase())) { event.preventDefault(); event.stopPropagation(); }
    }}>
      <Workbook key={`${generation}:${mode}:${busy}`} ref={book} data={draft.sheets} allowEdit={mode === "edit" && editable && !busy} lang={i18n.language.startsWith("zh") ? "zh" : "en"}
        showToolbar={false} showFormulaBar showSheetTabs cellContextMenu={[]} headerContextMenu={[]} sheetTabContextMenu={[]}
        onChange={change} hooks={{
          beforeAddSheet: () => false, beforeDeleteSheet: () => false, beforeUpdateSheetName: () => false,
          beforeUpdateCell: (row, column, value: unknown) => {
            const id = Number(book.current?.getSheet()?.id ?? 0);
            const formula = (typeof value === "string" && value.startsWith("=")) || (typeof value === "object" && value !== null && "f" in value);
            const allowed = mode === "edit" && !saving.current && canEditCell(draftRef.current.preview, id, row, column) && !formula;
            if (!allowed) setError("panel.files.sheet.valuesOnly");
            return allowed;
          },
        }} />
    </div>
  </div>;
}

function saveErrorKey(cause: unknown): string {
  return cause instanceof Error && cause.message.includes("SPREADSHEET_CONFLICT")
    ? "panel.files.sheet.conflict"
    : "panel.files.sheet.saveFailed";
}
