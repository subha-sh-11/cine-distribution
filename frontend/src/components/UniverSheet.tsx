"use client";

import { useEffect, useRef } from "react";
import "@univerjs/presets/lib/styles/preset-sheets-core.css";

/* eslint-disable @typescript-eslint/no-explicit-any */

// How long the user must be idle (no key/pointer/edit) before a collaborator's
// update is applied to their view. Applying one moves the active cell, so we
// wait for a pause instead of interrupting active data entry.
const IDLE_APPLY_MS = 1200;

// A whitespace-only value ("" or " ") is text; it makes "text + number" formulas
// evaluate to #VALUE!. Excel-origin pastes drop such junk into empty cells.
function isBlankText(v: any): boolean {
  return typeof v === "string" && v.trim() === "";
}

// After a paste, clear any whitespace/empty-string cell in the pasted region so
// dependent "+" formulas compute 0 (shown as "-") instead of #VALUE!. The region
// is the post-paste selection, widened to the clipboard's row/col count so a
// single-cell anchor selection still covers the whole pasted block.
function cleanPastedBlanks(univerAPI: any, params: any): void {
  try {
    const wb = univerAPI.getActiveWorkbook?.();
    const ws = wb?.getActiveSheet?.();
    if (!wb || !ws) return;
    const active = wb.getActiveRange?.();
    if (!active) return;
    const sr = active.getRow?.() ?? 0;
    const sc = active.getColumn?.() ?? 0;
    let nr = active.getHeight?.() ?? 1;
    let nc = active.getWidth?.() ?? 1;
    const text = params?.text;
    if (typeof text === "string" && text) {
      const lines = text.replace(/\r/g, "").split("\n");
      while (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
      nr = Math.max(nr, lines.length);
      nc = Math.max(nc, ...lines.map((l: string) => l.split("\t").length));
    }
    if (nr * nc > 40000) return; // safety cap — don't scan an enormous paste
    for (let r = 0; r < nr; r++)
      for (let c = 0; c < nc; c++) {
        const cell = ws.getRange(sr + r, sc + c);
        if (isBlankText(cell.getValue?.())) {
          if (cell.clear) cell.clear({ contentsOnly: true });
          else cell.setValueForCell({ v: null });
        }
      }
  } catch {}
}

// A full Excel/Google-Sheets-like spreadsheet (ribbon, formatting, formulas)
// powered by Univer. Loads an IWorkbookData snapshot and reports edits back.
export default function UniverSheet({
  snapshot,
  onChange,
  onReady,
  onRequestFind,
  readOnly,
}: {
  snapshot: any;
  onChange?: (snap: any) => void;
  onReady?: (api: any, phase: "mount" | "replace") => void;
  onRequestFind?: () => void;
  readOnly?: boolean;
}) {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<any>(null);
  const apiRef = useRef<any>(null);
  const lastRef = useRef("");
  const sigRef = useRef<(o: unknown) => string>((o) => JSON.stringify(o));
  const replacingRef = useRef(false); // applying an external snapshot in place
  const initialSnapRef = useRef(snapshot); // the snapshot we mounted with
  const onChangeRef = useRef(onChange);
  onChangeRef.current = onChange;
  const onReadyRef = useRef(onReady);
  onReadyRef.current = onReady;
  const onRequestFindRef = useRef(onRequestFind);
  onRequestFindRef.current = onRequestFind;
  // Live Format Painter status (0/1/2), mirrored here so the Esc handler can
  // cancel it without querying Univer's internal service.
  const painterStatusRef = useRef(0);
  // Source cell's row heights + column widths, captured when the painter is
  // armed. Univer's painter only copies cell STYLES (font, fill, border, number
  // format, alignment) + merges — Excel also carries the row height / column
  // width, so we copy those ourselves onto the target when it applies.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const painterSrcRef = useRef<any>(null);
  // Viewer (share role "viewer") → the whole workbook is read-only: no cell
  // editing, the edit toolbar is disabled, and undo/redo/word-delete are inert.
  const readOnlyRef = useRef(readOnly);
  readOnlyRef.current = readOnly;
  // Lock/unlock the workbook to match the access role. Retried across the mount
  // window because the permission service (and the workbook instance) aren't ready
  // the instant the sheet is created — a single early call is silently dropped, so
  // the sheet stays editable. Each retry reads the LIVE role, so it also catches
  // the role arriving slightly after mount.
  const applyEditable = (api: any) => {
    const set = () => {
      try {
        const wb = api?.getActiveWorkbook?.();
        if (wb?.setEditable) wb.setEditable(!readOnlyRef.current);
      } catch {}
    };
    set();
    [150, 500, 1000, 2000, 3500].forEach((ms) => setTimeout(set, ms));
  };
  // Excel-style Shift+click column-header range selection.
  const shiftRef = useRef(false); // Shift held at the last pointer/key event
  const anchorColRef = useRef<number | null>(null); // last plain column click
  const colEvtDisposeRef = useRef<any>(null); // ColumnHeaderClick disposer
  const editEvtDisposeRef = useRef<any[]>([]); // cell-edit start/end disposers
  const editingRef = useRef(false); // a cell is currently being edited
  // Timestamp of the user's last keyboard/pointer interaction. A remote update
  // moves the active cell (Univer's value-set command reselects the cell it
  // writes), so we NEVER apply one while the user is actively entering data —
  // only once they've paused. This is what stops "the cell jumping somewhere
  // else" when two people type at the same time.
  const lastInteractionRef = useRef(0);
  // Last server snapshot we applied (the common ancestor for diffing the next
  // remote update — so we only touch the cells other collaborators changed).
  const prevSnapRef = useRef<any>(snapshot);
  // A remote update that arrived WHILE this user was mid-edit — applied once the
  // edit finishes so we never yank the cursor or clobber a half-typed cell.
  const pendingSnapRef = useRef<any>(null);
  // Applies a remote snapshot in place (defined below, kept in a ref so the
  // edit-ended event handler can call the latest version).
  const applyRemoteRef = useRef<(snap: any) => void>(() => {});

  useEffect(() => {
    let disposed = false;
    let saveIv: ReturnType<typeof setInterval> | undefined;
    let flushIv: ReturnType<typeof setInterval> | undefined;
    // Track whether Shift is held at click time (for column range selection).
    // Also stamp the interaction time so remote updates hold off until the user
    // pauses (see lastInteractionRef) — no cursor jump mid-typing.
    const onKey = (e: KeyboardEvent) => {
      shiftRef.current = e.shiftKey;
      lastInteractionRef.current = Date.now();
    };
    const onDown = (e: MouseEvent) => {
      shiftRef.current = e.shiftKey;
      lastInteractionRef.current = Date.now();
    };
    // Ctrl/Cmd+F → open Univer's Find. Required because the grid is drawn on a
    // <canvas>: the browser's native find can't see cell text (it always shows
    // 0/0), so we route the shortcut to Univer's own search of the cell data.
    const onFind = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && !e.shiftKey && (e.key === "f" || e.key === "F")) {
        // ALWAYS suppress the browser's native find: the grid is a <canvas>, so
        // browser find can't see cell text (shows 0/0). Open our own lightweight
        // find bar instead (the parent renders it and searches the cell data).
        e.preventDefault();
        e.stopPropagation();
        onRequestFindRef.current?.();
      }
    };
    // Shift+Tab → move the active cell one column left (like Excel). The browser
    // otherwise moves focus out of the canvas, which clears the selection.
    const onShiftTab = (e: KeyboardEvent) => {
      if (e.key !== "Tab" || !e.shiftKey) return;
      const t = e.target as HTMLElement | null;
      // Leave real form fields (Find box, Share dialog) alone. Univer's own cell
      // editor is a contentEditable DIV, so we deliberately DON'T skip that.
      if (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA")) return;
      const api = apiRef.current;
      const wb = api?.getActiveWorkbook?.();
      const ws = wb?.getActiveSheet?.();
      const active = wb?.getActiveRange?.();
      const r = active?.getRange ? active.getRange() : null;
      if (!ws || !wb || !r) return; // no active cell → let the browser handle it
      e.preventDefault();
      e.stopImmediatePropagation();
      // Already at the first column → nothing to the left, so stay put (don't
      // wrap to the previous row).
      if ((r.startColumn ?? 0) <= 0) return;
      // Move the active cell one column left. Prefer Univer's own command (it
      // skips hidden columns, like Tab); fall back to a direct move.
      try {
        api.executeCommand("sheet.command.move-selection", { direction: 3 }); // LEFT
      } catch {
        try {
          const row = r.startRow ?? 0;
          const col = Math.max(0, (r.startColumn ?? 0) - 1);
          wb.setActiveRange(ws.getRange(row, col));
        } catch {}
      }
    };
    // Left arrow at the first column → stay put (Univer otherwise wraps to the
    // previous row's last cell). Skipped while editing so the text caret works.
    const onArrowLeft = (e: KeyboardEvent) => {
      if (e.key !== "ArrowLeft" || e.shiftKey || e.ctrlKey || e.metaKey || e.altKey)
        return;
      if (editingRef.current) return; // editing a cell → let the caret move
      const t = e.target as HTMLElement | null;
      if (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA")) return;
      const wb = apiRef.current?.getActiveWorkbook?.();
      const active = wb?.getActiveRange?.();
      const r = active?.getRange ? active.getRange() : null;
      if (!wb || !r) return;
      if ((r.startColumn ?? 0) <= 0) {
        // Already in the first column → nothing to the left.
        e.preventDefault();
        e.stopImmediatePropagation();
      }
    };
    // Ctrl+PageDown / Ctrl+PageUp → switch to the next / previous sheet tab
    // (the Excel & Google Sheets standard). Alt+Tab can't be used — Windows
    // reserves it for the app switcher, so the page never receives it.
    const onSheetSwitch = (e: KeyboardEvent) => {
      if (!e.ctrlKey || e.shiftKey || e.altKey || e.metaKey) return;
      if (e.key !== "PageDown" && e.key !== "PageUp") return;
      const wb = apiRef.current?.getActiveWorkbook?.();
      if (!wb) return;
      const sheets = wb.getSheets?.() || [];
      if (sheets.length < 2) return;
      const active = wb.getActiveSheet?.();
      const idx = sheets.findIndex(
        (s: any) => s.getSheetId?.() === active?.getSheetId?.()
      );
      if (idx < 0) return;
      e.preventDefault();
      e.stopPropagation();
      const dir = e.key === "PageDown" ? 1 : -1;
      const next = (idx + dir + sheets.length) % sheets.length;
      try {
        wb.setActiveSheet(sheets[next]);
      } catch {}
    };
    // Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z → route undo/redo to Univer even when focus
    // has left the sheet. Univer only handles these while its own cell editor is
    // focused; the moment focus moves to another element (a toolbar button, a
    // dialog, the page body) the keypress is lost and undo appears to do nothing.
    // We forward to Univer's command — but bail when a real text field / the
    // Univer cell editor is focused, so we never double-undo or steal a field's
    // own undo (Univer already handles it there).
    const onUndoRedo = (e: KeyboardEvent) => {
      if (readOnlyRef.current) return; // viewer → no undo/redo
      if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
      const k = e.key.toLowerCase();
      const isUndo = k === "z" && !e.shiftKey;
      const isRedo = k === "y" || (k === "z" && e.shiftKey);
      if (!isUndo && !isRedo) return;
      const ae = document.activeElement as HTMLElement | null;
      // Univer's cell editor is a contentEditable DIV that holds DOM focus even
      // when NOT editing (just selecting a cell — e.g. right after a Format
      // Painter apply). In that state the browser runs its own no-op editable
      // "undo" and swallows the key, so undo appears dead. We therefore only
      // defer to the editor while a cell is ACTUALLY being edited; otherwise we
      // drive the workbook undo ourselves. Real text fields (Find box, dialogs)
      // and any OTHER contentEditable keep their own undo.
      const isUniverEditor =
        !!ae && ae.id === "__editor___INTERNAL_EDITOR__DOCS_NORMAL";
      if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA")) return;
      if (ae && ae.isContentEditable && !isUniverEditor) return;
      if (isUniverEditor && editingRef.current) return;
      // Use the workbook facade's undo/redo: it calls focusUnit() first, so it
      // works even though DOM focus is on some element outside the sheet (the
      // raw undo command would target the "focused unit", which is nothing here).
      const wb = apiRef.current?.getActiveWorkbook?.();
      if (!wb) return;
      e.preventDefault();
      e.stopImmediatePropagation();
      try {
        if (isUndo) wb.undo();
        else wb.redo();
      } catch {}
    };
    window.addEventListener("keydown", onKey);
    window.addEventListener("keyup", onKey);
    window.addEventListener("pointerdown", onDown, true);
    window.addEventListener("keydown", onFind, true);
    window.addEventListener("keydown", onShiftTab, true);
    window.addEventListener("keydown", onArrowLeft, true);
    window.addEventListener("keydown", onSheetSwitch, true);
    // Ctrl+Backspace / Ctrl+Shift+Backspace → delete the previous WORD while
    // editing a cell (Univer's canvas editor only deletes a single character and
    // ignores the OS word-delete). We act against Univer's document model: read
    // the in-progress text, find the previous word boundary from the end, select
    // that span and run Univer's own delete command.
    const onWordDelete = (e: KeyboardEvent) => {
      if (readOnlyRef.current) return; // viewer → no editing
      if (e.key !== "Backspace" || !(e.ctrlKey || e.metaKey) || e.altKey) return;
      if (!editingRef.current) return; // only while a cell is being edited
      const api = apiRef.current;
      if (!api) return;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      let doc: any;
      try {
        doc = api.getActiveDocument?.();
      } catch {
        return;
      }
      if (!doc) return;
      let content = "";
      try {
        content = doc.getSnapshot?.().body?.dataStream ?? "";
      } catch {
        return;
      }
      content = content.replace(/[\r\n]+$/, ""); // drop the trailing paragraph marker
      const end = content.length;
      if (end <= 0) return;
      let i = end; // caret assumed at the end (the common case: mid-typing)
      while (i > 0 && /\s/.test(content[i - 1])) i--; // eat trailing spaces
      while (i > 0 && !/\s/.test(content[i - 1])) i--; // eat the word itself
      if (i >= end) return;
      e.preventDefault();
      e.stopImmediatePropagation();
      try {
        doc.setSelection(i, end);
        api.executeCommand("doc.command.delete-left");
      } catch {}
    };
    // Esc while the Format Painter is armed → cancel it (Excel behaviour).
    // Skipped while editing a cell / in a text field so it never steals the
    // editor's own Escape (which cancels the in-progress edit).
    const onPainterEsc = (e: KeyboardEvent) => {
      if (e.key !== "Escape" || painterStatusRef.current === 0) return;
      if (editingRef.current) return;
      const ae = document.activeElement as HTMLElement | null;
      if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA")) return;
      const api = apiRef.current;
      if (!api) return;
      e.preventDefault();
      e.stopImmediatePropagation();
      // Either command toggles the painter OFF when its status is non-zero.
      try {
        api.executeCommand("sheet.command.set-once-format-painter");
      } catch {}
    };
    window.addEventListener("keydown", onUndoRedo, true);
    window.addEventListener("keydown", onWordDelete, true);
    window.addEventListener("keydown", onPainterEsc, true);
    (async () => {
      const presets = await import("@univerjs/presets");
      const { createUniver, LocaleType, defaultTheme } = presets as any;
      const { UniverSheetsCorePreset } = await import(
        "@univerjs/presets/preset-sheets-core"
      );
      // Don't recalculate every formula on load. Uploaded sheets already carry
      // Excel's cached results in each cell, so recomputing thousands of formulas
      // at open time only spins Univer's calc-progress reporter — which drives a
      // React state loop ("Maximum update depth"), pegs the CPU and starves the
      // sync save/poll timers. NO_CALCULATION keeps the cached values (and the
      // formulas in the formula bar); edits still recalculate their dependents.
      const { CalculationMode } = await import("@univerjs/sheets-formula");
      const enUS = (
        (await import(
          "@univerjs/presets/preset-sheets-core/locales/en-US"
        )) as any
      ).default;
      if (disposed || !containerRef.current) return;

      // Find/Replace preset intentionally NOT loaded: its Find is a heavy modal
      // dialog. We provide our own lightweight in-sheet find bar instead (Ctrl+F
      // routes to it via onRequestFind).
      const { univer, univerAPI } = createUniver({
        locale: LocaleType.EN_US,
        locales: { [LocaleType.EN_US]: enUS },
        theme: defaultTheme,
        presets: [
          UniverSheetsCorePreset({
            container: containerRef.current,
            formula: { initialFormulaComputing: CalculationMode.NO_CALCULATION },
          }),
        ],
      });
      univerRef.current = univer;
      apiRef.current = univerAPI;
      // Dev aid: expose the facade so behaviors can be verified from the console.
      if (
        typeof window !== "undefined" &&
        process.env.NODE_ENV !== "production"
      )
        (window as any).__univerAPI = univerAPI;

      univerAPI.createWorkbook(
        snapshot || {
          id: "wb-empty",
          sheetOrder: ["s1"],
          sheets: {
            s1: {
              id: "s1",
              name: "Sheet1",
              cellData: {},
              rowCount: 200,
              columnCount: 40,
            },
          },
        }
      );
      // Hand the facade API to the parent so it can drive the sheet (e.g. the
      // special-shows column toggle).
      try {
        onReadyRef.current?.(univerAPI, "mount");
      } catch {}
      applyEditable(univerAPI); // viewer → lock the workbook read-only

      // Add a native Format Painter button to Univer's own formatting ribbon.
      // The preset ships the command + service (and its BrushIcon) but not the
      // toolbar entry, so we register it into the FORMAT group next to
      // bold/italic. Single-click = apply once; double-click = sticky (Univer's
      // `subId`); the button auto-highlights while armed. Row height / column
      // width are carried on top by the apply listener above.
      try {
        const [ui, sheetsUi, core, sheets, design, rx] = await Promise.all([
          import("@univerjs/ui"),
          import("@univerjs/sheets-ui"),
          import("@univerjs/core"),
          import("@univerjs/sheets"),
          import("@univerjs/design"),
          import("rxjs"),
        ]);
        const {
          IMenuManagerService,
          RibbonPosition,
          RibbonStartGroup,
          MenuItemType,
          getMenuHiddenObservable,
          ComponentManager,
          COLOR_PICKER_COMPONENT,
        } = ui as any;
        const {
          IFormatPainterService,
          SetOnceFormatPainterCommand,
          SetInfiniteFormatPainterCommand,
        } = sheetsUi as any;
        const { UniverInstanceType, IUniverInstanceService } = core as any;
        const {
          SetWorksheetRowHeightMutation,
          SetWorksheetRowHeightMutationFactory,
          SetWorksheetColWidthMutation,
          SetWorksheetColWidthMutationFactory,
        } = sheets as any;
        const { Observable } = rx as any;
        const injector = univerRef.current?.__getInjector?.();

        // Excel "Standard Colors" in every color picker (font + fill) — driven
        // from APP CODE so it ships in the bundle on every deploy (no fragile
        // node_modules patch / patch-package build step). Both the font-color and
        // fill-color menus render the shared COLOR_PICKER_COMPONENT and pass their
        // own onChange, so a wrapper that renders the real ColorPicker plus a
        // "Standard Colors" strip (calling that same onChange) applies to whichever
        // picker is open, with the correct command. We use the ComponentManager's
        // own React utils so it's the exact React instance Univer renders with.
        try {
          const componentManager = injector?.get?.(ComponentManager);
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const ColorPicker = (design as any)?.ColorPicker;
          const ru = componentManager?.reactUtils;
          if (
            componentManager?.register &&
            ColorPicker &&
            COLOR_PICKER_COMPONENT &&
            ru?.createElement
          ) {
            const h = ru.createElement;
            const useR = ru.useRef;
            const useEff = ru.useEffect;
            const STD = [
              "#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050",
              "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0",
            ];
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const StandardColorPicker = (props: any) => {
              const hostRef = useR(null);
              useEff(() => {
                const host = hostRef.current;
                if (!host) return;
                const presets = host.querySelector(
                  '[data-u-comp="color-picker-presets"]'
                );
                // Append INSIDE the presets grid (after the theme rows, before
                // "More Colors"), matching Excel's layout. Guarded so re-renders
                // don't duplicate it.
                if (!presets || presets.querySelector("[data-svf-std]")) return;
                const label = document.createElement("div");
                label.className =
                  "univer-text-xs univer-font-medium univer-text-gray-500 univer-mt-1";
                label.textContent = "Standard Colors";
                label.setAttribute("data-svf-std", "label");
                const row = document.createElement("div");
                row.className =
                  "univer-grid univer-grid-flow-col univer-items-center univer-justify-between univer-gap-2";
                row.setAttribute("data-svf-std", "row");
                STD.forEach((c) => {
                  const b = document.createElement("button");
                  b.type = "button";
                  b.className =
                    "univer-box-border univer-size-5 univer-cursor-pointer univer-rounded-full univer-border univer-border-solid univer-border-transparent univer-transition-shadow";
                  b.style.backgroundColor = c;
                  b.addEventListener("click", () => {
                    try {
                      props.onChange && props.onChange(c);
                    } catch {}
                  });
                  row.appendChild(b);
                });
                presets.appendChild(label);
                presets.appendChild(row);
              });
              return h(
                "div",
                { ref: hostRef, className: "univer-grid univer-gap-2" },
                h(ColorPicker, props)
              );
            };
            componentManager.register(COLOR_PICKER_COMPONENT, StandardColorPicker);
          }
        } catch {}

        const menuManager = injector?.get?.(IMenuManagerService);
        if (menuManager?.mergeMenu && SetOnceFormatPainterCommand?.id) {
          const menuItemFactory = (accessor: any) => {
            const fps = accessor.get(IFormatPainterService);
            return {
              id: SetOnceFormatPainterCommand.id,
              subId: SetInfiniteFormatPainterCommand.id,
              type: MenuItemType.BUTTON,
              icon: "BrushIcon",
              title: "Format Painter",
              tooltip: "Format Painter (double-click to keep applying)",
              activated$: new Observable((subscriber: any) => {
                let sub: any;
                try {
                  sub = fps.status$.subscribe((s: number) =>
                    subscriber.next(s !== 0)
                  );
                } catch {
                  subscriber.next(false);
                }
                return () => {
                  try {
                    sub?.unsubscribe?.();
                  } catch {}
                };
              }),
              hidden$: getMenuHiddenObservable(
                accessor,
                UniverInstanceType.UNIVER_SHEET
              ),
            };
          };
          menuManager.mergeMenu({
            [RibbonPosition.START]: {
              [RibbonStartGroup.FORMAT]: {
                [SetOnceFormatPainterCommand.id]: {
                  order: 0,
                  menuItemFactory,
                },
              },
            },
          });
        }

        // Univer's Format Painter copies cell STYLES only. Fold row height +
        // column width into the SAME paint by registering a hook: its onApply
        // returns the row-height / column-width mutations (with matching undos),
        // which Univer executes together with the style copy and pushes as ONE
        // undo entry — so a single Ctrl+Z reverts the entire paint. onStatusChange
        // both tracks the armed state (for Esc) and snapshots the source's
        // row heights / column widths while the source is still selected.
        const fps = injector?.get?.(IFormatPainterService);
        const uis = injector?.get?.(IUniverInstanceService);
        if (fps?.addHook && uis) {
          fps.addHook({
            id: "svf-format-painter-dimensions",
            priority: 0,
            onStatusChange: (status: number) => {
              painterStatusRef.current = status;
              if (status === 0) return;
              try {
                const wb = apiRef.current?.getActiveWorkbook?.();
                const ws = wb?.getActiveSheet?.();
                const r = wb?.getActiveRange?.()?.getRange?.();
                if (!ws || !r) return;
                const nRows = r.endRow - r.startRow + 1;
                const nCols = r.endColumn - r.startColumn + 1;
                const rowHeights: number[] = [];
                for (let i = 0; i < nRows; i++)
                  rowHeights.push(ws.getRowHeight?.(r.startRow + i));
                const colWidths: number[] = [];
                for (let j = 0; j < nCols; j++)
                  colWidths.push(ws.getColumnWidth?.(r.startColumn + j));
                painterSrcRef.current = { nRows, nCols, rowHeights, colWidths };
              } catch {}
            },
            onApply: (unitId: string, subUnitId: string, targetRange: any) => {
              const empty = { undos: [], redos: [] };
              const src = painterSrcRef.current;
              if (!src || !targetRange) return empty;
              try {
                const workbook = uis.getUniverSheetInstance?.(unitId);
                const worksheet = workbook?.getSheetBySheetId?.(subUnitId);
                if (!worksheet) return empty;
                let { startRow, endRow, startColumn, endColumn } = targetRange;
                // Painting from a single cell expands to the source's size,
                // exactly like Univer's style apply does.
                if (startRow === endRow && startColumn === endColumn) {
                  endRow = startRow + src.nRows - 1;
                  endColumn = startColumn + src.nCols - 1;
                }
                const redos: any[] = [];
                const undos: any[] = [];
                // Row heights — tile the source pattern, batching equal runs.
                const nT = endRow - startRow + 1;
                for (let tr = 0; tr < nT; ) {
                  const h = src.rowHeights[tr % src.nRows];
                  let run = 1;
                  while (
                    tr + run < nT &&
                    src.rowHeights[(tr + run) % src.nRows] === h
                  )
                    run++;
                  if (h > 0) {
                    const params = {
                      unitId,
                      subUnitId,
                      ranges: [
                        {
                          startRow: startRow + tr,
                          endRow: startRow + tr + run - 1,
                          startColumn: 0,
                          endColumn: 0,
                        },
                      ],
                      rowHeight: h,
                    };
                    const undoParams = SetWorksheetRowHeightMutationFactory(
                      params,
                      worksheet
                    );
                    redos.push({ id: SetWorksheetRowHeightMutation.id, params });
                    undos.push({
                      id: SetWorksheetRowHeightMutation.id,
                      params: undoParams,
                    });
                  }
                  tr += run;
                }
                // Column widths — same tiling + batching.
                const nC = endColumn - startColumn + 1;
                for (let tc = 0; tc < nC; ) {
                  const w = src.colWidths[tc % src.nCols];
                  let run = 1;
                  while (
                    tc + run < nC &&
                    src.colWidths[(tc + run) % src.nCols] === w
                  )
                    run++;
                  if (w > 0) {
                    const params = {
                      unitId,
                      subUnitId,
                      ranges: [
                        {
                          startRow: 0,
                          endRow: 0,
                          startColumn: startColumn + tc,
                          endColumn: startColumn + tc + run - 1,
                        },
                      ],
                      colWidth: w,
                    };
                    const undoParams = SetWorksheetColWidthMutationFactory(
                      params,
                      worksheet
                    );
                    redos.push({ id: SetWorksheetColWidthMutation.id, params });
                    undos.push({
                      id: SetWorksheetColWidthMutation.id,
                      params: undoParams,
                    });
                  }
                  tc += run;
                }
                return { undos, redos };
              } catch {
                return empty;
              }
            },
          });
        }
      } catch {}

      // Excel-style column range selection: click a column header, then
      // Shift+click another → select every column between them (inclusive).
      // We use Univer's ColumnHeaderClick event (authoritative clicked column)
      // and override the selection with the correct anchor→target range.
      try {
        const EVT =
          univerAPI.Event?.ColumnHeaderClick ?? "ColumnHeaderClick";
        colEvtDisposeRef.current = univerAPI.addEvent(EVT, (p: any) => {
          const col = p?.column;
          if (typeof col !== "number") return;
          const ws = p.worksheet;
          const wb = p.workbook;
          if (
            shiftRef.current &&
            anchorColRef.current != null &&
            anchorColRef.current !== col &&
            ws &&
            wb
          ) {
            const a = Math.min(anchorColRef.current, col);
            const b = Math.max(anchorColRef.current, col);
            try {
              const maxRows = ws.getMaxRows?.() ?? 1000;
              const range = ws.getRange({
                startRow: 0,
                endRow: Math.max(0, maxRows - 1),
                startColumn: a,
                endColumn: b,
                rangeType: 2, // RANGE_TYPE.COLUMN
              });
              wb.setActiveRange(range);
            } catch {}
            // keep the anchor so further shift-clicks extend from the same start
          } else {
            anchorColRef.current = col; // new anchor (plain click)
          }
        });
      } catch {}

      // Track cell edit mode so the arrow-key guard doesn't fight the text caret.
      try {
        const EV = univerAPI.Event;
        editEvtDisposeRef.current.push(
          univerAPI.addEvent(EV?.SheetEditStarted ?? "SheetEditStarted", () => {
            editingRef.current = true;
            lastInteractionRef.current = Date.now();
          }),
          univerAPI.addEvent(EV?.SheetEditEnded ?? "SheetEditEnded", () => {
            editingRef.current = false;
            // A collaborator's update landed while we were typing — apply it now
            // that the edit is committed (no cursor jump, no clobbered cell).
            if (pendingSnapRef.current) {
              const pending = pendingSnapRef.current;
              pendingSnapRef.current = null;
              // Defer a tick so Univer finishes committing this edit first.
              setTimeout(() => applyRemoteRef.current?.(pending), 0);
            }
          }),
          // Pasting from an external app (Excel) writes EMPTY source cells as ""
          // or " " (text). "text + number" evaluates to #VALUE!, so a template
          // formula row that references a pasted-empty cell shows #VALUE! instead
          // of blank. After each paste, convert those whitespace/empty-string
          // cells in the pasted region to truly-empty, so the formulas recompute
          // to 0 (which the report's number format then displays as "-"). Only
          // touches blank-text cells — never real values.
          univerAPI.addEvent(
            EV?.ClipboardPasted ?? "ClipboardPasted",
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (p: any) => setTimeout(() => cleanPastedBlanks(univerAPI, p), 0)
          )
        );
      } catch {}
      // Stable content signature — ignore Univer's volatile "rev" counters so we
      // don't fire a "change" (and a save) on every poll when nothing changed.
      const sig = (o: unknown) =>
        JSON.stringify(o, (k, v) => (k === "rev" ? undefined : v));
      sigRef.current = sig;
      // Seed from the ACTUAL created workbook so the first real edit — not the
      // initial state — triggers the first save.
      try {
        const wb0 = univerAPI.getActiveWorkbook?.();
        const snap0 = wb0?.getSnapshot ? wb0.getSnapshot() : wb0?.save?.();
        lastRef.current = sig(snap0 ?? snapshot ?? "");
      } catch {
        lastRef.current = sig(snapshot ?? "");
      }

      // Poll the workbook snapshot for edits → notify the parent (debounced feel).
      saveIv = setInterval(() => {
        if (replacingRef.current) return; // mid external-update; don't echo it back
        try {
          const wb = univerAPI.getActiveWorkbook?.();
          if (!wb) return;
          const snap = wb.getSnapshot ? wb.getSnapshot() : wb.save?.();
          if (!snap) return;
          const s = sig(snap);
          if (s !== lastRef.current) {
            lastRef.current = s;
            onChangeRef.current?.(snap);
          }
        } catch {}
      }, 1000);

      // Apply any deferred collaborator update the instant the user goes idle.
      // (A remote snapshot that lands while they're typing is stashed in
      // pendingSnapRef; this drains it once they've paused for IDLE_APPLY_MS so
      // their cursor is never yanked mid-entry, but updates still show promptly.)
      flushIv = setInterval(() => {
        if (replacingRef.current) return;
        const pending = pendingSnapRef.current;
        if (!pending) return;
        if (editingRef.current) return;
        if (Date.now() - lastInteractionRef.current < IDLE_APPLY_MS) return;
        pendingSnapRef.current = null;
        applyRemoteRef.current?.(pending);
      }, 400);
    })();

    return () => {
      disposed = true;
      if (saveIv) clearInterval(saveIv);
      if (flushIv) clearInterval(flushIv);
      window.removeEventListener("keydown", onKey);
      window.removeEventListener("keyup", onKey);
      window.removeEventListener("pointerdown", onDown, true);
      window.removeEventListener("keydown", onFind, true);
      window.removeEventListener("keydown", onShiftTab, true);
      window.removeEventListener("keydown", onArrowLeft, true);
      window.removeEventListener("keydown", onSheetSwitch, true);
      window.removeEventListener("keydown", onUndoRedo, true);
      window.removeEventListener("keydown", onWordDelete, true);
      window.removeEventListener("keydown", onPainterEsc, true);
      try {
        colEvtDisposeRef.current?.dispose?.();
      } catch {}
      try {
        editEvtDisposeRef.current.forEach((d) => d?.dispose?.());
        editEvtDisposeRef.current = [];
      } catch {}
      // Defer the Univer engine dispose out of React's synchronous unmount:
      // Univer tears down its own React root, and doing that *during* React's
      // render/commit throws "synchronously unmount a root while rendering".
      const toDispose = univerRef.current;
      univerRef.current = null;
      apiRef.current = null;
      setTimeout(() => {
        try {
          toDispose?.dispose?.();
        } catch {}
      }, 0);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Resolve a cell to a comparable/applicable shape: { v, f, style } — with the
  // style id DEREFERENCED to its object (two server snapshots may number styles
  // differently, so comparing raw ids gives false diffs).
  const resolveCell = (cell: any, snap: any) => {
    if (cell == null) return null;
    const s = cell.s;
    const style =
      s != null && typeof s === "object"
        ? s
        : s != null
        ? snap?.styles?.[s] ?? null
        : null;
    const v = cell.v ?? null;
    const f = cell.f ?? null;
    if (v == null && f == null && style == null) return null;
    return { v, f, style };
  };

  // Apply ONLY the cells (and hidden-column changes) that differ between the last
  // applied snapshot and the incoming one, straight onto the live workbook —
  // leaving scroll, selection and any in-progress edit untouched. Returns false
  // if the change is structural (sheet added/removed) and needs a full rebuild.
  const patchInPlace = (api: any, prevSnap: any, newSnap: any): boolean => {
    const wb = api.getActiveWorkbook?.();
    if (!wb) return false;
    const prevSheets = prevSnap?.sheets || {};
    const newSheets = newSnap?.sheets || {};
    const newIds = Object.keys(newSheets);
    const prevIds = Object.keys(prevSheets);
    if (prevIds.length !== newIds.length || prevIds.some((id) => !newSheets[id]))
      return false; // sheet set changed → rebuild
    for (const sid of newIds) {
      const ws = wb.getSheetBySheetId?.(sid);
      if (!ws) return false;
      // Hidden-column changes (cheap; keeps collaborators' hide state in sync).
      const oldCol = prevSheets[sid]?.columnData || {};
      const newCol = newSheets[sid]?.columnData || {};
      for (const c of new Set([...Object.keys(oldCol), ...Object.keys(newCol)])) {
        const oh = !!oldCol[c]?.hd;
        const nh = !!newCol[c]?.hd;
        if (oh !== nh) {
          try {
            if (nh) ws.hideColumns(Number(c), 1);
            else ws.showColumns(Number(c), 1);
          } catch {}
        }
      }
      // Cell value / formula / style diffs.
      const oldCd = prevSheets[sid]?.cellData || {};
      const newCd = newSheets[sid]?.cellData || {};
      for (const r of new Set([...Object.keys(oldCd), ...Object.keys(newCd)])) {
        const oldRow = oldCd[r] || {};
        const newRow = newCd[r] || {};
        for (const c of new Set([
          ...Object.keys(oldRow),
          ...Object.keys(newRow),
        ])) {
          const a = resolveCell(oldRow[c], prevSnap);
          const b = resolveCell(newRow[c], newSnap);
          if (JSON.stringify(a) === JSON.stringify(b)) continue;
          try {
            const rng = ws.getRange(Number(r), Number(c));
            if (b == null) {
              rng.setValueForCell({ v: null, f: null, s: null });
            } else {
              rng.setValueForCell({
                v: b.v ?? null,
                f: b.f ?? null,
                s: b.style ?? null,
              });
            }
          } catch {}
        }
      }
    }
    return true;
  };

  // Last-resort full rebuild (structural change) — capture the viewport +
  // selection first and restore them after, so the user isn't thrown to the top.
  const fullRebuild = (api: any, snap: any) => {
    const cur = api.getActiveWorkbook?.();
    const id = cur?.getId?.();
    let sid: any = null;
    let a1: any = null;
    let scroll: any = null;
    try {
      const ws = cur?.getActiveSheet?.();
      sid = ws?.getSheetId?.();
      a1 = cur?.getActiveRange?.()?.getA1Notation?.();
      scroll = ws?.getScrollState?.();
    } catch {}
    if (id && api.disposeUnit) {
      try {
        api.disposeUnit(id);
      } catch {}
    }
    api.createWorkbook(snap);
    try {
      onReadyRef.current?.(api, "replace");
    } catch {}
    applyEditable(api); // re-assert read-only after a structural rebuild
    try {
      const wb = api.getActiveWorkbook?.();
      const ws = sid ? wb?.getSheetBySheetId?.(sid) : null;
      if (ws) {
        wb.setActiveSheet?.(ws);
        if (a1) wb.setActiveRange?.(ws.getRange(a1));
        if (scroll)
          ws.scrollToCell?.(
            scroll.sheetViewStartRow || 0,
            scroll.sheetViewStartColumn || 0
          );
      }
    } catch {}
  };

  // Keep the apply function current so the edit-ended handler calls the latest.
  applyRemoteRef.current = (snap: any) => {
    const api = apiRef.current;
    if (!api || !snap) return;
    const sig = sigRef.current;
    if (sig(snap) === lastRef.current) return; // matches what's shown (our own save)
    // Don't apply while the user is editing OR still actively entering data.
    // Writing a remote cell reselects it, so applying mid-typing yanks the
    // active cell away. Stash the newest snapshot and let the flush timer apply
    // it the moment the user pauses (idle for IDLE_APPLY_MS). This is the fix
    // for the cell jumping when several people type at once.
    if (
      editingRef.current ||
      Date.now() - lastInteractionRef.current < IDLE_APPLY_MS
    ) {
      pendingSnapRef.current = snap;
      return;
    }
    replacingRef.current = true;
    try {
      // Capture this user's viewport + selection BEFORE the patch. Applying a
      // remote/round-tripped snapshot can move the cursor — hiding/showing a
      // column runs Univer's SetColHidden command, which reselects the affected
      // columns and yanks the active cell "somewhere else" while the user is
      // working. We restore exactly where they were afterwards. (fullRebuild
      // already restores its own view, so only guard the in-place patch path.)
      let savedSel: string | null = null;
      let savedScroll: any = null;
      try {
        const wb0 = api.getActiveWorkbook?.();
        const ws0 = wb0?.getActiveSheet?.();
        savedSel = wb0?.getActiveRange?.()?.getA1Notation?.() ?? null;
        savedScroll = ws0?.getScrollState?.() ?? null;
      } catch {}
      const ok = patchInPlace(api, prevSnapRef.current, snap);
      if (!ok) fullRebuild(api, snap);
      else if (savedSel || savedScroll) {
        // Put the user's selection + viewport back exactly where they were.
        const restore = () => {
          try {
            const wb1 = api.getActiveWorkbook?.();
            const ws1 = wb1?.getActiveSheet?.();
            if (!ws1) return;
            if (savedSel) wb1.setActiveRange?.(ws1.getRange(savedSel));
            // Only restore scroll when we actually captured a position — else
            // scrollToCell(0,0) would itself jump the viewport to the top.
            const sr = savedScroll?.sheetViewStartRow;
            const sc = savedScroll?.sheetViewStartColumn;
            if (Number.isFinite(sr) || Number.isFinite(sc))
              ws1.scrollToCell?.(Number(sr) || 0, Number(sc) || 0);
          } catch {}
        };
        restore();
        // Writing a remote cell can move the selection on a LATER frame (Univer
        // flushes its selection update asynchronously), which would override the
        // synchronous restore above. Re-assert once more next frame — but only if
        // the user hasn't clicked/typed since, so we never fight a fresh action.
        const tAtApply = lastInteractionRef.current;
        try {
          requestAnimationFrame(() => {
            if (lastInteractionRef.current === tAtApply) restore();
          });
        } catch {}
      }
      prevSnapRef.current = snap;
      // Re-seed the change signature from the ACTUAL workbook so the save loop
      // doesn't echo this remote update straight back to the server.
      try {
        const wb = api.getActiveWorkbook?.();
        const after = wb?.getSnapshot ? wb.getSnapshot() : wb?.save?.();
        lastRef.current = sig(after ?? snap);
      } catch {
        lastRef.current = sig(snap);
      }
    } catch {}
    setTimeout(() => (replacingRef.current = false), 400);
  };

  // When the parent passes a NEW snapshot (a remote collaborator's save), apply
  // just the changed cells IN PLACE — never rebuilding the workbook, so the other
  // user's scroll position, selection and in-progress edit are all preserved.
  useEffect(() => {
    if (snapshot === initialSnapRef.current) return; // that's the mount value
    if (!apiRef.current || !snapshot) return;
    applyRemoteRef.current?.(snapshot);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [snapshot]);

  // The access role is fetched after mount, so re-assert the workbook's editable
  // state whenever readOnly changes (viewer → locked, editor → unlocked).
  useEffect(() => {
    if (apiRef.current) applyEditable(apiRef.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [readOnly]);

  return (
    <div
      ref={containerRef}
      className={readOnly ? "univer-readonly-view" : undefined}
      style={{ width: "100%", height: "100%", minHeight: 0 }}
    />
  );
}
