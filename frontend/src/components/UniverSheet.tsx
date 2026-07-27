"use client";

import { useEffect, useRef } from "react";
import "@univerjs/presets/lib/styles/preset-sheets-core.css";

/* eslint-disable @typescript-eslint/no-explicit-any */

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
}: {
  snapshot: any;
  onChange?: (snap: any) => void;
  onReady?: (api: any, phase: "mount" | "replace") => void;
  onRequestFind?: () => void;
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
  // Excel-style Shift+click column-header range selection.
  const shiftRef = useRef(false); // Shift held at the last pointer/key event
  const anchorColRef = useRef<number | null>(null); // last plain column click
  const colEvtDisposeRef = useRef<any>(null); // ColumnHeaderClick disposer
  const editEvtDisposeRef = useRef<any[]>([]); // cell-edit start/end disposers
  const editingRef = useRef(false); // a cell is currently being edited
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
    // Track whether Shift is held at click time (for column range selection).
    const onKey = (e: KeyboardEvent) => (shiftRef.current = e.shiftKey);
    const onDown = (e: MouseEvent) => (shiftRef.current = e.shiftKey);
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
    window.addEventListener("keydown", onKey);
    window.addEventListener("keyup", onKey);
    window.addEventListener("pointerdown", onDown, true);
    window.addEventListener("keydown", onFind, true);
    window.addEventListener("keydown", onShiftTab, true);
    window.addEventListener("keydown", onArrowLeft, true);
    window.addEventListener("keydown", onSheetSwitch, true);
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
    })();

    return () => {
      disposed = true;
      if (saveIv) clearInterval(saveIv);
      window.removeEventListener("keydown", onKey);
      window.removeEventListener("keyup", onKey);
      window.removeEventListener("pointerdown", onDown, true);
      window.removeEventListener("keydown", onFind, true);
      window.removeEventListener("keydown", onShiftTab, true);
      window.removeEventListener("keydown", onArrowLeft, true);
      window.removeEventListener("keydown", onSheetSwitch, true);
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
    // Mid-edit → stash it and apply when the edit ends (don't yank the cursor).
    if (editingRef.current) {
      pendingSnapRef.current = snap;
      return;
    }
    replacingRef.current = true;
    try {
      const ok = patchInPlace(api, prevSnapRef.current, snap);
      if (!ok) fullRebuild(api, snap);
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

  return (
    <div
      ref={containerRef}
      style={{ width: "100%", height: "100%", minHeight: 0 }}
    />
  );
}
