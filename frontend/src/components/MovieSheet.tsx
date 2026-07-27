"use client";

import "jspreadsheet-ce/dist/jspreadsheet.css";
import "jsuites/dist/jsuites.css";

import Link from "next/link";
import dynamic from "next/dynamic";
import { useEffect, useRef, useState } from "react";
import { Theatre, Screen, inr } from "@/data/theatres";
import { getMovie, refreshMovies } from "@/lib/movies";
// Static import (small, pure functions) so Fast Refresh always loads the latest
// converter — a dynamically-imported chunk can be cached stale in the browser.
import {
  excelToUniverSnapshot,
  parseThemePalette,
  setThemePalette,
  ensureFullBorder,
} from "@/lib/xlsxToUniver";

// Full Excel-like engine (ribbon + formatting) for uploaded .xlsx files.
const UniverSheet = dynamic(() => import("./UniverSheet"), { ssr: false });

type Cell = string | number;
type Day = { name: string; data: Cell[][] };

const PRICING_PAIRS = 5; // RATE/AUD class pairs (from the Nizam list)
const CORE_COLUMNS = [
  { type: "text", title: "S.NO", width: 56, align: "center" },
  { type: "text", title: "CENTRE NAME", width: 150 },
  { type: "text", title: "THEATRE NAME", width: 260 },
  { type: "text", title: "FORMAT", width: 90, align: "center" },
  { type: "numeric", title: "H.F.C", width: 90 },
  { type: "numeric", title: "H.F.A", width: 70 },
];
const PRICING_COLUMNS = Array.from({ length: PRICING_PAIRS }).flatMap(() => [
  { type: "numeric", title: "RATE", width: 66, align: "center" },
  { type: "numeric", title: "AUD", width: 66, align: "center" },
]);
const META_COLUMNS = [
  { type: "text", title: "Terms", width: 180 },
  { type: "numeric", title: "WEEK RENT", width: 90, align: "center" },
  { type: "numeric", title: "DAY RENT", width: 90, align: "center" },
  { type: "numeric", title: "Shows", width: 64, align: "center" },
];
const BASE_COLUMNS = [...CORE_COLUMNS, ...PRICING_COLUMNS, ...META_COLUMNS];
const BASE_LEN = BASE_COLUMNS.length; // 6 + 10 + 4 = 20
const CORE_LEN = CORE_COLUMNS.length; // 6
const RATE_COL_INDEX = CORE_LEN; // first RATE (primary, drives Show)
const TERMS_COL_INDEX = CORE_LEN + PRICING_PAIRS * 2; // 16
const WEEKRENT_COL_INDEX = TERMS_COL_INDEX + 1; // 17
const DAYRENT_COL_INDEX = TERMS_COL_INDEX + 2; // 18
const SHOWS_COUNT_COL_INDEX = TERMS_COL_INDEX + 3; // 19
// Terms + Week/Day rent collapse together with the "Show terms" toggle.
const TERMS_GROUP_COLS = [TERMS_COL_INDEX, WEEKRENT_COL_INDEX, DAYRENT_COL_INDEX];
// 0-based indexes of the pricing (RATE/AUD) columns, for collapse.
const PRICING_COL_INDEXES = Array.from(
  { length: PRICING_PAIRS * 2 },
  (_, k) => CORE_LEN + k
);
const FIXED_SHOWS = ["Noon", "Matinee", "1st", "2nd"]; // always present after specials
const DEFAULT_SPL = 3;
const TOTAL_COLUMNS = [
  { type: "numeric", title: "Gross", width: 100, align: "center" },
  { type: "numeric", title: "Nett", width: 100, align: "center" },
  { type: "numeric", title: "Share", width: 100, align: "center" },
  { type: "numeric", title: "Audience", width: 100, align: "center" },
];
const GT_COLUMNS = TOTAL_COLUMNS;
// Bump when the column layout changes so stale saved sheets auto-rebuild.
const SHEET_VERSION = "og-18";
// Blank spacer rows inserted after each theatre's screen rows (breathing room).
const SPACER_ROWS = 1;
// Blank rows after each red section TOTAL row.
const TOTAL_GAP_ROWS = 3;

/** Show names for a given number of special shows. */
function showsFor(splCount: number): string[] {
  const specials = Array.from({ length: splCount }, (_, i) => `Spl-${i + 1}`);
  return [...specials, ...FIXED_SHOWS];
}

/** Column definitions for a given number of special shows. */
function columnsFor(splCount: number) {
  const showCols = showsFor(splCount).flatMap(() => [
    { type: "numeric", title: "Show", width: 92, align: "center" },
    { type: "numeric", title: "Aud", width: 70, align: "center" },
  ]);
  return [...BASE_COLUMNS, ...showCols, ...TOTAL_COLUMNS, ...GT_COLUMNS];
}

function ncolsFor(splCount: number) {
  return BASE_LEN + showsFor(splCount).length * 2 + 8;
}

// 0-based indexes of the special-show columns (Show+Aud) for hide/show.
function splColIndexes(splCount: number): number[] {
  return Array.from({ length: splCount * 2 }, (_, k) => BASE_LEN + k);
}

// Convert a 0-based column index to a spreadsheet letter (A, B, … Z, AA…).
function colLetter(index: number): string {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Grouped header (super row). Special-show groups are omitted when hidden;
// the pricing (RATE/AUD) block collapses under a "Rates" group; the meta blank
// span (Terms + Shows) shrinks by 1 when the Terms column is hidden.
function nestedHeadersFor(
  splCount: number,
  hidden: boolean,
  termsHidden: boolean,
  ratesHidden: boolean
) {
  const shows = showsFor(splCount).filter(
    (sh) => !(hidden && sh.startsWith("Spl"))
  );
  // meta = Terms + WEEK RENT + DAY RENT + Shows; the first 3 hide with "terms"
  const metaSpan = META_COLUMNS.length - (termsHidden ? TERMS_GROUP_COLS.length : 0);
  const head: { title: string; colspan: number }[] = [
    { title: "", colspan: CORE_LEN }, // S.NO … H.F.A
  ];
  if (!ratesHidden) head.push({ title: "Rates", colspan: PRICING_PAIRS * 2 });
  head.push({ title: "", colspan: metaSpan }); // Terms + Shows
  return [
    [
      ...head,
      ...shows.map((sh) => ({ title: sh, colspan: 2 })),
      { title: "Today", colspan: TOTAL_COLUMNS.length },
      { title: "GT (Grand Total)", colspan: GT_COLUMNS.length },
    ],
  ];
}

function exportTitlesFor(splCount: number): string[] {
  return [
    ...BASE_COLUMNS.map((c) => c.title),
    ...showsFor(splCount).flatMap((sh) => [`${sh} Show`, `${sh} Aud`]),
    "Today Gross",
    "Today Nett",
    "Today Share",
    "Today Audience",
    "GT Gross",
    "GT Nett",
    "GT Share",
    "GT Audience",
  ];
}

function titlesFor(splCount: number, ncols: number): string[] {
  const out = exportTitlesFor(splCount).slice(0, ncols);
  while (out.length < ncols) out.push("");
  return out;
}

function toNum(v: unknown): Cell {
  if (v === null || v === undefined || v === "") return "";
  const n = Number(v);
  return Number.isFinite(n) ? n : (v as string);
}

/** The 5 RATE/AUD pricing pairs for a screen, from its class slabs (padded). */
function pricingCells(s: Screen): Cell[] {
  const slabs = s.slabs || [];
  const out: Cell[] = [];
  for (let i = 0; i < PRICING_PAIRS; i++) {
    const sl = slabs[i];
    out.push(toNum(sl?.rate), toNum(sl?.aud));
  }
  return out;
}

const EMPTY_TOTALS: Cell[] = ["", "", "", "", "", "", "", ""];

/** Top (current-week) term percentage from a Terms string, e.g.
 *  "On Nett 52.5%-42.5%-40%..." → 0.525. Returns 0 if none found. */
function topTermPct(terms: string | undefined): number {
  const m = /([\d.]+)\s*%/.exec(terms ?? "");
  return m ? (Number(m[1]) || 0) / 100 : 0;
}

/** Weekly rent from a rent-based Terms string. Handles Indian grouping with
 *  commas OR dots ("6,03,000" / "1.26,000") and "Rent Per Day" (×7 → weekly).
 *  e.g. "Rent Per Week Rs:6,03,000/-" → 603000; "Rent Per Day Rs:14,000" → 98000. */
function weeklyRentOf(terms: string | undefined): number {
  if (!terms || !/rent/i.test(terms)) return 0;
  const m = /rs[:.]?\s*([\d.,]+)/i.exec(terms);
  if (!m) return 0;
  let v = Number(m[1].replace(/[.,]/g, "")) || 0; // strip , and . separators
  if (/per\s*day/i.test(terms)) v *= 7; // daily rent → weekly
  return v;
}

/** WEEK RENT / DAY RENT cell values for a theatre — rent halls only, else "-". */
function rentCols(terms: string | undefined): { week: Cell; day: Cell } {
  const rent = weeklyRentOf(terms);
  if (rent <= 0) return { week: "-", day: "-" };
  return { week: rent, day: Math.round(rent / 7) };
}

/** Today Share formula for a theatre based on its Terms, with the actual Nett
 *  and Shows baked in as numbers (regenerated on every change).
 *  % terms → Nett × %.  Rent terms → Nett − (weeklyRent/28) × Shows.  Else "". */
function shareFormula(terms: string, nett: number, shows: number): string {
  const pct = topTermPct(terms);
  if (pct > 0) return `=${nett}*${pct}`;
  const rent = weeklyRentOf(terms);
  if (rent > 0) return `=${nett}-(${rent}/28)*${shows}`;
  return "";
}

/** Theatre-level Today/GT totals (single number = sum across ALL the theatre's
 *  screen rows), placed on the theatre's first row.
 *  Today Share = Today Nett × top term %. GT cells start = Today; the cumulative
 *  across days is applied afterwards by recomputeGT(). */
function theatreTotalsCells(
  splCount: number,
  theatreRows: number[],
  terms = ""
): Cell[] {
  const nShows = showsFor(splCount).length;
  const grossRefs: string[] = [];
  const audRefs: string[] = [];
  for (const r of theatreRows) {
    for (let k = 0; k < nShows; k++) {
      grossRefs.push(colLetter(BASE_LEN + 2 * k) + r);
      audRefs.push(colLetter(BASE_LEN + 2 * k + 1) + r);
    }
  }
  const firstR = theatreRows[0];
  const todayGrossIdx = BASE_LEN + nShows * 2;
  const G = colLetter(todayGrossIdx);
  const N = colLetter(todayGrossIdx + 1);
  const S = colLetter(todayGrossIdx + 2);
  const A = colLetter(todayGrossIdx + 3);
  void terms;
  void N;
  return [
    "=" + grossRefs.join("+"), // Today Gross (whole theatre)
    `=${G}${firstR}/1.18`, // Today Nett = Gross ÷ 1.18 (strip 18% GST)
    "", // Today Share — set by recomputeNett (with real Nett & Shows numbers)
    "=" + audRefs.join("+"), // Today Audience (whole theatre)
    `=${G}${firstR}`, // GT Gross (cumulative applied by recomputeGT)
    `=${N}${firstR}`, // GT Nett
    `=${S}${firstR}`, // GT Share
    `=${A}${firstR}`, // GT Audience
  ];
}

// ── Single source of truth for the sheet's row layout ──
// The Nizam list is split into sections (the S.NO restarts at 1 for each);
// after each section we emit a red TOTAL row summing that section's rows.
type LayoutRow =
  | { kind: "band"; label: string }
  | {
      kind: "screen";
      tId: number;
      sIdx: number;
      first: boolean;
      type: "single" | "multiplex";
      name: string;
    }
  | { kind: "spacer"; type: "single" | "multiplex" }
  | { kind: "total"; start: number; end: number };

function computeLayout(theatres: Theatre[]): LayoutRow[] {
  const rows: LayoutRow[] = [];
  const firstDist = theatres[0]?.district ?? null;
  let bandDist: string | null = null; // for DIST band emission (resets per section)
  let seenDist: string | null = null; // last district seen (for part-boundary detection)
  let secStart: number | null = null; // 0-based index where the current section began

  const closeSection = () => {
    if (secStart === null) return;
    rows.push({ kind: "total", start: secStart + 1, end: rows.length }); // 1-based range
    // blank gap rows after the red TOTAL (always visible)
    for (let k = 0; k < TOTAL_GAP_ROWS; k++) rows.push({ kind: "band", label: "" });
    secStart = null;
    bandDist = null; // next section restarts its district bands
  };

  for (const t of theatres) {
    // The Nizam list is two parts (multiplex block, then single-screen block);
    // each part starts again from the first district → close the section there.
    if (seenDist !== null && t.district === firstDist && seenDist !== firstDist)
      closeSection();
    if (secStart === null) secStart = rows.length;
    if (t.district !== bandDist) {
      rows.push({ kind: "band", label: `DIST : ${t.district}` });
      bandDist = t.district;
    }
    t.screens.forEach((s, i) =>
      rows.push({
        kind: "screen",
        tId: t.id,
        sIdx: i,
        first: i === 0,
        type: t.type,
        name: s.name ?? "",
      })
    );
    for (let k = 0; k < SPACER_ROWS; k++) rows.push({ kind: "spacer", type: t.type });
    seenDist = t.district;
  }
  closeSection();
  return rows;
}

/** 1-based sheet row numbers that are section TOTAL rows. */
function totalRowNumbers(theatres: Theatre[]): number[] {
  const out: number[] = [];
  computeLayout(theatres).forEach((r, i) => {
    if (r.kind === "total") out.push(i + 1);
  });
  return out;
}

/** 1-based row numbers of the "DIST : …" district band rows. */
function distRowNumbers(theatres: Theatre[]): number[] {
  const out: number[] = [];
  computeLayout(theatres).forEach((r, i) => {
    if (r.kind === "band" && /^DIST\s*:/.test(r.label)) out.push(i + 1);
  });
  return out;
}

/** Static Nizam sheet rows in OG format; show cells empty, theatre-level totals. */
function buildData(theatres: Theatre[], splCount: number): Cell[][] {
  const layout = computeLayout(theatres);
  const nShows = showsFor(splCount).length;
  const ncols = ncolsFor(splCount);
  const pad = (row: Cell[]) => {
    while (row.length < ncols) row.push("");
    return row;
  };
  // screen row numbers (1-based) per theatre, for theatre-level total formulas
  const screenRows: Record<number, number[]> = {};
  layout.forEach((r, i) => {
    if (r.kind === "screen") (screenRows[r.tId] ||= []).push(i + 1);
  });
  const byId: Record<number, Theatre> = {};
  for (const t of theatres) byId[t.id] = t;

  return layout.map((r) => {
    if (r.kind === "band") return pad([r.label]);
    if (r.kind === "spacer") return pad([]);
    if (r.kind === "total") return sectionTotalRow(splCount, r.start, r.end, ncols);
    // screen row
    const t = byId[r.tId];
    const s = t.screens[r.sIdx];
    return pad([
      r.first ? t.sno : "",
      r.first ? t.centre : "",
      s.name ?? "",
      s.format ?? "",
      toNum(s.hfc),
      toNum(s.hfa),
      ...pricingCells(s),
      r.first ? t.terms ?? "" : "", // Terms
      r.first ? rentCols(t.terms).week : "", // WEEK RENT
      r.first ? rentCols(t.terms).day : "", // DAY RENT
      "", // Shows
      ...Array(nShows * 2).fill(""),
      ...(r.first
        ? theatreTotalsCells(splCount, screenRows[r.tId], t.terms ?? "")
        : EMPTY_TOTALS),
    ]);
  });
}

/** A red section-total row summing rows [start..end] (1-based, inclusive). */
function sectionTotalRow(
  splCount: number,
  start: number,
  end: number,
  ncols: number
): Cell[] {
  const nShows = showsFor(splCount).length;
  const row: Cell[] = Array(ncols).fill("");
  row[2] = "TOTAL"; // THEATRE NAME column
  const sum = (col: number) =>
    `=SUM(${colLetter(col)}${start}:${colLetter(col)}${end})`;
  row[4] = sum(4); // H.F.C
  row[5] = sum(5); // H.F.A
  row[SHOWS_COUNT_COL_INDEX] = sum(SHOWS_COUNT_COL_INDEX); // Shows
  // show groups + the 8 Today/GT columns
  const from = BASE_LEN;
  const to = BASE_LEN + nShows * 2 + 8;
  for (let c = from; c < to; c++) row[c] = sum(c);
  return row;
}

function sanitizeSheetName(n: string): string {
  return n.replace(/[\\/?*[\]:]/g, " ").slice(0, 31) || "Sheet";
}

export default function MovieSheet({
  theatres,
  movieId,
}: {
  theatres: Theatre[];
  movieId: string;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const headerRef = useRef<HTMLDivElement>(null);
  const tabsSlotRef = useRef<HTMLDivElement>(null); // fixed slot for the day tabs
  const fileInputRef = useRef<HTMLInputElement>(null); // hidden Excel upload input
  const appendInputRef = useRef<HTMLInputElement>(null); // append-import input
  const [shareOpen, setShareOpen] = useState(false); // share dialog open
  const cardRef = useRef<HTMLDivElement>(null); // H.F.A class-breakdown hover card
  const cardRowRef = useRef<number | null>(null); // row currently shown in the card
  const cardPinnedRef = useRef(false); // card pinned open by a click
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const jssRef = useRef<any>(null);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sheetsRef = useRef<any[]>([]);
  const namesRef = useRef<string[]>([]);
  const [filter, setFilter] = useState<"all" | "multiplex" | "single">("all");
  const [splHidden, setSplHidden] = useState(false);
  const [showTerms, setShowTerms] = useState(false); // Terms column hidden by default
  const showTermsRef = useRef(false);
  const [splCount, setSplCount] = useState(DEFAULT_SPL);
  const splCountRef = useRef(DEFAULT_SPL); // source of truth for render/save
  const splHiddenRef = useRef(false);
  const ratesHiddenRef = useRef(false); // pricing (RATE/AUD) block collapsed
  const applyingRef = useRef(false); // guard against recursive onchange
  const followRaf = useRef<number | null>(null); // rAF id for scroll-follow
  const enhanceRef = useRef<() => void>(() => {}); // latest per-day enhancer
  // ── Server sync (shared sheet, auto-save + poll) ──
  const serverVersionRef = useRef(0); // last version we've seen/written
  const saveTimerRef = useRef<number | null>(null); // debounced save timer
  const lastEditRef = useRef(0); // ts of last local edit (don't clobber via poll)
  const pullingRef = useRef(false); // applying a server pull (suppress save)
  const roleRef = useRef<"editor" | "viewer">("editor"); // access role
  const [readOnly, setReadOnly] = useState(false);
  // Every movie is a Univer spreadsheet now (the old OG Nizam sheet is retired).
  const plainModeRef = useRef(true); // always Univer mode
  const [plainMode, setPlainMode] = useState(true); // mirror for UI
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const univerSnapRef = useRef<any>(null); // current Univer workbook snapshot
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [univerSnap, setUniverSnap] = useState<any>(null); // snapshot to render
  const [univerKey, setUniverKey] = useState(0); // remount Univer on external pull
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const univerApiRef = useRef<any>(null); // Univer facade API (drives col toggle)
  // Per-user view state (last active sheet + scroll + selection) so a reload /
  // reopen lands exactly where the user left off. Stored in localStorage.
  const lastViewRef = useRef<string>(""); // last-saved view state (JSON)
  const viewIvRef = useRef<number | null>(null); // periodic view-state saver
  const viewSettleIvRef = useRef<number | null>(null); // initial restore window
  // ── Lightweight in-sheet find bar (Ctrl+F) — live search as you type ──
  const [findOpen, setFindOpen] = useState(false);
  const [findQuery, setFindQuery] = useState("");
  // matches on the active sheet: {r,c}; findIdx = current match (0-based, -1 none)
  const [findMatches, setFindMatches] = useState<{ r: number; c: number }[]>([]);
  const [findIdx, setFindIdx] = useState(-1);
  const findInputRef = useRef<HTMLInputElement | null>(null);
  const [univerSplHidden, setUniverSplHidden] = useState(true); // Spl cols collapsed
  const univerReadyRef = useRef(false); // a real snapshot is loaded (guard saves)
  const [loading, setLoading] = useState(true); // loading the shared copy
  const [saveState, setSaveState] = useState<"idle" | "saving" | "saved">(
    "idle"
  ); // save-status toast
  const saveStateTimer = useRef<number | null>(null);
  const [me, setMe] = useState<{ email: string | null; isAdmin?: boolean }>({
    email: null,
  }); // signed-in user
  const saveKey = `svf.moviesheet.${movieId}`;

  // Show-entry dialog (right-click a Show/Aud cell → add per-class audience).
  const [entry, setEntry] = useState<{
    y: number;
    showCol: number;
    audCol: number;
    name: string;
    show: string;
    cap: number;
    storeKey: string;
  } | null>(null);
  const [entryRows, setEntryRows] = useState<{ price: string; aud: string }[]>([]);
  const [formulasOpen, setFormulasOpen] = useState(false); // "how it's calculated" panel
  // Cell inspector (click a derived cell → see/edit how it's calculated)
  const [inspect, setInspect] = useState<{
    x: number;
    y: number;
    label: string;
    why: string;
    worked: string;
    value: string;
    formula: string; // "" when the cell is a plain computed number (not editable)
    values?: string; // the formula with cell refs replaced by actual numbers
    ded?: { tid: number; value: number; day: string; overridden: boolean }; // Nett cell
    input?: boolean; // true when the cell is a direct (typed) input, not calculated
  } | null>(null);

  // theatreId -> { first: firstRowNum, rows: allScreenRowNums } (1-based)
  function buildTheatreRows(): Record<number, { first: number; rows: number[] }> {
    const map: Record<number, { first: number; rows: number[] }> = {};
    computeLayout(theatres).forEach((r, i) => {
      if (r.kind !== "screen") return;
      const rn = i + 1;
      const e = (map[r.tId] ||= { first: rn, rows: [] });
      e.rows.push(rn);
    });
    return map;
  }

  const isFilled = (v: unknown) => v !== "" && v !== null && v !== undefined;

  // Auto-fill a Show cell = RATE × Aud for a given data row (0-based y).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function computeShow(ws: any, y: number, showCol: number) {
    const rate = Number(ws.getValueFromCoords(RATE_COL_INDEX, y)) || 0;
    const audRaw = ws.getValueFromCoords(showCol + 1, y);
    if (!isFilled(audRaw)) {
      ws.setValueFromCoords(showCol, y, "", true);
    } else {
      ws.setValueFromCoords(showCol, y, rate * (Number(audRaw) || 0), true);
    }
  }

  // Recompute the Shows count (0.5 per filled Show + 0.5 per filled Aud) for a theatre.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function recomputeShows(ws: any, tid: number, theatreRows: Record<number, { first: number; rows: number[] }>) {
    const info = theatreRows[tid];
    if (!info) return;
    const nShows = showsFor(splCountRef.current).length;
    let count = 0;
    for (const r of info.rows) {
      for (let k = 0; k < nShows; k++) {
        const sv = ws.getValueFromCoords(BASE_LEN + 2 * k, r - 1);
        const av = ws.getValueFromCoords(BASE_LEN + 2 * k + 1, r - 1);
        if (isFilled(sv)) count += 0.5;
        if (isFilled(av)) count += 0.5;
      }
    }
    ws.setValueFromCoords(SHOWS_COUNT_COL_INDEX, info.first - 1, count || "", true);
  }

  // Sum the theatre's raw Show / Aud cells (these are plain numbers, so reading
  // them is reliable — unlike the Today formulas, which need processed=true).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  // Cheap raw numeric read of a cell.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function numAt(ws: any, x: number, y: number): number {
    const raw = ws.getValueFromCoords(x, y);
    if (raw === "" || raw == null) return 0;
    const n = Number(raw);
    return Number.isFinite(n) ? n : 0;
  }
  // Read a cell's COMPUTED value (processed=true) — handles formula cells.
  // Callers must batch these BEFORE writing any cells (writes dirty the sheet and
  // make each subsequent processed read trigger a full recalc → freeze).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function computedNum(ws: any, x: number, y: number): number {
    const disp = ws.getValueFromCoords(x, y, true);
    const n = Number(String(disp ?? "").replace(/[₹,\s]/g, ""));
    return Number.isFinite(n) ? n : 0;
  }
  // Gross / Audience for a theatre = its Today-Gross / Today-Audience cell (the
  // authoritative SUM that's actually displayed). Falls back to a raw sum of the
  // Show/Aud cells. Reads the computed cell, so batch calls before writes.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function grossOfTheatre(ws: any, info: { first: number; rows: number[] }): number {
    const nShows = showsFor(splCountRef.current).length;
    const g = computedNum(ws, BASE_LEN + nShows * 2, info.first - 1);
    if (g > 0) return g;
    let s = 0;
    for (const r of info.rows)
      for (let k = 0; k < nShows; k++) s += numAt(ws, BASE_LEN + 2 * k, r - 1);
    return s;
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function audOfTheatre(ws: any, info: { first: number; rows: number[] }): number {
    const nShows = showsFor(splCountRef.current).length;
    const a = computedNum(ws, BASE_LEN + nShows * 2 + 3, info.first - 1);
    if (a > 0) return a;
    let s = 0;
    for (const r of info.rows)
      for (let k = 0; k < nShows; k++) s += numAt(ws, BASE_LEN + 2 * k + 1, r - 1);
    return s;
  }

  // ── Per-theatre fixed deduction (editable; hardcoded defaults from OG) ──
  function loadDeductions(): Record<string, number> {
    if (typeof window === "undefined") return {};
    try {
      return JSON.parse(localStorage.getItem(`svf.deduction.${movieId}`) || "{}");
    } catch {
      return {};
    }
  }
  function deductionFor(tid: number): number {
    const all = loadDeductions();
    if (all[tid] != null) return Number(all[tid]) || 0;
    const t = theatres.find((x) => x.id === tid);
    return t && t.type === "single" ? 500 : 0; // OG: most singles 500, multiplex 0
  }
  function saveDeduction(tid: number, val: number) {
    const all = loadDeductions();
    all[tid] = val;
    localStorage.setItem(`svf.deduction.${movieId}`, JSON.stringify(all));
  }

  // Manual Nett-formula overrides (per day + theatre). When set, recomputeNett
  // writes the user's formula instead of the auto per-class one.
  function loadNettOverrides(): Record<string, string> {
    if (typeof window === "undefined") return {};
    try {
      return JSON.parse(localStorage.getItem(`svf.nettoverride.${movieId}`) || "{}");
    } catch {
      return {};
    }
  }
  function saveNettOverride(day: string, tid: number, formula: string) {
    const all = loadNettOverrides();
    all[`${day}|${tid}`] = formula;
    localStorage.setItem(`svf.nettoverride.${movieId}`, JSON.stringify(all));
  }
  function clearNettOverride(day: string, tid: number) {
    const all = loadNettOverrides();
    delete all[`${day}|${tid}`];
    localStorage.setItem(`svf.nettoverride.${movieId}`, JSON.stringify(all));
  }

  // Total audience per ticket-price class for a theatre on a day (from the
  // per-class show-entry breakdowns saved by the Add-show-entry dialog).
  function classAudFor(day: string, tid: number): Record<number, number> {
    const entries = loadShowEntries();
    const out: Record<number, number> = {};
    const prefix = `${day}|${tid}|`;
    for (const k of Object.keys(entries)) {
      if (!k.startsWith(prefix)) continue;
      for (const r of entries[k]) {
        const p = Number(r.price) || 0;
        const a = Number(r.aud) || 0;
        if (p > 0 && a > 0) out[p] = (out[p] || 0) + a;
      }
    }
    return out;
  }

  // GST is per class: >₹105 → 18%, ≤₹105 → 5% (both GST-inclusive extraction).
  function gstOfClasses(cls: Record<number, number>): number {
    let gst = 0;
    for (const p of Object.keys(cls)) {
      const price = Number(p);
      const coll = price * cls[Number(p)];
      gst += price > 105 ? (coll * 18) / 118 : (coll * 5) / 105;
    }
    return gst;
  }

  // Build the OG-style Nett formula for a theatre: Gross − per-class GST − deduction.
  // High classes (>₹105) at 18/118, low (≤₹105) at 5/105 — exactly like OG.
  // The gross is baked in as a literal number (regenerated on every change), so
  // the formula reads with real figures instead of a cell reference.
  function nettFormulaFor(day: string, tid: number, gross: number): string {
    const ded = deductionFor(tid);
    const cls = classAudFor(day, tid);
    const highs: string[] = [];
    const lows: string[] = [];
    for (const p of Object.keys(cls)) {
      const price = Number(p);
      const term = `${price}*${cls[price]}`;
      (price > 105 ? highs : lows).push(term);
    }
    if (!highs.length && !lows.length) {
      // no class breakdown → flat 18%
      return ded
        ? `=${gross}-(${gross}*18/118+${ded})`
        : `=${gross}-${gross}*18/118`;
    }
    const parts: string[] = [];
    if (highs.length) parts.push(`(${highs.join("+")})*18/118`);
    if (lows.length) parts.push(`(${lows.join("+")})*5/105`);
    if (ded) parts.push(String(ded));
    return `=${gross}-(${parts.join("+")})`;
  }

  // Refresh the Terms cell from the DB (theatres prop) when it still holds the
  // generic default — so exact OG terms (incl. rent) show without a full rebuild.
  const DEFAULT_TERMS = "On Nett 55%-45%-40%-35%-30%";
  function syncTerms() {
    const sheets = sheetsRef.current;
    if (!sheets.length) return;
    const theatreRows = buildTheatreRows();
    const byId: Record<number, Theatre> = {};
    for (const t of theatres) byId[t.id] = t;
    const prev = applyingRef.current;
    applyingRef.current = true;
    try {
      for (const ws of sheets) {
        for (const [tidStr, info] of Object.entries(theatreRows)) {
          const t = byId[Number(tidStr)];
          if (!t) continue;
          const cell = String(
            ws.getValueFromCoords(TERMS_COL_INDEX, info.first - 1) ?? ""
          );
          if (cell !== (t.terms ?? "") && (cell === DEFAULT_TERMS || cell === "")) {
            ws.setValueFromCoords(TERMS_COL_INDEX, info.first - 1, t.terms ?? "", true);
          }
        }
      }
    } catch {}
    applyingRef.current = prev;
  }

  // Refill Show/Aud cells from saved per-class show entries (survives rebuilds).
  function restoreShowEntries() {
    const entries = loadShowEntries();
    const keys = Object.keys(entries);
    if (!keys.length) return;
    const rowMap = buildRowMap(); // `${tid}|${screen}` -> rowNum (1-based)
    const idxByName: Record<string, number> = {};
    namesRef.current.forEach((n, i) => (idxByName[n] = i));
    const prev = applyingRef.current;
    applyingRef.current = true;
    try {
      for (const key of keys) {
        const parts = key.split("|");
        if (parts.length < 4) continue;
        const day = parts[0];
        const tid = parts[1];
        const group = Number(parts[parts.length - 1]);
        const screen = parts.slice(2, -1).join("|");
        const si = idxByName[day];
        const ws = si != null ? sheetsRef.current[si] : null;
        const row = rowMap[`${tid}|${screen}`];
        if (!ws || !row) continue;
        const rows = entries[key] || [];
        const totalAud = rows.reduce((n, r) => n + (Number(r.aud) || 0), 0);
        const totalShow = rows.reduce(
          (n, r) => n + (Number(r.price) || 0) * (Number(r.aud) || 0),
          0
        );
        const showCol = BASE_LEN + 2 * group;
        ws.setValueFromCoords(showCol, row - 1, totalShow || "", true);
        ws.setValueFromCoords(showCol + 1, row - 1, totalAud || "", true);
      }
      // refresh Shows count for every theatre
      const theatreRows = buildTheatreRows();
      for (const ws of sheetsRef.current)
        for (const tid of Object.keys(theatreRows))
          recomputeShows(ws, Number(tid), theatreRows);
    } catch {}
    applyingRef.current = prev;
  }

  // Write the per-class Nett formula into each theatre's Nett cell.
  function recomputeNett(onlyTid?: number) {
    const sheets = sheetsRef.current;
    if (!sheets.length) return;
    const nShows = showsFor(splCountRef.current).length;
    const todayIdx = BASE_LEN + nShows * 2; // Gross col; +1 = Nett
    const theatreRows = buildTheatreRows();
    const overrides = loadNettOverrides();
    const prevGuard = applyingRef.current;
    applyingRef.current = true;
    try {
      sheets.forEach((ws, si) => {
        const day = namesRef.current[si] ?? `Day ${si + 1}`;
        const entries = Object.entries(theatreRows).filter(
          ([tidStr]) => !onlyTid || Number(tidStr) === onlyTid
        );
        // PHASE 1 — read every theatre's gross (computed cells) BEFORE writing
        // anything, so the formula cache stays warm (one recalc, not one/theatre).
        const grossMap: Record<number, number> = {};
        for (const [tidStr, info] of entries)
          grossMap[Number(tidStr)] = grossOfTheatre(ws, info);
        // PHASE 2 — write Nett & Share.
        for (const [tidStr, info] of entries) {
          const tid = Number(tidStr);
          const fr = info.first;
          const gross = grossMap[tid];
          const ov = overrides[`${day}|${tid}`];
          if (gross <= 0 && !ov) {
            ws.setValueFromCoords(todayIdx + 1, fr - 1, "", true);
            ws.setValueFromCoords(todayIdx + 2, fr - 1, "", true); // clear Share
            continue;
          }
          ws.setValueFromCoords(
            todayIdx + 1,
            fr - 1,
            ov || nettFormulaFor(day, tid, gross), // manual override wins
            true
          );
          // (re)build the Today Share formula with the ACTUAL Nett & Shows numbers.
          const cls = classAudFor(day, tid);
          const gst = Object.keys(cls).length
            ? gstOfClasses(cls)
            : (gross * 18) / 118;
          let nettVal = Math.max(0, gross - gst - deductionFor(tid));
          if (ov) {
            const rv = Number(ws.getValueFromCoords(todayIdx + 1, fr - 1, true));
            if (Number.isFinite(rv)) nettVal = rv; // override → use its computed value
          }
          nettVal = Math.round(nettVal * 100) / 100;
          const shows = Number(ws.getValueFromCoords(SHOWS_COUNT_COL_INDEX, fr - 1)) || 0;
          const terms = String(ws.getValueFromCoords(TERMS_COL_INDEX, fr - 1) ?? "");
          ws.setValueFromCoords(
            todayIdx + 2,
            fr - 1,
            shareFormula(terms, nettVal, shows) || "",
            true
          );
        }
      });
    } catch {}
    applyingRef.current = prevGuard;
  }

  // Cumulative GT across the day tabs: each day's GT = Σ Today from Day 1..that day.
  // Writes plain numbers into the 4 GT columns on each theatre's first row.
  function recomputeGT(onlyTid?: number) {
    const sheets = sheetsRef.current;
    if (!sheets.length) return;
    const nShows = showsFor(splCountRef.current).length;
    const todayIdx = BASE_LEN + nShows * 2; // Today Gross col; +1 Nett +2 Share +3 Aud
    const gtIdx = todayIdx + 4; // GT Gross col
    const theatreRows = buildTheatreRows();
    const prevGuard = applyingRef.current;
    applyingRef.current = true;
    try {
      const entries = Object.entries(theatreRows).filter(
        ([tidStr]) => !onlyTid || Number(tidStr) === onlyTid
      );
      // PHASE 1 — read every (theatre, day) Today value with NO writes, so the
      // processed reads hit the warm formula cache (one recalc, not one each).
      const today: Record<number, number[][]> = {}; // tid -> per-day [g,n,s,a]
      for (const [tidStr, info] of entries) {
        const tid = Number(tidStr);
        const y = info.first - 1;
        today[tid] = sheets.map((ws, si) => {
          const day = namesRef.current[si] ?? `Day ${si + 1}`;
          const gross = grossOfTheatre(ws, info);
          const aud = audOfTheatre(ws, info);
          let nett = 0;
          let share = 0;
          if (gross > 0) {
            const cls = classAudFor(day, tid);
            const gst = Object.keys(cls).length
              ? gstOfClasses(cls)
              : (gross * 18) / 118;
            nett = Math.max(0, gross - gst - deductionFor(tid));
            const terms = String(ws.getValueFromCoords(TERMS_COL_INDEX, y) ?? "");
            const pct = topTermPct(terms);
            if (pct > 0) share = nett * pct;
            else {
              const rent = weeklyRentOf(terms);
              if (rent > 0) {
                const shows =
                  Number(ws.getValueFromCoords(SHOWS_COUNT_COL_INDEX, y)) || 0;
                share = nett - (rent / 28) * shows;
              }
            }
          }
          return [gross, nett, share, aud];
        });
      }
      // PHASE 2 — accumulate across days and write the GT cells.
      for (const [tidStr, info] of entries) {
        const tid = Number(tidStr);
        const y = info.first - 1;
        const cum = [0, 0, 0, 0];
        sheets.forEach((ws, si) => {
          const t = today[tid][si];
          for (let j = 0; j < 4; j++) {
            cum[j] += t[j];
            ws.setValueFromCoords(
              gtIdx + j,
              y,
              cum[j] ? Math.round(cum[j] * 100) / 100 : "",
              true
            );
          }
        });
      }
    } catch {}
    applyingRef.current = prevGuard;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function handleCellChange(ws: any, x: number, y: number) {
    if (plainModeRef.current) return; // uploaded plain grid → no OG recompute
    const nShows = showsFor(splCountRef.current).length;
    // Editing any inline pricing cell (RATE/AUD class pairs) recomputes the
    // row's House-full (H.F.C = Σ rate×aud, H.F.A = Σ aud) and re-prices shows.
    if (x >= CORE_LEN && x < CORE_LEN + PRICING_PAIRS * 2) {
      let hfc = 0;
      let hfa = 0;
      for (let i = 0; i < PRICING_PAIRS; i++) {
        const rate = Number(ws.getValueFromCoords(CORE_LEN + 2 * i, y)) || 0;
        const aud = Number(ws.getValueFromCoords(CORE_LEN + 2 * i + 1, y)) || 0;
        hfc += rate * aud;
        hfa += aud;
      }
      ws.setValueFromCoords(4, y, hfc || "", true); // H.F.C
      ws.setValueFromCoords(5, y, hfa || "", true); // H.F.A
      for (let k = 0; k < nShows; k++) computeShow(ws, y, BASE_LEN + 2 * k);
    }
    const inShows = x >= BASE_LEN && x < BASE_LEN + nShows * 2;
    if (inShows && (x - BASE_LEN) % 2 === 1) {
      // an Aud cell changed → auto-fill its Show (no capacity validation)
      computeShow(ws, y, x - 1);
    }
    const rowTheatre = buildRowTheatre();
    const tid = rowTheatre[y + 1];
    if (tid) {
      recomputeShows(ws, tid, buildTheatreRows());
      recomputeNett(tid); // per-class GST Nett
      recomputeGT(tid); // refresh cumulative GT for this theatre across all days
    }
  }

  // row number (1-based) -> theatre id, for detecting the selected theatre.
  function buildRowTheatre(): Record<number, number> {
    const map: Record<number, number> = {};
    computeLayout(theatres).forEach((r, i) => {
      if (r.kind === "screen") map[i + 1] = r.tId;
    });
    return map;
  }

  function setSpl(n: number) {
    splCountRef.current = n;
    setSplCount(n);
  }

  // row number (1-based) -> 'band' | 'multiplex' | 'single', mirroring buildData order
  function buildRowTypes(): Record<number, "band" | "multiplex" | "single"> {
    const map: Record<number, "band" | "multiplex" | "single"> = {};
    computeLayout(theatres).forEach((r, i) => {
      const rn = i + 1;
      if (r.kind === "band" || r.kind === "total") map[rn] = "band"; // always visible
      else map[rn] = r.type; // screen & spacer follow the theatre's type
    });
    return map;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function applyFilterTo(sheets: any[], mode: "all" | "multiplex" | "single") {
    const types = buildRowTypes();
    const hide: number[] = [];
    const show: number[] = [];
    for (const [rowStr, type] of Object.entries(types)) {
      const y = Number(rowStr) - 1; // 0-based
      if (type === "band" || mode === "all") show.push(y);
      else if (mode === "multiplex") (type === "multiplex" ? show : hide).push(y);
      else (type === "single" ? show : hide).push(y);
    }
    for (const ws of sheets) {
      // Per-row calls with individual try/catch: one out-of-range index must
      // not abort the whole batch (which would leave rows un-hidden).
      const n =
        (ws?.options?.data?.length as number | undefined) ??
        Number.POSITIVE_INFINITY;
      for (const y of show) {
        if (y >= 0 && y < n) try { ws.showRow(y); } catch {}
      }
      for (const y of hide) {
        if (y >= 0 && y < n) try { ws.hideRow(y); } catch {}
      }
    }
  }

  function applyFilter(mode: "all" | "multiplex" | "single") {
    setFilter(mode);
    applyFilterTo(sheetsRef.current, mode);
  }

  function toggleSpecialShows() {
    const next = !splHiddenRef.current;
    splHiddenRef.current = next;
    setSplHidden(next);
    // Direct re-render (no fade) so the header rebuilds instantly; data is
    // unchanged so skip the GT recompute to keep it snappy.
    renderSheet(currentDaysForSave(), next, true);
  }

  function toggleRates() {
    ratesHiddenRef.current = !ratesHiddenRef.current;
    renderSheet(currentDaysForSave(), undefined, true);
  }

  function toggleTerms() {
    const next = !showTermsRef.current;
    showTermsRef.current = next;
    setShowTerms(next);
    renderSheet(currentDaysForSave(), undefined, true); // data unchanged → skip recompute
  }

  // Rebuild the trailing 8 total cells for every row after a column change,
  // placing theatre-level totals on each theatre's first row.
  function withRebuiltTotals(
    days: Day[],
    newSpl: number,
    mutate: (r: Cell[]) => Cell[]
  ): Day[] {
    const theatreRows = buildTheatreRows();
    const firstToTheatre: Record<number, number> = {};
    for (const [tid, info] of Object.entries(theatreRows)) {
      firstToTheatre[info.first] = Number(tid);
    }
    return days.map((d) => ({
      name: d.name,
      data: d.data.map((row, idx) => {
        const r = mutate([...row]);
        const tid = firstToTheatre[idx + 1];
        const totals =
          tid !== undefined
            ? theatreTotalsCells(newSpl, theatreRows[tid].rows)
            : EMPTY_TOTALS;
        r.splice(r.length - 8, 8, ...totals);
        return r;
      }),
    }));
  }

  // Remove the k-th special show (0-based): drop its 2 columns from every day.
  function removeSpecialShow(k: number) {
    const oldSpl = splCountRef.current;
    if (k < 0 || k >= oldSpl) return;
    const newSpl = oldSpl - 1;
    const removeAt = BASE_LEN + 2 * k;
    const days = withRebuiltTotals(currentDaysForSave(), newSpl, (r) => {
      r.splice(removeAt, 2);
      return r;
    });
    setSpl(newSpl);
    saveDays(days);
    renderSheet(days);
  }

  // Add another special show (Spl-N): insert its 2 columns before the fixed shows.
  function addSpecialShow() {
    const oldSpl = splCountRef.current;
    const newSpl = oldSpl + 1;
    const insertAt = BASE_LEN + 2 * oldSpl;
    const days = withRebuiltTotals(currentDaysForSave(), newSpl, (r) => {
      r.splice(insertAt, 0, "", "");
      return r;
    });
    setSpl(newSpl);
    saveDays(days);
    renderSheet(days);
  }


  function currentDaysForSave(): Day[] {
    return sheetsRef.current.map((ws, i) => ({
      name: namesRef.current[i] ?? `Day ${i + 1}`,
      data: ws.getData(),
    }));
  }

  function saveDays(days: Day[]) {
    // For uploaded (Univer) sheets, the snapshot can be many MB — too big for
    // localStorage's ~5MB quota. The SERVER (Postgres) is the source of truth
    // for those; localStorage only records the flag. Wrapped so a quota error
    // can never block an upload.
    try {
      const payload = plainModeRef.current
        ? {
            v: SHEET_VERSION,
            splCount: splCountRef.current,
            plain: true,
            univer: univerSnapRef.current, // deduped → usually fits
          }
        : {
            v: SHEET_VERSION,
            splCount: splCountRef.current,
            days,
            plain: false,
          };
      localStorage.setItem(saveKey, JSON.stringify(payload));
    } catch {
      // Snapshot too big for localStorage → keep just the flag; the SERVER copy
      // (no size limit) still persists via serverSave() and restores on reload.
      try {
        localStorage.setItem(
          saveKey,
          JSON.stringify({ v: SHEET_VERSION, splCount: splCountRef.current, plain: true })
        );
      } catch {}
    }
  }

  // Returns { splCount, days, plain, univer }; rebuilds fresh for legacy formats.
  function loadSaved(): {
    splCount: number;
    days: Day[];
    plain: boolean;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    univer?: any;
  } {
    try {
      const raw = localStorage.getItem(saveKey);
      if (raw) {
        const parsed = JSON.parse(raw);
        // Uploaded (plain / Univer) sheet — the big snapshot lives on the server;
        // localStorage only flags it. serverLoad() fills the snapshot on mount.
        if (parsed?.plain) {
          return {
            splCount:
              typeof parsed.splCount === "number" ? parsed.splCount : DEFAULT_SPL,
            days: [],
            plain: true,
            univer: parsed.univer, // usually undefined; server provides it
          };
        }
        if (
          parsed &&
          parsed.v === SHEET_VERSION &&
          Array.isArray(parsed.days) &&
          parsed.days.length
        ) {
          const sc =
            typeof parsed.splCount === "number" ? parsed.splCount : DEFAULT_SPL;
          return { splCount: sc, days: parsed.days as Day[], plain: false };
        }
      }
    } catch {}
    return {
      splCount: DEFAULT_SPL,
      days: [{ name: "Day 1", data: buildData(theatres, DEFAULT_SPL) }],
      plain: false,
    };
  }

  function persist() {
    try {
      saveDays(currentDaysForSave());
    } catch {}
    if (!pullingRef.current) {
      lastEditRef.current = Date.now();
      serverSave(); // debounced push to the shared server copy
    }
  }

  // ── Shared-sheet server sync ──
  // The full editable state (days + side-stores) as one JSON blob.
  function fullBlob() {
    const read = (k: string) => {
      try {
        return JSON.parse(localStorage.getItem(k) || "null");
      } catch {
        return null;
      }
    };
    return {
      v: SHEET_VERSION,
      splCount: splCountRef.current,
      plain: plainModeRef.current,
      univer: plainModeRef.current ? univerSnapRef.current : undefined,
      days: plainModeRef.current ? [] : currentDaysForSave(),
      deduction: read(`svf.deduction.${movieId}`),
      nettoverride: read(`svf.nettoverride.${movieId}`),
      showentry: read(`svf.showentry.${movieId}`),
    };
  }
  // Hydrate localStorage from a server blob, then re-render the sheet from it.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  // Clean up a Univer snapshot loaded from an OLDER converter: any cell value
  // that got stringified to "[object Object]" is reset to 0 so stale saved
  // copies display correctly without needing a re-upload.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function sanitizeUniver(snap: any): any {
    try {
      // Force MIDDLE vertical alignment everywhere so cell values are centered
      // (Univer draws styled cells that have no explicit vt at the TOP, so they
      // look like they "float"). Done at load time so EVERY movie — even ones
      // uploaded earlier — gets it without needing a re-upload. vt: 2 = MIDDLE.
      if (snap && typeof snap === "object") {
        snap.defaultStyle = { ...(snap.defaultStyle || {}), vt: 2 };
        const styles = snap.styles;
        if (styles && typeof styles === "object") {
          for (const id of Object.keys(styles)) {
            const st = styles[id];
            if (st && typeof st === "object") {
              st.vt = 2;
              // Strikethrough on report cells is copy-paste cruft (Excel files
              // accumulate a "bold+strike+underline" font). Drop the strike AND
              // the underline that rides on the same cruft font, but keep a
              // standalone underline (e.g. the report title).
              if (st.st) {
                delete st.st;
                delete st.ul;
              }
              // A filled cell hides the sheet gridlines, so give it a full
              // thin-black border — otherwise adjacent filled cells (the green
              // audience column) read as one solid block with no line between.
              if (st.bg) ensureFullBorder(st);
            }
          }
        }
      }
      const sheets = snap?.sheets;
      if (!sheets) return snap;
      for (const sid of Object.keys(sheets)) {
        const sheet = sheets[sid];
        if (sheet && typeof sheet === "object") {
          sheet.defaultStyle = { ...(sheet.defaultStyle || {}), vt: 2 };
          // Dark gridlines so every cell (incl. empty ones) shows a crisp border,
          // matching the black boxes on data cells. Applied at load so movies
          // uploaded before this change get it without a re-upload.
          sheet.showGridlines = 1;
          sheet.gridlinesColor = "#000000";
        }
        // Floor every row height so the 14pt report font isn't crammed into a
        // too-short row (values overlap top-to-bottom — the "puffed" look). Done
        // at load so movies uploaded before the converter fix render right too.
        const rowData = sheet?.rowData;
        if (rowData && typeof rowData === "object") {
          for (const rk of Object.keys(rowData)) {
            const rr = rowData[rk];
            if (rr && typeof rr === "object" && typeof rr.h === "number" && rr.h < 22)
              rr.h = 22;
          }
        }
        const cd = sheet?.cellData;
        if (!cd) continue;
        for (const r of Object.keys(cd)) {
          const row = cd[r];
          for (const c of Object.keys(row)) {
            const cell = row[c];
            if (cell && cell.v === "[object Object]") cell.v = 0;
            if (cell && typeof cell.v === "object") cell.v = 0;
            // Whitespace-only text ("" or " ") left by an external paste makes
            // "text + number" formulas evaluate to #VALUE!. Drop it (keeping any
            // style) so the formula computes 0, shown as "-" by the number format.
            if (
              cell &&
              !cell.f &&
              typeof cell.v === "string" &&
              cell.v.trim() === ""
            )
              delete cell.v;
            // Inline style object on the cell → force middle + strip strike cruft.
            if (cell && cell.s && typeof cell.s === "object") {
              cell.s.vt = 2;
              if (cell.s.st) {
                delete cell.s.st;
                delete cell.s.ul;
              }
              if (cell.s.bg) ensureFullBorder(cell.s);
            }
          }
        }
      }
    } catch {}
    return snap;
  }

  // inPlace=true → update the workbook WITHOUT remounting Univer (no blank
  // flash for other viewers when a remote edit arrives). Used by the poll.
  function applyBlob(blob: any, inPlace = false) {
    if (!blob) return;
    // Uploaded Excel (Univer) copy — swap the workbook snapshot.
    if (blob.plain && blob.univer) {
      blob.univer = sanitizeUniver(blob.univer);
      pullingRef.current = true;
      plainModeRef.current = true;
      setPlainMode(true);
      univerSnapRef.current = blob.univer;
      univerReadyRef.current = true;
      setUniverSnap(blob.univer);
      // Only remount (blank flash) for the initial load / upload, NOT for
      // remote collaborative updates — those apply in place inside UniverSheet.
      if (!inPlace) setUniverKey((k) => k + 1);
      try {
        saveDays([]);
      } catch {}
      window.setTimeout(() => (pullingRef.current = false), 150);
      return;
    }
    if (!Array.isArray(blob.days)) return;
    pullingRef.current = true;
    try {
      plainModeRef.current = !!blob.plain;
      setPlainMode(!!blob.plain);
      localStorage.setItem(
        saveKey,
        JSON.stringify({
          v: blob.v ?? SHEET_VERSION,
          splCount: blob.splCount ?? DEFAULT_SPL,
          days: blob.days,
          plain: !!blob.plain,
        })
      );
      if (blob.deduction)
        localStorage.setItem(`svf.deduction.${movieId}`, JSON.stringify(blob.deduction));
      if (blob.nettoverride)
        localStorage.setItem(`svf.nettoverride.${movieId}`, JSON.stringify(blob.nettoverride));
      if (blob.showentry)
        localStorage.setItem(`svf.showentry.${movieId}`, JSON.stringify(blob.showentry));
      const saved = loadSaved();
      setSpl(saved.splCount);
      renderSheet(saved.days);
      window.setTimeout(() => {
        if (!plainModeRef.current) {
          syncTerms();
          restoreShowEntries();
          recomputeNett();
          recomputeGT();
        }
        pullingRef.current = false;
      }, 150);
    } catch {
      pullingRef.current = false;
    }
  }
  // ── gzip compression (the uploaded snapshot can be ~10MB → too big to send
  // raw). Compress to base64 before saving; decompress on load. ──
  function bytesToB64(bytes: Uint8Array): string {
    let bin = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk)
      bin += String.fromCharCode(...bytes.subarray(i, i + chunk));
    return btoa(bin);
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  async function compress(obj: any): Promise<any> {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const CS = (globalThis as any).CompressionStream;
      if (!CS) return obj; // fallback: send raw
      const stream = new Blob([JSON.stringify(obj)])
        .stream()
        .pipeThrough(new CS("gzip"));
      const buf = await new Response(stream).arrayBuffer();
      return { __gz: bytesToB64(new Uint8Array(buf)) };
    } catch {
      return obj;
    }
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  async function decompress(data: any): Promise<any> {
    if (!data || !data.__gz) return data;
    try {
      // Fast native base64→bytes via a data: URL (avoids slow per-char decode).
      const bytes = await (
        await fetch("data:application/octet-stream;base64," + data.__gz)
      ).arrayBuffer();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const DS = (globalThis as any).DecompressionStream;
      const stream = new Blob([bytes]).stream().pipeThrough(new DS("gzip"));
      return JSON.parse(await new Response(stream).text());
    } catch {
      return null;
    }
  }

  // Save-status toast helpers.
  function markSaving() {
    setSaveState("saving");
    if (saveStateTimer.current) clearTimeout(saveStateTimer.current);
  }
  function markSaved() {
    setSaveState("saved");
    if (saveStateTimer.current) clearTimeout(saveStateTimer.current);
    saveStateTimer.current = window.setTimeout(() => setSaveState("idle"), 2000);
  }

  // Debounced save of the whole blob to the server (last-write-wins).
  function serverSave() {
    if (roleRef.current !== "editor") return;
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = window.setTimeout(async () => {
      markSaving();
      try {
        const res = await fetch(`/api/movies/${movieId}/sheet`, {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ data: await compress(fullBlob()) }),
        });
        if (res.ok) {
          serverVersionRef.current = (await res.json()).version ?? 0;
          markSaved();
        } else setSaveState("idle");
      } catch {
        setSaveState("idle");
      }
    }, 700);
  }
  // Force an immediate save (Ctrl+S / Enter-after-edit) — bypass the debounce.
  async function flushSave() {
    if (roleRef.current !== "editor") return;
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    markSaving();
    try {
      const res = await fetch(`/api/movies/${movieId}/sheet`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ data: await compress(fullBlob()) }),
      });
      if (res.ok) {
        serverVersionRef.current = (await res.json()).version ?? 0;
        markSaved();
      } else setSaveState("idle");
    } catch {
      setSaveState("idle");
    }
  }
  // Load the shared copy from the server. Returns true if applied.
  async function serverLoad(): Promise<boolean> {
    try {
      const res = await fetch(`/api/movies/${movieId}/sheet`, { cache: "no-store" });
      if (!res.ok) return false;
      const j = await res.json();
      roleRef.current = j.role === "viewer" ? "viewer" : "editor";
      setReadOnly(j.role === "viewer");
      serverVersionRef.current = j.version ?? 0;
      const data = await decompress(j.data);
      // Only the Univer spreadsheet is used now. Legacy OG-only saves (days,
      // no univer) are ignored → the movie opens as a blank Univer sheet.
      if (data && data.univer) {
        applyBlob(data);
        return true;
      }
    } catch {}
    return false;
  }

  function renderSheet(
    days: Day[],
    hiddenOverride?: boolean,
    skipRecompute = false,
    ratesOverride?: boolean
  ) {
    const el = ref.current;
    const jspreadsheet = jssRef.current;
    if (!el || !jspreadsheet) return;
    const hidden = hiddenOverride ?? splHiddenRef.current;
    const termsHidden = !showTermsRef.current;
    const ratesHidden = ratesOverride ?? ratesHiddenRef.current;
    const sc = splCountRef.current;
    const plain = plainModeRef.current; // uploaded Excel → plain generic grid
    // preserve scroll position across the rebuild (so collapse doesn't jump to top)
    const scroller = el.parentElement;
    const savedTop = scroller?.scrollTop ?? 0;
    const savedLeft = scroller?.scrollLeft ?? 0;
    // Remove the previous render's tab bar (it was moved into the slot, so
    // el.innerHTML="" below won't clear it) — prevents stacked tab bars.
    tabsSlotRef.current?.replaceChildren();
    el.innerHTML = "";
    namesRef.current = days.map((d) => d.name);
    const instance = jspreadsheet(el, {
      tabs: true,
      allowExport: false, // disable built-in Ctrl+S CSV download
      worksheets: days.map((d) => {
        if (plain) {
          // Uploaded Excel: show it faithfully — one generic column per data
          // column, no OG nested headers / fixed layout.
          const ncols = Math.max(
            1,
            d.data.reduce((m, r) => Math.max(m, r.length), 1)
          );
          return {
            data: d.data,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            columns: Array.from({ length: ncols }, () => ({ width: 120 })) as any,
            worksheetName: d.name,
            minDimensions: [ncols, Math.max(d.data.length, 1)],
            editable: true,
            columnResize: true,
            rowResize: true,
            columnSorting: false,
            allowInsertColumn: true,
            allowInsertRow: true,
            allowManualInsertColumn: true,
            tableOverflow: false,
          };
        }
        return {
          data: d.data,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          columns: columnsFor(sc) as any,
          worksheetName: d.name,
          minDimensions: [ncolsFor(sc), 1],
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          nestedHeaders: nestedHeadersFor(sc, hidden, termsHidden, ratesHidden) as any,
          editable: true,
          columnResize: true,
          rowResize: true,
          columnSorting: false, // so double-click renames the header instead of sorting
          allowInsertColumn: true,
          allowInsertRow: true,
          allowManualInsertColumn: true,
          allowRenameColumn: true,
          // The surrounding container scrolls (both axes); let the table render at
          // its natural size so horizontal + vertical scrolling always works.
          tableOverflow: false,
        };
      }),
      // Right-click a Show/Aud cell → add an "Add show entry…" option that
      // opens the per-class audience dialog.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      contextMenu: (_ws: any, x: any, y: any, _e: any, items: any[]) => {
        const xi = Number(x);
        const yi = Number(y);
        const nShows = showsFor(splCountRef.current).length;
        const inShows =
          Number.isFinite(yi) && xi >= BASE_LEN && xi < BASE_LEN + nShows * 2;
        // Only for single-screen theatres (not multiplex / band / spacer rows).
        const isSingle = buildRowTypes()[yi + 1] === "single";
        if (inShows && isSingle) {
          items.unshift({ type: "line" });
          items.unshift({
            title: "Add show entry…",
            onclick: () => openShowEntry(xi, yi),
          });
        }
        return items;
      },
      // Follow the selected cell with the scroll container when navigating by
      // arrow keys (tableOverflow is off, so jspreadsheet won't scroll itself).
      // Batched in rAF + early-out-when-visible so it never thrashes layout on
      // the large table during rapid key-repeat.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      onselection: (ws: any, _x1: any, _y1: any, x2: any, y2: any) => {
        if (followRaf.current) cancelAnimationFrame(followRaf.current);
        followRaf.current = requestAnimationFrame(() => {
          const scroller = el.parentElement;
          if (!scroller) return;
          const body = (ws?.tbody as HTMLElement | undefined) ?? undefined;
          const row = body?.children[Number(y2)] as HTMLElement | undefined;
          const td = row?.children[Number(x2) + 1] as HTMLElement | undefined;
          if (!td) return;
          const thead = ws?.thead as HTMLElement | undefined;
          const headH = thead ? thead.offsetHeight : 0;
          const pad = 4;
          // Single reflow: read all geometry, then decide whether to write.
          const c = scroller.getBoundingClientRect();
          const t = td.getBoundingClientRect();
          let dTop = 0;
          if (t.top < c.top + headH + pad) dTop = t.top - (c.top + headH + pad);
          else if (t.bottom > c.bottom - pad) dTop = t.bottom - (c.bottom - pad);
          let dLeft = 0;
          if (t.left < c.left + pad) dLeft = t.left - (c.left + pad);
          else if (t.right > c.right - pad) dLeft = t.right - (c.right - pad);
          if (dTop) scroller.scrollTop += dTop; // no-op write skipped when 0
          if (dLeft) scroller.scrollLeft += dLeft;
        });
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      onchange: (ws: any, _cell: any, x: number, y: number) => {
        // Ignore programmatic writes (recompute*/pull) entirely — they self-guard
        // and persist once at the end; only real user edits fall through here.
        if (applyingRef.current) return;
        applyingRef.current = true;
        try {
          handleCellChange(ws, Number(x), Number(y));
        } finally {
          applyingRef.current = false;
        }
        persist();
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      onafterchanges: () => {
        if (!applyingRef.current) persist();
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      oninsertcolumn: () => {
        persist();
        setTimeout(enhanceSheet, 30);
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      oninsertrow: () => persist(),
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      ondeletecolumn: () => {
        persist();
        setTimeout(enhanceSheet, 30);
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      ondeleterow: () => persist(),
    });
    sheetsRef.current = Array.isArray(instance) ? instance : [instance];
    // Hide columns SYNCHRONOUSLY (before the browser paints) so the grouped
    // header — built assuming these are hidden — always lines up. Doing this in
    // a delayed timeout caused a visible misalignment flash (esp. Terms).
    if (!plain)
      for (const ws of sheetsRef.current) {
        try {
          if (hidden) ws.hideColumn(splColIndexes(splCountRef.current));
          if (termsHidden) ws.hideColumn(TERMS_GROUP_COLS);
          if (ratesHidden) ws.hideColumn(PRICING_COL_INDEXES);
        } catch {}
      }
    // Move the fresh tab bar into the fixed slot immediately (synchronously),
    // so there's never a window where an old + new bar both exist.
    pinTabs();
    // restore scroll right away (before the browser paints the rebuilt table)
    if (scroller) {
      scroller.scrollTop = savedTop;
      scroller.scrollLeft = savedLeft;
    }
    setTimeout(() => {
      pinTabs();
      // Uploaded plain grid: no OG headers/formulas/styling — just show it.
      if (plain) {
        if (scroller) {
          scroller.scrollTop = savedTop;
          scroller.scrollLeft = savedLeft;
        }
        return;
      }
      if (filter !== "all") applyFilterTo(sheetsRef.current, filter);
      // Apply the per-day styling + collapse chevrons to EVERY worksheet (all
      // existing days and any newly added day), not just the active one, so
      // every tab — including brand-new days — looks & behaves identical.
      for (const ws of sheetsRef.current) {
        enhanceSheet(ws);
        stickyHeaders(ws);
        setSectionDividers(ws);
        styleTotalRows(ws);
        styleDistRows(ws);
      }
      if (!skipRecompute) {
        recomputeNett(); // per-class GST Nett
        recomputeGT(); // cumulative GT across all day tabs
      }
      // restore the scroll position captured before the rebuild
      if (scroller) {
        scroller.scrollTop = savedTop;
        scroller.scrollLeft = savedLeft;
      }
    }, 50);
  }

  // Red highlight on the section TOTAL rows (like the OG cinema grand-total line).
  // Operates on the ACTIVE worksheet (each day tab) so styling applies to all days.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function styleTotalRows(ws: any = activeWs()) {
    const tbody = ws?.tbody as HTMLElement | undefined;
    if (!tbody) return;
    const totals = new Set(totalRowNumbers(theatres));
    Array.from(tbody.children).forEach((tr, i) => {
      (tr as HTMLElement).classList.toggle("total-row", totals.has(i + 1));
    });
  }

  // District band rows ("DIST : …") → merge the whole row into ONE cell so the
  // district name reads clearly across the sheet (bold, prominent).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function styleDistRows(ws: any = activeWs()) {
    const tbody = ws?.tbody as HTMLElement | undefined;
    if (!tbody) return;
    const dist = new Set(distRowNumbers(theatres));
    Array.from(tbody.children).forEach((trEl, i) => {
      const tr = trEl as HTMLElement;
      const isDist = dist.has(i + 1);
      tr.classList.toggle("dist-row", isDist);
      const tds = Array.from(tr.children) as HTMLTableCellElement[];
      if (!isDist) {
        // undo a previous merge if this row is no longer a district band
        if (tds[1]?.colSpan > 1) {
          tds[1].colSpan = 1;
          tds.forEach((td) => (td.style.display = ""));
        }
        return;
      }
      // tds[0] = row-number gutter; tds[1] = first data cell (holds the label).
      const first = tds[1];
      if (!first) return;
      let span = 0;
      for (let k = 1; k < tds.length; k++) {
        span++;
        if (k > 1) tds[k].style.display = "none";
      }
      first.colSpan = span;
      first.style.display = "";
    });
  }

  // Make both header rows stick to the top while scrolling (active worksheet).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function stickyHeaders(ws: any = activeWs()) {
    const thead = ws?.thead as HTMLElement | undefined;
    if (!thead) return;
    const r1 = thead.querySelector("tr:first-child") as HTMLElement | null;
    const r2 = thead.querySelector("tr:last-child") as HTMLElement | null;
    const h1 = r1?.offsetHeight ?? 0;
    // The day tabs are moved out of the scroller (pinTabs), so the column
    // headers pin to the very top of the scroll area.
    r1?.querySelectorAll("td").forEach((td) => {
      const t = td as HTMLElement;
      t.style.position = "sticky";
      t.style.top = "0px";
      t.style.zIndex = "7";
    });
    if (r2 && r2 !== r1) {
      r2.querySelectorAll("td").forEach((td) => {
        const t = td as HTMLElement;
        t.style.position = "sticky";
        t.style.top = `${h1}px`;
        t.style.zIndex = "7";
      });
    }
  }

  // Move jspreadsheet's native day-tab bar out of the scrolling sheet and into
  // the fixed slot below the toolbar, so the tabs stay visible while scrolling.
  function pinTabs() {
    const el = ref.current;
    const slot = tabsSlotRef.current;
    if (!el || !slot) return;
    // Move the freshly-rendered tab bar into the slot (replaceChildren clears
    // any previous bar sitting there).
    const fresh = el.querySelector(".jtabs-headers") as HTMLElement | null;
    if (fresh && fresh.parentElement !== slot) slot.replaceChildren(fresh);
    // Safety: a re-render can leave the old instance's bar behind — remove any
    // stray tab bars so only the one in the slot survives (no stacking).
    document.querySelectorAll(".jtabs-headers").forEach((bar) => {
      if (bar.parentElement !== slot) bar.remove();
    });
    // Re-apply the per-day styling whenever a different day tab is clicked
    // (jspreadsheet has no tab-change event, so we listen on the bar once).
    if (slot.dataset.tabHook !== "1") {
      slot.dataset.tabHook = "1";
      slot.addEventListener("click", () => {
        // let jspreadsheet switch the active worksheet first, then style it
        setTimeout(() => enhanceRef.current(), 0);
        setTimeout(() => enhanceRef.current(), 60);
      });
    }
  }

  // Re-apply all per-day visual enhancements to the currently active worksheet.
  function enhanceActive() {
    const ws = activeWs();
    if (!ws) return;
    enhanceSheet(ws);
    stickyHeaders(ws);
    setSectionDividers(ws);
    styleTotalRows(ws);
    styleDistRows(ws);
    if (filter !== "all") applyFilterTo([ws], filter);
  }
  // keep the tab-click listener calling the freshest closure (current filter etc.)
  enhanceRef.current = enhanceActive;

  // Thin bold dividers between the major sections (core | Rates | shows | Today | GT).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function setSectionDividers(ws: any = activeWs()) {
    const el = ws?.table as HTMLElement | undefined;
    if (!el) return;
    const nShows = showsFor(splCountRef.current).length;
    // Only the major section boundaries (not between every show group).
    const edges = new Set<number>();
    edges.add(CORE_LEN - 1); // H.F.A → before Rates
    if (!ratesHiddenRef.current) edges.add(CORE_LEN + PRICING_PAIRS * 2 - 1); // → before meta
    edges.add(SHOWS_COUNT_COL_INDEX); // Shows → before the show groups
    edges.add(BASE_LEN + nShows * 2 - 1); // last show → before Today
    edges.add(BASE_LEN + nShows * 2 + 3); // last Today col → before GT
    el.querySelectorAll("td.sec-edge").forEach((td) => td.classList.remove("sec-edge"));
    edges.forEach((ix) => {
      // matches both body cells and header title cells (both carry data-x)
      el.querySelectorAll(`td[data-x="${ix}"]`).forEach((td) =>
        td.classList.add("sec-edge")
      );
    });
  }

  // Build & show the H.F.A class-breakdown card for a data row (0-based y).
  function showHfaCard(y: number, px: number, py: number) {
    const card = cardRef.current;
    const ws = activeWs();
    if (!card || !ws) return;
    const rows: string[] = [];
    let totalAud = 0;
    let totalGross = 0;
    let n = 0;
    for (let i = 0; i < PRICING_PAIRS; i++) {
      const rate = Number(ws.getValueFromCoords(CORE_LEN + 2 * i, y)) || 0;
      const aud = Number(ws.getValueFromCoords(CORE_LEN + 2 * i + 1, y)) || 0;
      if (rate <= 0 && aud <= 0) continue;
      n++;
      totalAud += aud;
      totalGross += rate * aud;
      rows.push(
        `<tr><td>Class ${n}</td><td>₹${inr(rate)}</td><td>${inr(aud)}</td><td>₹${inr(
          rate * aud
        )}</td></tr>`
      );
    }
    if (!n) {
      hideHfaCard();
      return;
    }
    // Rebuild the content only when the hovered row changes.
    if (cardRowRef.current !== y) {
      cardRowRef.current = y;
      const name = String(ws.getValueFromCoords(2, y) ?? ""); // Theatre/screen name
      card.innerHTML =
        `<h5>${name || "Pricing"} — ${n} class${n > 1 ? "es" : ""}</h5>` +
        `<table><thead><tr><th>Class</th><th>Rate</th><th>Aud</th><th>Gross</th></tr></thead>` +
        `<tbody>${rows.join("")}</tbody>` +
        `<tfoot><tr><td>Total</td><td></td><td>${inr(totalAud)}</td><td>₹${inr(
          totalGross
        )}</td></tr></tfoot></table>`;
    }
    card.style.display = "block";
    // position near the cursor, kept within the viewport
    const cw = card.offsetWidth || 260;
    const ch = card.offsetHeight || 160;
    let left = px + 14;
    let top = py + 14;
    if (left + cw > window.innerWidth - 8) left = px - cw - 14;
    if (top + ch > window.innerHeight - 8) top = window.innerHeight - ch - 8;
    card.style.left = `${Math.max(8, left)}px`;
    card.style.top = `${Math.max(8, top)}px`;
  }

  function hideHfaCard() {
    if (cardRef.current) cardRef.current.style.display = "none";
    cardRowRef.current = null;
    cardPinnedRef.current = false;
  }

  // Describe a derived cell (name + why). Returns null for plain input cells.
  function cellRole(x: number): { name: string; why: string } | null {
    const nShows = showsFor(splCountRef.current).length;
    const todayIdx = BASE_LEN + nShows * 2;
    const gtIdx = todayIdx + 4;
    if (x === 4) return { name: "H.F.C", why: "Σ (class RATE × class AUD) — house-full gross." };
    if (x === 5) return { name: "H.F.A", why: "Σ class AUD — house-full audience (all seats)." };
    if (x === SHOWS_COUNT_COL_INDEX)
      return { name: "Shows", why: "0.5 per filled Show + 0.5 per filled Aud." };
    if (x >= BASE_LEN && x < todayIdx) {
      if ((x - BASE_LEN) % 2 === 0)
        return { name: "Show (₹)", why: "RATE × Audience for this slot." };
      return null; // Audience is a manual input
    }
    if (x === todayIdx) return { name: "Today Gross", why: "Σ all Show cells of the theatre." };
    if (x === todayIdx + 1)
      return {
        name: "Today Nett",
        why: "Gross − GST − deduction. GST: 18% for tickets >₹105, 5% for ≤₹105.",
      };
    if (x === todayIdx + 2)
      return {
        name: "Today Share",
        why: "Percentage terms → Nett × %. Rent terms → Nett − (weekly rent ÷ 28) × Shows.",
      };
    if (x === todayIdx + 3)
      return { name: "Today Audience", why: "Σ all Aud cells of the theatre." };
    if (x >= gtIdx && x < gtIdx + 4) {
      const names = ["GT Gross", "GT Nett", "GT Share", "GT Audience"];
      return { name: names[x - gtIdx], why: "Cumulative across day tabs (previous GT + today)." };
    }
    return null;
  }

  // Build a live "worked example" line for the common derived cells.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function workedFor(ws: any, x: number, y: number): string {
    const nShows = showsFor(splCountRef.current).length;
    const todayIdx = BASE_LEN + nShows * 2;
    const v = (c: number) => Number(ws.getValueFromCoords(c, y)) || 0;
    const n = (x: number) => inr(Math.round(x));
    if (x === todayIdx + 1) return `${n(v(todayIdx))} ÷ 1.18 = ${n(v(todayIdx + 1))}`;
    if (x === todayIdx + 2) {
      const terms = String(ws.getValueFromCoords(TERMS_COL_INDEX, y) ?? "");
      const pct = topTermPct(terms);
      if (pct > 0)
        return `${n(v(todayIdx + 1))} × ${(pct * 100).toFixed(1)}% = ${n(v(todayIdx + 2))}`;
      const rent = weeklyRentOf(terms);
      if (rent > 0)
        return `${n(v(todayIdx + 1))} − (${inr(rent)}/28)×${n(
          v(SHOWS_COUNT_COL_INDEX)
        )} = ${n(v(todayIdx + 2))}`;
      return "";
    }
    if (x >= BASE_LEN && x < todayIdx && (x - BASE_LEN) % 2 === 0)
      return `RATE ${n(v(RATE_COL_INDEX))} × Aud ${n(v(x + 1))} = ${n(v(x))}`;
    if (x === 4) {
      const parts: string[] = [];
      for (let i = 0; i < PRICING_PAIRS; i++) {
        const r = v(CORE_LEN + 2 * i), a = v(CORE_LEN + 2 * i + 1);
        if (r > 0 || a > 0) parts.push(`${n(r)}×${n(a)}`);
      }
      return parts.length ? `${parts.join(" + ")} = ${n(v(4))}` : "";
    }
    // H.F.A — sum of class audiences
    if (x === 5) {
      const parts: string[] = [];
      for (let i = 0; i < PRICING_PAIRS; i++) {
        const a = v(CORE_LEN + 2 * i + 1);
        if (a > 0) parts.push(n(a));
      }
      return parts.length ? `${parts.join(" + ")} = ${n(v(5))}` : "";
    }
    // Today Gross / Today Audience — Σ of every Show / Aud cell of the theatre
    if (x === todayIdx || x === todayIdx + 3) {
      const tid = buildRowTheatre()[y + 1] ?? 0;
      const info = buildTheatreRows()[tid];
      if (!info) return "";
      const off = x === todayIdx ? 0 : 1;
      const parts: string[] = [];
      for (const r of info.rows)
        for (let k = 0; k < nShows; k++) {
          const val =
            Number(ws.getValueFromCoords(BASE_LEN + 2 * k + off, r - 1)) || 0;
          if (val) parts.push(n(val));
        }
      const total =
        x === todayIdx ? grossOfTheatre(ws, info) : audOfTheatre(ws, info);
      return parts.length ? `${parts.join(" + ")} = ${n(total)}` : "";
    }
    // Shows count — 0.5 per filled Show + 0.5 per filled Aud
    if (x === SHOWS_COUNT_COL_INDEX) {
      const tid = buildRowTheatre()[y + 1] ?? 0;
      const info = buildTheatreRows()[tid];
      if (!info) return "";
      let sFilled = 0,
        aFilled = 0;
      for (const r of info.rows)
        for (let k = 0; k < nShows; k++) {
          if (Number(ws.getValueFromCoords(BASE_LEN + 2 * k, r - 1))) sFilled++;
          if (Number(ws.getValueFromCoords(BASE_LEN + 2 * k + 1, r - 1)))
            aFilled++;
        }
      const total = Number(v(SHOWS_COUNT_COL_INDEX)).toString();
      return `${sFilled} Show×0.5 + ${aFilled} Aud×0.5 = ${total}`;
    }
    // GT columns — previous cumulative + today
    {
      const gtIdx = todayIdx + 4;
      if (x >= gtIdx && x < gtIdx + 4) {
        const today = v(todayIdx + (x - gtIdx));
        const total = v(x);
        return `previous ${n(total - today)} + today ${n(today)} = ${n(total)}`;
      }
    }
    return "";
  }

  // A1 column letters ("AE") → 0-based column index.
  function colLettersToIndex(letters: string): number {
    let n = 0;
    for (const ch of letters.toUpperCase())
      n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  // Rewrite a formula, replacing every cell reference (AE2, X5, …) with that
  // cell's actual computed number — the "values used" version.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function formulaWithValues(ws: any, formula: string): string {
    return formula.replace(/\$?([A-Za-z]+)\$?(\d+)/g, (m, col, row) => {
      const disp = ws.getValueFromCoords(
        colLettersToIndex(col),
        Number(row) - 1,
        true
      );
      const s = String(disp ?? "").replace(/[₹,\s]/g, "");
      const num = Number(s);
      return s !== "" && Number.isFinite(num) ? String(Math.round(num)) : m;
    });
  }

  // Open the inspector for a clicked derived cell (0-based coords).
  function openInspector(x: number, y: number) {
    const ws = activeWs();
    if (!ws) return;
    const theatre = String(ws.getValueFromCoords(2, y) ?? "").trim();
    const role = cellRole(x);
    if (!role) {
      // Not a calculated cell — if it holds a value, show it as a direct input
      // (so every populated cell explains itself); otherwise close the panel.
      const el0 = ref.current?.querySelector(`td[data-x="${x}"][data-y="${y}"]`);
      const val0 = el0
        ? (el0.textContent ?? "").trim()
        : String(ws.getValueFromCoords(x, y) ?? "").trim();
      if (!val0) {
        setInspect(null);
        return;
      }
      setInspect({
        x,
        y,
        label: `${theatre || `Row ${y + 1}`} — input`,
        why: "Direct input — you type this value; it isn't calculated from other cells.",
        worked: "",
        value: val0,
        formula: "",
        input: true,
      });
      return;
    }
    // show the DISPLAYED (computed) value, not the raw formula
    const cellEl = ref.current?.querySelector(`td[data-x="${x}"][data-y="${y}"]`);
    const value = cellEl
      ? (cellEl.textContent ?? "").trim()
      : String(ws.getValueFromCoords(x, y) ?? "");
    // raw formula (if any) lives in the worksheet data array
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const raw = (ws.options?.data as any)?.[y]?.[x];
    const formula = typeof raw === "string" && raw.startsWith("=") ? raw : "";

    // Nett cell → show the per-class GST breakdown + an editable deduction.
    const nShows = showsFor(splCountRef.current).length;
    const todayIdx = BASE_LEN + nShows * 2;
    let worked = workedFor(ws, x, y);
    let ded:
      | { tid: number; value: number; day: string; overridden: boolean }
      | undefined;
    if (x === todayIdx + 1) {
      const tid = buildRowTheatre()[y + 1] ?? 0;
      const day = activeDayName();
      const cls = classAudFor(day, tid);
      const info = buildTheatreRows()[tid];
      const gross = info ? grossOfTheatre(ws, info) : 0;
      const parts = Object.keys(cls).map((p) => {
        const price = Number(p);
        const rate = price > 105 ? "18%" : "5%";
        return `${price}×${inr(cls[price])} @${rate}`;
      });
      const gst = Object.keys(cls).length ? gstOfClasses(cls) : (gross * 18) / 118;
      const dv = deductionFor(tid);
      const overridden = loadNettOverrides()[`${day}|${tid}`] != null;
      worked = overridden
        ? "Manually overridden formula (auto-calc off for this cell)."
        : (parts.length ? `GST on: ${parts.join(", ")}` : "GST: flat 18%") +
          ` → GST ${inr(Math.round(gst))}, deduction ${inr(dv)}. Nett = ${inr(
            gross
          )} − ${inr(Math.round(gst))} − ${inr(dv)}`;
      ded = { tid, value: dv, day, overridden };
    }

    // Values-substituted version of the formula (cell refs → actual numbers).
    const values = formula ? formulaWithValues(ws, formula) : "";

    setInspect({
      x,
      y,
      label: `${theatre || `Row ${y + 1}`} — ${role.name}`,
      why: role.why,
      worked,
      value,
      formula,
      values,
      ded,
    });
  }

  // Apply an edited fixed deduction; clears any manual override so the auto
  // per-class formula (with the new deduction) takes over again.
  function applyDeduction(tid: number, value: number) {
    if (!inspect?.ded) return;
    saveDeduction(tid, value);
    clearNettOverride(inspect.ded.day, tid);
    recomputeNett(tid);
    recomputeGT(tid);
    persist();
    openInspector(inspect.x, inspect.y);
  }

  // Save a manual Nett formula override for this theatre/day (edit like Excel).
  function applyNettFormula(formula: string) {
    if (!inspect?.ded) return;
    let f = formula.trim();
    if (f && !f.startsWith("=")) f = "=" + f;
    saveNettOverride(inspect.ded.day, inspect.ded.tid, f);
    recomputeNett(inspect.ded.tid);
    recomputeGT(inspect.ded.tid);
    persist();
    openInspector(inspect.x, inspect.y);
  }

  // Drop the override → back to the auto per-class calculation.
  function resetNett() {
    if (!inspect?.ded) return;
    clearNettOverride(inspect.ded.day, inspect.ded.tid);
    recomputeNett(inspect.ded.tid);
    recomputeGT(inspect.ded.tid);
    persist();
    openInspector(inspect.x, inspect.y);
  }

  // Apply an edited formula from the inspector back into the cell.
  function applyInspectFormula(newFormula: string) {
    if (!inspect) return;
    const ws = activeWs();
    if (!ws) return;
    let f = newFormula.trim();
    if (f && !f.startsWith("=")) f = "=" + f;
    const prev = applyingRef.current;
    applyingRef.current = true;
    try {
      ws.setValueFromCoords(inspect.x, inspect.y, f, true);
    } catch {}
    applyingRef.current = prev;
    persist();
    const cellEl = ref.current?.querySelector(
      `td[data-x="${inspect.x}"][data-y="${inspect.y}"]`
    );
    const value = cellEl
      ? (cellEl.textContent ?? "").trim()
      : String(ws.getValueFromCoords(inspect.x, inspect.y) ?? "");
    setInspect({ ...inspect, formula: f, value, worked: workedFor(ws, inspect.x, inspect.y) });
  }

  // localStorage store for per-class show entries (so reopening restores them).
  function loadShowEntries(): Record<string, { price: string; aud: string }[]> {
    if (typeof window === "undefined") return {};
    try {
      return JSON.parse(localStorage.getItem(`svf.showentry.${movieId}`) || "{}");
    } catch {
      return {};
    }
  }
  function saveShowEntryRows(key: string, rows: { price: string; aud: string }[]) {
    const all = loadShowEntries();
    all[key] = rows;
    localStorage.setItem(`svf.showentry.${movieId}`, JSON.stringify(all));
  }

  function activeDayName(): string {
    const active = jssRef.current?.current;
    let idx = sheetsRef.current.findIndex((ws) => ws === active);
    if (idx < 0) idx = 0;
    return namesRef.current[idx] ?? "Day 1";
  }

  // Open the per-show entry dialog for a Show/Aud cell (0-based coords).
  function openShowEntry(colX: number, y: number) {
    const ws = activeWs();
    if (!ws) return;
    const group = Math.floor((colX - BASE_LEN) / 2);
    const showCol = BASE_LEN + 2 * group;
    const audCol = showCol + 1;
    const name = String(ws.getValueFromCoords(2, y) ?? "");
    const showName = showsFor(splCountRef.current)[group] ?? "Show";
    const cap = Number(ws.getValueFromCoords(5, y)) || 0;
    const tid = buildRowTheatre()[y + 1] ?? 0;
    const storeKey = `${activeDayName()}|${tid}|${name}|${group}`;
    // Restore a previously saved breakdown for this exact slot, if any…
    const saved = loadShowEntries()[storeKey];
    if (Array.isArray(saved) && saved.length) {
      setEntryRows(saved.map((r) => ({ price: String(r.price ?? ""), aud: String(r.aud ?? "") })));
    } else {
      // …otherwise pre-fill class prices from the rate card, blank audience.
      const rowsPrefill: { price: string; aud: string }[] = [];
      for (let i = 0; i < PRICING_PAIRS; i++) {
        const rate = Number(ws.getValueFromCoords(CORE_LEN + 2 * i, y)) || 0;
        if (rate > 0) rowsPrefill.push({ price: String(rate), aud: "" });
      }
      if (!rowsPrefill.length) rowsPrefill.push({ price: "", aud: "" });
      setEntryRows(rowsPrefill);
    }
    setEntry({ y, showCol, audCol, name, show: showName, cap, storeKey });
  }

  function saveShowEntry() {
    if (!entry) return;
    const ws = activeWs();
    if (!ws) return;
    const totalAud = entryRows.reduce((n, r) => n + (Number(r.aud) || 0), 0);
    const totalShow = entryRows.reduce(
      (n, r) => n + (Number(r.price) || 0) * (Number(r.aud) || 0),
      0
    );
    applyingRef.current = true;
    try {
      ws.setValueFromCoords(entry.audCol, entry.y, totalAud || "", true);
      ws.setValueFromCoords(entry.showCol, entry.y, totalShow || "", true);
      const tid = buildRowTheatre()[entry.y + 1];
      if (tid) {
        recomputeShows(ws, tid, buildTheatreRows());
      }
    } catch {}
    applyingRef.current = false;
    // remember the per-class breakdown so reopening this slot restores it
    saveShowEntryRows(entry.storeKey, entryRows);
    const tid2 = buildRowTheatre()[entry.y + 1];
    if (tid2) {
      recomputeNett(tid2); // uses the breakdown just saved
      recomputeGT(tid2);
    }
    persist();
    setEntry(null);
  }

  useEffect(() => {
    let destroyed = false;
    (async () => {
      const jspreadsheet = (await import("jspreadsheet-ce")).default;
      if (destroyed) return;
      jssRef.current = jspreadsheet;
      // Every movie is a Univer spreadsheet. Load its saved snapshot if any;
      // otherwise it stays a blank sheet until the user types or uploads Excel.
      plainModeRef.current = true;
      setPlainMode(true);
      const saved = loadSaved();
      if (saved.univer) {
        const clean = sanitizeUniver(saved.univer);
        univerSnapRef.current = clean;
        univerReadyRef.current = true;
        setUniverSnap(clean);
        setUniverKey((k) => k + 1);
      }
      // Pull the SHARED server copy; then allow saving (blank or loaded).
      setTimeout(async () => {
        if (destroyed) return;
        await serverLoad();
        univerReadyRef.current = true; // edits (even on a blank sheet) now save
        setLoading(false);
      }, 300);
      // Safety: never spin forever if the network hangs.
      setTimeout(() => !destroyed && setLoading(false), 6000);
    })();

    // Poll the shared copy; pull in others' edits (last-write-wins). Guarded so
    // it can never repeatedly rebuild the sheet: skip while hidden, while a pull
    // is in flight, and right after a local edit; bump the version BEFORE
    // applying so a slow apply can't re-trigger itself.
    const pollIv = setInterval(async () => {
      if (destroyed || document.hidden) return;
      if (pullingRef.current) return;
      if (Date.now() - lastEditRef.current < 3000) return;
      try {
        const res = await fetch(`/api/movies/${movieId}/sheet?meta=1`, {
          cache: "no-store",
        });
        if (!res.ok) return;
        const j = await res.json();
        if ((j.version ?? 0) > serverVersionRef.current) {
          serverVersionRef.current = j.version ?? 0; // claim it first
          pullingRef.current = true;
          const full = await (
            await fetch(`/api/movies/${movieId}/sheet`, { cache: "no-store" })
          ).json();
          const fdata = await decompress(full.data);
          if (fdata && fdata.univer) applyBlob(fdata, true); // in-place, no blank
          else pullingRef.current = false;
        }
      } catch {
        pullingRef.current = false;
      }
    }, 2000); // poll for collaborators' edits (cheap version check every 2s)

    // Hover the H.F.A cell (logical column 5) → show the class breakdown card.
    // Click it → pin the card open until you click elsewhere.
    const el = ref.current;
    const hfaY = (e: MouseEvent): number | null => {
      const td = (e.target as HTMLElement)?.closest?.("td");
      if (!td || td.getAttribute("data-x") !== "5") return null;
      const y = Number(td.getAttribute("data-y"));
      return Number.isFinite(y) ? y : null;
    };
    const onOver = (e: MouseEvent) => {
      if (cardPinnedRef.current) return; // pinned → hover doesn't move it
      const y = hfaY(e);
      if (y === null) {
        hideHfaCard();
        return;
      }
      showHfaCard(y, e.clientX, e.clientY);
    };
    const onLeave = () => {
      if (!cardPinnedRef.current) hideHfaCard();
    };
    const onClick = (e: MouseEvent) => {
      // H.F.A cell → pin the class-breakdown card
      const y = hfaY(e);
      if (y !== null) {
        cardPinnedRef.current = false;
        cardRowRef.current = null;
        showHfaCard(y, e.clientX, e.clientY);
        cardPinnedRef.current = true;
        return;
      }
      hideHfaCard();
      // any other derived cell → open the formula inspector
      const td = (e.target as HTMLElement)?.closest?.("td");
      const dx = Number(td?.getAttribute("data-x"));
      const dy = Number(td?.getAttribute("data-y"));
      if (Number.isFinite(dx) && Number.isFinite(dy)) openInspector(dx, dy);
    };
    el?.addEventListener("mousemove", onOver);
    el?.addEventListener("mouseleave", onLeave);
    el?.addEventListener("click", onClick);

    // Ctrl+S / Cmd+S → save immediately to the DB (works in both sheet modes).
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && (e.key === "s" || e.key === "S")) {
        e.preventDefault();
        if (plainModeRef.current) flushSave();
        else {
          persist();
          flushSave();
        }
      }
    };
    window.addEventListener("keydown", onKey, true);

    // Persist the user's position on tab-hide / navigation-away too, so an abrupt
    // close (before the 1s tracker fires) still remembers where they were.
    const onLeavePage = () => {
      try {
        if (univerApiRef.current) saveViewState(univerApiRef.current);
      } catch {}
    };
    window.addEventListener("beforeunload", onLeavePage);
    document.addEventListener("visibilitychange", onLeavePage);

    return () => {
      destroyed = true;
      clearInterval(pollIv);
      onLeavePage(); // save final position before this movie view tears down
      if (viewIvRef.current) clearInterval(viewIvRef.current);
      viewIvRef.current = null;
      if (viewSettleIvRef.current) clearInterval(viewSettleIvRef.current);
      viewSettleIvRef.current = null;
      window.removeEventListener("keydown", onKey, true);
      window.removeEventListener("beforeunload", onLeavePage);
      document.removeEventListener("visibilitychange", onLeavePage);
      el?.removeEventListener("mousemove", onOver);
      el?.removeEventListener("mouseleave", onLeave);
      el?.removeEventListener("click", onClick);
      if (ref.current) ref.current.innerHTML = "";
      sheetsRef.current = [];
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [movieId]);

  function currentDays(): Day[] {
    return sheetsRef.current.map((ws, i) => ({
      name: namesRef.current[i] ?? `Day ${i + 1}`,
      data: ws.getData(),
    }));
  }

  function nextDayName(days: Day[]): string {
    let max = 0;
    for (const d of days) {
      const m = /^Day\s+(\d+)$/i.exec(d.name.trim());
      if (m) max = Math.max(max, Number(m[1]));
    }
    return `Day ${max + 1 || days.length + 1}`;
  }

  function addDay() {
    const days = currentDays();
    days.push({ name: nextDayName(days), data: buildData(theatres, splCountRef.current) });
    saveDays(days);
    // New day starts with Rates & Special-shows COLLAPSED (user can expand).
    splHiddenRef.current = true;
    setSplHidden(true);
    ratesHiddenRef.current = true;
    renderSheet(days, true, false, true); // force spl + rates collapsed
  }

  function deleteDayAt(idx: number) {
    if (sheetsRef.current.length <= 1) {
      alert("At least one day is required.");
      return;
    }
    const name = namesRef.current[idx] ?? `Day ${idx + 1}`;
    if (!confirm(`Delete "${name}"? Its data will be lost.`)) return;
    const days = currentDays();
    days.splice(idx, 1);
    saveDays(days);
    renderSheet(days);
  }

  function deleteDay() {
    const active = jssRef.current?.current;
    let idx = sheetsRef.current.findIndex((ws) => ws === active);
    if (idx < 0) idx = sheetsRef.current.length - 1;
    deleteDayAt(idx);
  }

  // Inline rename of a column header (double-click), and ×-on-hover on day tabs.
  // Runs on the given worksheet's header (defaults to the active day) so the
  // collapse chevrons appear on EVERY day, not just Day 1.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function enhanceSheet(ws: any = activeWs()) {
    const thead = ws?.thead as HTMLElement | undefined;

    // Group-header enhancements on the nested (super) row:
    //  • ▾ chevron to collapse specials (on Spl-1), ▸ to expand (on Noon)
    //  • × on hover to remove a specific special show
    const nestedRow = thead?.querySelector("tr");
    if (nestedRow) {
      const cells = Array.from(nestedRow.children) as HTMLElement[];
      const collapsed = splHiddenRef.current;

      const addChevron = (
        cell: HTMLElement,
        sym: string,
        title: string,
        onClick: () => void,
        variant: "rates" | "spl"
      ) => {
        if (cell.querySelector(".spl-collapse")) return;
        const ch = document.createElement("span");
        const isPill = sym.includes("▸"); // expand affordance → solid pill
        ch.className = `spl-collapse spl-collapse--${variant}${isPill ? " spl-collapse--pill" : ""}`;
        ch.textContent = sym;
        ch.title = title;
        ch.addEventListener("mousedown", (e) => e.stopPropagation());
        ch.addEventListener("click", (e) => {
          e.stopPropagation();
          onClick();
        });
        cell.insertBefore(ch, cell.firstChild);
      };

      // Special-shows collapse: ▾ on Spl-1; when collapsed, ▸ Spl pill on the
      // meta cell (the empty cell immediately before the first show group),
      // so it sits in its own cell — not glued to "Noon".
      if (!collapsed) {
        const spl1 = cells.find((c) => /^Spl-1/.test((c.textContent ?? "").trim()));
        if (spl1) addChevron(spl1, "▾ ", "Collapse special shows", toggleSpecialShows, "spl");
      } else {
        const firstShowIdx = cells.findIndex((c) =>
          /^(Noon|Spl-1)/.test((c.textContent ?? "").trim())
        );
        const meta = firstShowIdx > 0 ? cells[firstShowIdx - 1] : undefined;
        if (meta) addChevron(meta, "▸ Spl", "Show special shows", toggleSpecialShows, "spl");
      }

      // Rates (pricing) collapse (green): ▾ on "Rates", ▸ Rates pill on the
      // core block (S.NO … H.F.A) when hidden.
      if (!ratesHiddenRef.current) {
        const rates = cells.find((c) => /^Rates/.test((c.textContent ?? "").trim()));
        if (rates) addChevron(rates, "▾ ", "Collapse rate columns", toggleRates, "rates");
      } else {
        const core = cells.find((c) => c.getAttribute("colspan") === String(CORE_LEN));
        if (core) addChevron(core, "▸ Rates", "Show rate columns", toggleRates, "rates");
      }

      // × to remove each special show
      let k = 0;
      cells.forEach((c) => {
        const txt = (c.textContent ?? "").trim();
        if (/Spl-\d+/.test(txt)) {
          const myK = k++;
          if (c.querySelector(".spl-remove")) return;
          const x = document.createElement("span");
          x.className = "spl-remove";
          x.textContent = "×";
          x.title = "Remove this special show";
          x.addEventListener("mousedown", (e) => e.stopPropagation());
          x.addEventListener("click", (e) => {
            e.stopPropagation();
            removeSpecialShow(myK);
          });
          c.appendChild(x);
        }
      });
    }

    // header double-click → inline rename (only the real column-title row,
    // not the grouped "nested" row above it)
    thead?.querySelectorAll("tr:last-child td").forEach((tdEl, i) => {
      if (i === 0) return; // row-number corner
      const td = tdEl as HTMLElement & { dataset: DOMStringMap };
      if (td.dataset.renameBound) return;
      td.dataset.renameBound = "1";
      td.addEventListener("dblclick", (e) => {
        e.stopPropagation();
        e.preventDefault();
        const ws = activeWs();
        if (!ws) return;
        const col = i - 1;
        const current = ws.getHeader ? ws.getHeader(col) : td.textContent ?? "";
        const prev = td.innerHTML;
        const input = document.createElement("input");
        input.value = current || "";
        input.style.width = "92%";
        input.style.font = "inherit";
        input.style.textAlign = "center";
        td.textContent = "";
        td.appendChild(input);
        input.focus();
        input.select();
        let done = false;
        const finish = (save: boolean) => {
          if (done) return;
          done = true;
          if (save) {
            try {
              ws.setHeader(col, input.value);
            } catch {
              td.innerHTML = prev;
            }
          } else {
            td.innerHTML = prev;
          }
        };
        input.addEventListener("blur", () => finish(true));
        input.addEventListener("keydown", (ev) => {
          if (ev.key === "Enter") {
            ev.preventDefault();
            input.blur();
          } else if (ev.key === "Escape") {
            finish(false);
          }
        });
      });
    });

    // × close button on each day tab (shown on hover via CSS) — the tab bar now
    // lives in the fixed slot, not inside the sheet container.
    const headers = tabsSlotRef.current?.querySelector(".jtabs-headers");
    if (headers) {
      const skip = ["jtabs-add", "jtabs-border", "jtabs-controls", "jtabs-prev", "jtabs-next"];
      let idx = 0;
      Array.from(headers.children).forEach((child) => {
        const tab = child as HTMLElement;
        if (skip.some((c) => tab.classList.contains(c))) return;
        const myIdx = idx++;
        if (tab.querySelector(".day-close")) return;
        const x = document.createElement("span");
        x.className = "day-close";
        x.textContent = "×";
        x.title = "Delete this day";
        x.addEventListener("mousedown", (e) => e.stopPropagation());
        x.addEventListener("click", (e) => {
          e.stopPropagation();
          deleteDayAt(myIdx);
        });
        tab.appendChild(x);
      });
    }
  }

  function activeWs() {
    return jssRef.current?.current ?? sheetsRef.current[0];
  }

  function addColumn() {
    const ws = activeWs();
    if (!ws) return;
    const selected: number[] = ws.getSelectedColumns?.() ?? [];
    if (selected.length > 0) ws.insertColumn(1, Math.min(...selected), true);
    else ws.insertColumn(1);
  }

  function removeColumn() {
    const ws = activeWs();
    if (!ws) return;
    const selected: number[] = ws.getSelectedColumns?.() ?? [];
    const totalCols = ws.getData?.()?.[0]?.length ?? 0;
    if (selected.length > 0) ws.deleteColumn(Math.min(...selected), selected.length);
    else if (totalCols > 0) ws.deleteColumn(totalCols - 1, 1);
  }

  // Map (theatreId|screenName) -> jspreadsheet row number (1-based), mirroring buildData.
  function buildRowMap() {
    const map: Record<string, number> = {};
    computeLayout(theatres).forEach((r, i) => {
      if (r.kind === "screen") map[`${r.tId}|${r.name}`] = i + 1;
    });
    return map;
  }

  // Pull representative submissions into the active day's show cells.
  async function pullRepData(silent = false) {
    const ws = activeWs();
    if (!ws) return;
    let reports: Array<{
      theatreId: string;
      screenName: string;
      time: string;
      totalSold: number;
      actualCollection: number;
    }> = [];
    try {
      const res = await fetch("/api/reports/all");
      if (!res.ok) throw new Error();
      reports = (await res.json()).reports || [];
    } catch {
      if (!silent) alert("Could not load representative data.");
      return;
    }
    const rowMap = buildRowMap();
    const rowTheatre = buildRowTheatre();
    const theatreRows = buildTheatreRows();
    const affected = new Set<number>();
    let filled = 0;
    applyingRef.current = true;
    for (const r of reports) {
      const rowNum = rowMap[`${r.theatreId}|${r.screenName}`];
      const slot = showsFor(splCountRef.current).indexOf(r.time);
      if (!rowNum || slot < 0) continue;
      const collCol = BASE_LEN + 2 * slot; // Show (₹)
      const audCol = collCol + 1; // Aud
      try {
        // set audience from the rep; Show auto-computes = RATE × Aud
        ws.setValueFromCoords(audCol, rowNum - 1, r.totalSold ?? 0, true);
        computeShow(ws, rowNum - 1, collCol);
        const tid = rowTheatre[rowNum];
        if (tid) affected.add(tid);
        filled++;
      } catch {}
    }
    affected.forEach((tid) => recomputeShows(ws, tid, theatreRows));
    applyingRef.current = false;
    recomputeNett(); // per-class GST Nett
    recomputeGT(); // refresh cumulative GT after pulling rep data
    persist();
    if (!silent) {
      alert(
        filled > 0
          ? `Filled ${filled} show entr${filled === 1 ? "y" : "ies"} from representatives into this day.`
          : "No matching representative entries found for these theatres/shows."
      );
    }
  }

  function resetSheet() {
    if (
      !confirm(
        "Clear this sheet to a blank spreadsheet? All current content will be lost."
      )
    )
      return;
    // Clear to a blank spreadsheet.
    plainModeRef.current = true;
    setPlainMode(true);
    const blank = {
      id: "wb-" + Math.random().toString(36).slice(2),
      sheetOrder: ["s1"],
      sheets: {
        s1: { id: "s1", name: "Sheet1", cellData: {}, rowCount: 500, columnCount: 52 },
      },
      styles: {},
      resources: [],
    };
    univerSnapRef.current = blank;
    univerReadyRef.current = true;
    setUniverSnap(blank);
    setUniverKey((k) => k + 1);
    saveDays([]);
    lastEditRef.current = Date.now();
    serverSave();
  }

  // Resolve the movie name from the shared DB list (works on a direct deep-link
  // load where the in-memory cache is still empty).
  const [movieName, setMovieName] = useState("Movie");
  useEffect(() => {
    const m = getMovie(movieId);
    if (m) {
      setMovieName(m.name);
      return;
    }
    refreshMovies().then((list) => {
      const found = list.find((x) => x.id === movieId);
      if (found) setMovieName(found.name);
    });
  }, [movieId]);

  useEffect(() => {
    fetch("/api/whoami", { cache: "no-store" })
      .then((r) => r.json())
      .then(setMe)
      .catch(() => {});
  }, []);

  async function exportExcel() {
    const XLSX = await import("xlsx");
    const wb = XLSX.utils.book_new();
    currentDays().forEach((d) => {
      const ncols = d.data.reduce((m, r) => Math.max(m, r.length), ncolsFor(splCountRef.current));
      const aoa = [titlesFor(splCountRef.current, ncols), ...d.data];
      const sheet = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(wb, sheet, sanitizeSheetName(d.name));
    });
    XLSX.writeFile(wb, `${movieName}.xlsx`);
  }

  // Upload an .xlsx → parse WITH styling (ExcelJS) → render faithfully in Univer.
  async function importExcel(file: File) {
    if (roleRef.current !== "editor") {
      alert("You have view-only access to this movie.");
      return;
    }
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ExcelJSmod: any = await import("exceljs");
      const ExcelJS = ExcelJSmod.default ?? ExcelJSmod;
      const buf = await file.arrayBuffer();
      // Read the file's real theme palette so fills/fonts match exactly.
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const JSZipMod: any = await import("jszip");
        const JSZip = JSZipMod.default ?? JSZipMod;
        const zip = await JSZip.loadAsync(buf);
        const themeXml = await zip.file("xl/theme/theme1.xml")?.async("string");
        if (themeXml) setThemePalette(parseThemePalette(themeXml));
      } catch {}
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb: any = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      const snap = excelToUniverSnapshot(wb, movieName || "Uploaded");
      // Switch into full-Excel (Univer) mode and render the uploaded file.
      plainModeRef.current = true;
      setPlainMode(true);
      univerSnapRef.current = snap;
      univerReadyRef.current = true;
      setUniverSnap(snap);
      setUniverKey((k) => k + 1);
      saveDays(currentDaysForSave()); // persist plain flag + snapshot (localStorage)
      lastEditRef.current = Date.now();
      serverSave();
    } catch (e) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const msg = (e as any)?.message || String(e);
      alert(
        "Could not read that file: " +
          msg +
          "\n\nPlease upload a valid .xlsx file (.xls / .csv are not supported)."
      );
    }
  }

  // Univer reported an edit → keep the snapshot and push to the shared server copy.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function onUniverChange(snap: any) {
    if (roleRef.current !== "editor") return;
    // Don't let the initial EMPTY workbook (before the real snapshot loads)
    // overwrite the saved data.
    if (!univerReadyRef.current) return;
    univerSnapRef.current = snap;
    lastEditRef.current = Date.now();
    serverSave();
  }

  // Columns whose header reads "Spl - N" (special shows) on a given sheet snapshot.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function splColsOf(sheetSnap: any): number[] {
    const cd = sheetSnap?.cellData;
    if (!cd) return [];
    const cols = new Set<number>();
    for (const r of Object.keys(cd)) {
      if (Number(r) > 40) continue; // headers live near the top
      const row = cd[r];
      for (const c of Object.keys(row)) {
        const v = row[c]?.v;
        if (typeof v === "string" && /^\s*spl\s*-/i.test(v)) cols.add(Number(c));
      }
    }
    return [...cols];
  }

  // Whether the current workbook has any special-show columns (→ show the toggle).
  function hasSplColumns(): boolean {
    const snap = univerSnapRef.current;
    if (!snap?.sheets) return false;
    return Object.values(snap.sheets).some(
      (s) => splColsOf(s).length > 0
    );
  }

  // Show/hide the special-show columns across every sheet (day tab).
  function setSplVisibility(hidden: boolean) {
    const api = univerApiRef.current;
    const snap = univerSnapRef.current;
    if (!api || !snap?.sheets) return;
    try {
      const wb = api.getActiveWorkbook?.();
      if (!wb) return;
      const sheets = wb.getSheets?.() || [];
      for (const ws of sheets) {
        const id = ws.getSheetId?.();
        const cols = splColsOf(snap.sheets[id]);
        for (const c of cols) {
          if (hidden) ws.hideColumns?.(c, 1);
          else ws.showColumns?.(c, 1);
        }
      }
    } catch {}
  }

  function toggleSpl() {
    const next = !univerSplHidden;
    setUniverSplHidden(next);
    setSplVisibility(next);
  }

  // ── Per-user view state: last active sheet + scroll + selection ──
  // Stored per movie in localStorage so a refresh / reopen returns the user to
  // exactly where they were (their own browser only — never shared/synced).
  const viewKey = `svf-view:${movieId}`;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function currentViewState(api: any): any | null {
    try {
      const wb = api?.getActiveWorkbook?.();
      const ws = wb?.getActiveSheet?.();
      if (!wb || !ws) return null;
      return {
        sheetId: ws.getSheetId?.() ?? null,
        scroll: ws.getScrollState?.() ?? null,
        sel: wb.getActiveRange?.()?.getA1Notation?.() ?? null,
      };
    } catch {
      return null;
    }
  }
  function readViewState(): any | null {
    try {
      const s = localStorage.getItem(viewKey);
      return s ? JSON.parse(s) : null;
    } catch {
      return null;
    }
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function saveViewState(api: any) {
    const st = currentViewState(api);
    if (!st || (st.sel == null && st.scroll == null)) return;
    const s = JSON.stringify(st);
    if (s === lastViewRef.current) return;
    lastViewRef.current = s;
    try {
      localStorage.setItem(viewKey, s);
    } catch {}
  }
  // Restore the saved position (or A1 if none), then begin tracking so the next
  // reload lands here. Applied repeatedly for a short window because hiding
  // columns makes Univer auto-select them slightly after load.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function restoreAndTrackView(api: any) {
    const saved = readViewState();
    const start = Date.now();
    if (viewSettleIvRef.current) clearInterval(viewSettleIvRef.current);
    const iv = window.setInterval(() => {
      try {
        const wb = api.getActiveWorkbook?.();
        if (wb) {
          // Return to the saved sheet tab.
          if (saved?.sheetId && wb.getSheetBySheetId?.(saved.sheetId)) {
            const s = wb.getSheetBySheetId(saved.sheetId);
            if (s && s.getSheetId?.() !== wb.getActiveSheet?.()?.getSheetId?.())
              wb.setActiveSheet?.(s);
          }
          const ws = wb.getActiveSheet?.();
          if (ws) {
            const wantSel = saved?.sel || "A1";
            const curSel = wb.getActiveRange?.()?.getA1Notation?.();
            if (curSel !== wantSel) {
              try {
                wb.setActiveRange(ws.getRange(wantSel));
              } catch {}
            }
            if (saved?.scroll) {
              try {
                ws.scrollToCell?.(
                  saved.scroll.sheetViewStartRow || 0,
                  saved.scroll.sheetViewStartColumn || 0
                );
              } catch {}
            }
          }
        }
      } catch {}
      // After the settle window, hand control back to the user and start saving.
      if (Date.now() - start > 1400) {
        clearInterval(iv);
        viewSettleIvRef.current = null;
        lastViewRef.current = JSON.stringify(currentViewState(api) || {});
        if (viewIvRef.current) clearInterval(viewIvRef.current);
        viewIvRef.current = window.setInterval(
          () => saveViewState(api),
          1000
        ) as unknown as number;
      }
    }, 200) as unknown as number;
    viewSettleIvRef.current = iv;
  }

  // ── Find bar helpers: live-search the ACTIVE sheet and jump to matches ──
  // Searches the current sheet's cell values/formulas (case-insensitive), like a
  // browser find, and returns the matching cells in reading order.
  function computeFindMatches(query: string): { r: number; c: number }[] {
    const api = univerApiRef.current;
    const snap = univerSnapRef.current;
    if (!api || !snap?.sheets || !query.trim()) return [];
    let sid: string | undefined;
    try {
      sid = api.getActiveWorkbook?.()?.getActiveSheet?.()?.getSheetId?.();
    } catch {}
    if (!sid) return [];
    const cd = snap.sheets[sid]?.cellData || {};
    const q = query.toLowerCase();
    const out: { r: number; c: number }[] = [];
    for (const r of Object.keys(cd)) {
      const row = cd[r];
      for (const c of Object.keys(row)) {
        const cell = row[c];
        const text =
          cell?.v != null ? String(cell.v) : cell?.f != null ? String(cell.f) : "";
        if (text && text.toLowerCase().includes(q))
          out.push({ r: Number(r), c: Number(c) });
      }
    }
    out.sort((a, b) => a.r - b.r || a.c - b.c);
    return out;
  }
  function jumpToMatch(m: { r: number; c: number } | undefined) {
    if (!m) return;
    const api = univerApiRef.current;
    try {
      const wb = api.getActiveWorkbook?.();
      const ws = wb?.getActiveSheet?.();
      if (wb && ws) {
        wb.setActiveRange(ws.getRange(m.r, m.c)); // select (highlights) the match
        ws.scrollToCell?.(m.r, m.c); // and bring it into view
      }
    } catch {}
  }
  // Runs on every keystroke — recompute matches and jump to the first one.
  function runFind(query: string) {
    setFindQuery(query);
    const matches = computeFindMatches(query);
    setFindMatches(matches);
    if (matches.length) {
      setFindIdx(0);
      jumpToMatch(matches[0]);
    } else {
      setFindIdx(-1);
    }
  }
  function stepFind(dir: 1 | -1) {
    if (!findMatches.length) return;
    const n = (findIdx + dir + findMatches.length) % findMatches.length;
    setFindIdx(n);
    jumpToMatch(findMatches[n]);
  }
  function openFindBar() {
    setFindOpen(true);
    setTimeout(() => {
      findInputRef.current?.focus();
      findInputRef.current?.select();
    }, 0);
  }
  function closeFindBar() {
    setFindOpen(false);
  }

  // Import a local Excel INTO the current workbook — its sheets are appended as
  // new tabs so the user can keep working (not a replace).
  async function appendExcel(file: File) {
    if (roleRef.current !== "editor") {
      alert("You have view-only access to this movie.");
      return;
    }
    if (!plainModeRef.current || !univerSnapRef.current) {
      // Not in Excel mode yet → first upload just loads it.
      importExcel(file);
      return;
    }
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ExcelJSmod: any = await import("exceljs");
      const ExcelJS = ExcelJSmod.default ?? ExcelJSmod;
      const buf = await file.arrayBuffer();
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const JSZipMod: any = await import("jszip");
        const JSZip = JSZipMod.default ?? JSZipMod;
        const zip = await JSZip.loadAsync(buf);
        const themeXml = await zip.file("xl/theme/theme1.xml")?.async("string");
        if (themeXml) setThemePalette(parseThemePalette(themeXml));
      } catch {}
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb: any = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      const incoming = excelToUniverSnapshot(wb, file.name);
      // Merge incoming sheets into the current workbook snapshot (be tolerant of
      // a live snapshot that lacks/renames sheetOrder).
      const cur = JSON.parse(JSON.stringify(univerSnapRef.current));
      cur.sheets = cur.sheets || {};
      cur.styles = cur.styles || {};
      if (!Array.isArray(cur.sheetOrder))
        cur.sheetOrder = Object.keys(cur.sheets);

      // CRITICAL: both workbooks use numeric style ids ("1","2",…). Merge the
      // incoming styles under FRESH ids and remap every imported cell's `s`
      // reference — otherwise the new sheet points at the wrong (or missing)
      // styles and Univer can fail to build the workbook (nothing shows up).
      const incStyles = incoming.styles || {};
      const styleRemap: Record<string, string> = {};
      let styleSeq = 0;
      const freshStyleId = () => {
        let id: string;
        do {
          id = "imp-s" + ++styleSeq;
        } while (cur.styles[id]);
        return id;
      };
      for (const oldId of Object.keys(incStyles)) {
        const nid = freshStyleId();
        styleRemap[oldId] = nid;
        cur.styles[nid] = incStyles[oldId];
      }

      const names = new Set<string>(
        cur.sheetOrder.map((id: string) => cur.sheets[id]?.name)
      );
      const added: string[] = [];
      let seq = 0;
      for (const sid of incoming.sheetOrder) {
        const sheet = incoming.sheets[sid];
        // Point this sheet's cells at the remapped style ids.
        const cd = sheet.cellData || {};
        for (const r of Object.keys(cd)) {
          const rowCells = cd[r];
          for (const c of Object.keys(rowCells)) {
            const cell = rowCells[c];
            if (cell && cell.s != null && styleRemap[cell.s] != null)
              cell.s = styleRemap[cell.s];
          }
        }
        const newId = `imp-${Date.now().toString(36)}-${seq++}`;
        sheet.id = newId;
        let nm = sheet.name || "Sheet";
        let n = 1;
        while (names.has(nm)) nm = `${sheet.name} (${n++})`;
        names.add(nm);
        sheet.name = nm;
        cur.sheets[newId] = sheet;
        cur.sheetOrder.push(newId);
        added.push(nm);
      }
      univerSnapRef.current = cur;
      univerReadyRef.current = true;
      setUniverSnap(cur);
      setUniverKey((k) => k + 1);
      saveDays(currentDaysForSave());
      lastEditRef.current = Date.now();
      serverSave();
      alert(
        added.length
          ? `Added ${added.length} sheet${added.length > 1 ? "s" : ""}: ${added.join(
              ", "
            )}\n\nLook for the new tab${
              added.length > 1 ? "s" : ""
            } at the far right of the tab bar.`
          : "That file had no sheets to import."
      );
    } catch (e) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const msg = (e as any)?.message || String(e);
      alert(
        "Could not import that file: " +
          msg +
          "\n\n(Only .xlsx files are supported — .xls/.csv won't work for import.)"
      );
    }
  }

  // Download the current spreadsheet as a real .xlsx (values + styling).
  async function downloadExcel() {
    const snap = univerSnapRef.current;
    if (!snap || !snap.sheets) {
      alert("Nothing to download yet — the sheet is empty.");
      return;
    }
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ExcelJSmod: any = await import("exceljs");
      const ExcelJS = ExcelJSmod.default ?? ExcelJSmod;
      const { univerSnapshotToExcel } = await import("@/lib/univerToExcel");
      const wb = await univerSnapshotToExcel(snap, ExcelJS);
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${(movieName || "sheet").replace(/[^\w.-]+/g, "_")}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    } catch {
      alert("Could not generate the Excel file.");
    }
  }

  async function exportPDF() {
    const jsPDFmod = await import("jspdf");
    const autoTable = (await import("jspdf-autotable")).default;
    const doc = new jsPDFmod.default({ orientation: "landscape" });
    currentDays().forEach((d, idx) => {
      if (idx > 0) doc.addPage();
      const ncols = d.data.reduce((m, r) => Math.max(m, r.length), ncolsFor(splCountRef.current));
      doc.setFontSize(12);
      doc.text(`${movieName} — ${d.name}`, 14, 12);
      autoTable(doc, {
        head: [titlesFor(splCountRef.current, ncols)],
        body: d.data.map((r) => r.map((c) => (c === "" ? "" : String(c)))),
        startY: 16,
        styles: { fontSize: 6, cellPadding: 1 },
        headStyles: { fillColor: [65, 126, 55] },
      });
    });
    doc.save(`${movieName}.pdf`);
  }

  return (
    <div className="flex h-screen flex-col overflow-hidden bg-canvas">
      {/* Save-status toast (Ctrl+S / auto-save) */}
      {saveState !== "idle" && (
        <div
          className="fixed left-5 top-5 z-[200] flex items-center gap-2 rounded-full px-4 py-2 text-xs font-semibold text-white shadow-lg"
          style={{
            backgroundColor:
              saveState === "saving" ? "#6b7280" : "var(--brand-600)",
            transition: "opacity 200ms ease",
          }}
        >
          {saveState === "saving" ? (
            <>
              <span className="inline-block h-3 w-3 animate-spin rounded-full border-2 border-white/40 border-t-white" />
              Saving…
            </>
          ) : (
            <>✓ Saved</>
          )}
        </div>
      )}
      <div
        ref={headerRef}
        className="movie-toolbar flex shrink-0 flex-wrap items-center justify-between gap-2 px-6 py-3"
      >
        <div className="flex items-center gap-3">
          <Link
            href="/movies"
            className="rounded-md border border-line bg-surface px-2.5 py-1.5 text-xs font-medium text-body hover:bg-chip"
          >
            ← Movies
          </Link>
          <div>
            <h1
              className="text-base font-bold"
              style={{ color: "var(--brand-600)" }}
            >
              {movieName}
            </h1>
            <p className="text-[11px] text-faint">
              Spreadsheet · shared &amp; auto-saved
            </p>
          </div>
          {me.email && (
            <span
              title={`Signed in as ${me.email}`}
              className="ml-1 rounded-full border border-brand-300 bg-surface px-2.5 py-1 text-[11px] font-semibold"
              style={{ color: "var(--brand-700)" }}
            >
              {me.isAdmin ? "👑 " : "👤 "}
              {me.email}
            </span>
          )}
        </div>
        <div className="flex flex-wrap items-center gap-2">
          {!plainMode && (
          <>
          <button
            onClick={addDay}
            className="rounded-md bg-brand-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-brand-700"
          >
            + Add day
          </button>
          <button
            onClick={deleteDay}
            title="Delete the current day tab"
            className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-rose-50 hover:text-rose-600"
          >
            ✕ Delete day
          </button>
          <span className="mx-1 h-4 w-px bg-line" />
          <button
            onClick={addColumn}
            className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-chip"
          >
            + Column
          </button>
          <button
            onClick={removeColumn}
            className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-chip"
          >
            − Column
          </button>
          <span className="mx-1 h-4 w-px bg-line" />
          <button
            onClick={addSpecialShow}
            title="Add another special show (Spl-N)"
            className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-chip"
          >
            + Special show
          </button>
          <button
            onClick={toggleTerms}
            title="Show/hide the Terms column"
            className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-chip"
          >
            {showTerms ? "Hide terms" : "Show terms"}
          </button>
          <button
            onClick={() => setFormulasOpen(true)}
            title="See how Gross, Nett, Share, GT etc. are calculated"
            className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-chip"
          >
            ƒ Formulas
          </button>
          <select
            value={filter}
            onChange={(e) =>
              applyFilter(e.target.value as "all" | "multiplex" | "single")
            }
            className="rounded-md border border-line bg-surface px-2 py-1.5 text-xs font-medium text-body outline-none focus:border-brand-400"
          >
            <option value="all">All theatres</option>
            <option value="multiplex">Multiplexes only</option>
            <option value="single">Single screens only</option>
          </select>
          <span className="mx-1 h-4 w-px bg-line" />
          <button
            onClick={() => pullRepData(false)}
            title="Fill this day's show cells from representative submissions"
            className="rounded-md bg-brand-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-brand-700"
          >
            ⟳ Pull rep data
          </button>
          <span className="mx-1 h-4 w-px bg-line" />
          </>
          )}
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            className="hidden"
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) importExcel(f);
              e.target.value = "";
            }}
          />
          <button
            onClick={() => fileInputRef.current?.click()}
            title="Upload an Excel/CSV file (replaces the current view)"
            className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-chip"
          >
            ⬆ Upload Excel
          </button>
          {plainMode && (
            <>
              <input
                ref={appendInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                className="hidden"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) appendExcel(f);
                  e.target.value = "";
                }}
              />
              <button
                onClick={() => appendInputRef.current?.click()}
                title="Import a local Excel and add its sheets to this workbook"
                className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-chip"
              >
                ➕ Import into workbook
              </button>
              <button
                onClick={downloadExcel}
                title="Download this spreadsheet as an .xlsx file"
                className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-chip"
              >
                ⬇ Download Excel
              </button>
              {hasSplColumns() && (
                <button
                  onClick={toggleSpl}
                  title="Show or hide the special-show (Spl) columns"
                  className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-chip"
                >
                  {univerSplHidden
                    ? "▸ Show special shows"
                    : "▾ Hide special shows"}
                </button>
              )}
            </>
          )}
          <button
            onClick={() => setShareOpen(true)}
            title="Share this movie with others"
            className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-chip"
          >
            🔗 Share
          </button>
          {readOnly && (
            <span className="rounded bg-amber-100 px-2 py-1 text-[11px] font-semibold text-amber-700">
              View only
            </span>
          )}
          <span className="mx-1 h-4 w-px bg-line" />
          {!plainMode && (
            <>
              <button
                onClick={exportExcel}
                className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-chip"
              >
                Export Excel
              </button>
              <button
                onClick={exportPDF}
                className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-chip"
              >
                Export PDF
              </button>
            </>
          )}
          <button
            onClick={resetSheet}
            className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-body hover:bg-chip"
          >
            Reset
          </button>
        </div>
      </div>

      {/* Day tabs live here — a fixed, non-scrolling slot just below the toolbar
          so they stay visible no matter how far the sheet is scrolled. The
          native jspreadsheet tab bar is moved into this slot after each render. */}
      <div
        ref={tabsSlotRef}
        className="shrink-0 border-b border-line bg-canvas px-2 empty:hidden"
      />

      {/* sheet + side panels live in a horizontal row so panels sit BESIDE the
          sheet (below the toolbar) and never cover the header or its scrollbar */}
      <div className="flex min-h-0 flex-1">
        {plainMode ? (
          <div className="min-h-0 flex-1 p-3">
            {/* The spreadsheet lives inside a rounded card instead of bleeding
                edge-to-edge across the whole window. */}
            <div className="relative h-full w-full overflow-hidden rounded-xl border border-line bg-surface shadow-sm">
              {loading && (
                <div className="pointer-events-none absolute inset-0 z-20 flex items-center justify-center">
                  <div
                    className="flex items-center gap-3 rounded-xl border border-line px-5 py-4 text-sm font-medium text-body shadow-pop"
                    style={{ backgroundColor: "var(--surface)" }}
                  >
                    <span className="inline-block h-5 w-5 animate-spin rounded-full border-2 border-brand-300 border-t-brand-600" />
                    Loading spreadsheet…
                  </div>
                </div>
              )}
              {/* Lightweight in-sheet find bar (Ctrl+F) — searches live as you
                  type and jumps to matches, like a browser's find. */}
              {findOpen && (
                <div className="absolute right-3 top-3 z-30 flex items-center gap-1 rounded-lg border border-line bg-surface px-2 py-1.5 shadow-pop">
                  <input
                    ref={findInputRef}
                    value={findQuery}
                    onChange={(e) => runFind(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter") {
                        e.preventDefault();
                        stepFind(e.shiftKey ? -1 : 1);
                      } else if (e.key === "Escape") {
                        e.preventDefault();
                        closeFindBar();
                      }
                    }}
                    placeholder="Find in sheet"
                    className="w-44 bg-transparent px-1 text-sm text-body outline-none placeholder:text-faint"
                  />
                  <span className="min-w-[2.75rem] px-1 text-center text-xs tabular-nums text-faint">
                    {findQuery
                      ? findMatches.length
                        ? `${findIdx + 1}/${findMatches.length}`
                        : "0/0"
                      : ""}
                  </span>
                  <button
                    onClick={() => stepFind(-1)}
                    disabled={!findMatches.length}
                    title="Previous (Shift+Enter)"
                    className="rounded px-1.5 py-0.5 text-sm text-body hover:bg-chip disabled:opacity-40"
                  >
                    ↑
                  </button>
                  <button
                    onClick={() => stepFind(1)}
                    disabled={!findMatches.length}
                    title="Next (Enter)"
                    className="rounded px-1.5 py-0.5 text-sm text-body hover:bg-chip disabled:opacity-40"
                  >
                    ↓
                  </button>
                  <button
                    onClick={closeFindBar}
                    title="Close (Esc)"
                    className="rounded px-1.5 py-0.5 text-sm text-faint hover:bg-chip"
                  >
                    ✕
                  </button>
                </div>
              )}
              <UniverSheet
                key={univerKey}
                snapshot={univerSnap}
                onChange={onUniverChange}
                onRequestFind={openFindBar}
                onReady={(api, phase) => {
                  univerApiRef.current = api;
                  // Reflect the current toggle state on the freshly built book.
                  setSplVisibility(univerSplHidden);
                  // On first mount, return the user to exactly where they left
                  // off last time (saved sheet + scroll + selection), or A1 on a
                  // first-ever open. This also clears the stray column highlight
                  // that hiding columns can leave behind. Remote pulls (phase
                  // "replace") preserve position in place, so skip them here.
                  if (phase === "mount") {
                    restoreAndTrackView(api);
                  }
                }}
              />
            </div>
          </div>
        ) : (
          <div className="min-h-0 flex-1 overflow-auto px-2 pb-2">
            <div ref={ref} />
          </div>
        )}

        {formulasOpen && (
          <FormulaPanel open={formulasOpen} onClose={() => setFormulasOpen(false)} />
        )}
        {inspect && (
          <CellInspector
            key={`${inspect.x}:${inspect.y}`}
            data={inspect}
            onApply={applyInspectFormula}
            onApplyDeduction={applyDeduction}
            onApplyNettFormula={applyNettFormula}
            onResetNett={resetNett}
            onClose={() => setInspect(null)}
          />
        )}
      </div>

      <div ref={cardRef} className="hfa-card" style={{ display: "none" }} />

      {shareOpen && (
        <ShareDialog movieId={movieId} onClose={() => setShareOpen(false)} />
      )}

      {entry && (
        <ShowEntryDialog
          entry={entry}
          rows={entryRows}
          setRows={setEntryRows}
          onSave={saveShowEntry}
          onClose={() => setEntry(null)}
        />
      )}
    </div>
  );
}

// ── Cell inspector: shows how a derived value is reached; edit its formula ──
function CellInspector({
  data,
  onApply,
  onApplyDeduction,
  onApplyNettFormula,
  onResetNett,
  onClose,
}: {
  data: {
    label: string;
    why: string;
    worked: string;
    value: string;
    formula: string;
    values?: string;
    ded?: { tid: number; value: number; day: string; overridden: boolean };
    input?: boolean;
  };
  onApply: (formula: string) => void;
  onApplyDeduction: (tid: number, value: number) => void;
  onApplyNettFormula: (formula: string) => void;
  onResetNett: () => void;
  onClose: () => void;
}) {
  const [draft, setDraft] = useState(data.formula);
  const [dedDraft, setDedDraft] = useState(String(data.ded?.value ?? ""));
  // keep the editors in sync when the panel rebuilds for the same cell
  useEffect(() => setDraft(data.formula), [data.formula]);
  useEffect(() => setDedDraft(String(data.ded?.value ?? "")), [data.ded?.value]);
  const isNett = !!data.ded;
  const editable = data.formula !== "";
  return (
    <>
      <aside
        className="flex h-full w-[360px] shrink-0 flex-col overflow-hidden border-l border-line"
        style={{ backgroundColor: "var(--surface)" }}
      >
        <div className="flex items-start justify-between gap-3 border-b border-line px-5 py-4">
          <div>
            <h3 className="text-sm font-semibold text-strong">{data.label}</h3>
            <p className="text-xs text-faint">How this value is calculated</p>
          </div>
          <button
            onClick={onClose}
            className="rounded-md p-1 text-faint hover:bg-chip hover:text-body"
          >
            <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" stroke="currentColor" strokeWidth={2}>
              <path d="M6 6l12 12M18 6L6 18" />
            </svg>
          </button>
        </div>

        <div className="flex-1 overflow-y-auto px-5 py-4">
          <div className="mb-4">
            <div className="text-[11px] uppercase tracking-wide text-faint">Value</div>
            <div className="text-2xl font-semibold text-strong">{data.value || "—"}</div>
          </div>

          <div className="mb-4 rounded-lg border border-line bg-muted/40 px-3 py-2">
            <div className="text-[11px] uppercase tracking-wide text-faint">Why</div>
            <p className="mt-1 text-sm text-body">{data.why}</p>
            {data.worked ? (
              <p className="mt-2 rounded bg-chip px-2 py-1 font-mono text-[12px] text-body">
                {data.worked}
              </p>
            ) : null}
          </div>

          {data.ded ? (
            <div className="mb-4">
              <div className="mb-1 text-[11px] uppercase tracking-wide text-faint">
                Fixed deduction (₹)
              </div>
              <input
                type="number"
                value={dedDraft}
                onChange={(e) => setDedDraft(e.target.value)}
                className="w-40 rounded-md border border-line bg-surface px-3 py-2 text-sm outline-none focus:border-brand-400"
              />
              <p className="mt-1 text-[11px] text-faint">
                Per-theatre charge subtracted after GST. GST is fixed: 18% for
                tickets &gt; ₹105, 5% for ≤ ₹105.
              </p>
              <button
                onClick={() =>
                  onApplyDeduction(data.ded!.tid, Number(dedDraft) || 0)
                }
                className="mt-3 rounded-md bg-brand-600 px-4 py-2 text-sm font-semibold text-white hover:bg-brand-700"
              >
                Apply
              </button>
            </div>
          ) : null}

          <div>
            <div className="mb-1 flex items-center justify-between text-[11px] uppercase tracking-wide text-faint">
              <span>Formula</span>
              {isNett && data.ded?.overridden ? (
                <span className="rounded bg-amber-100 px-1.5 py-0.5 text-[10px] font-semibold normal-case text-amber-700">
                  overridden
                </span>
              ) : null}
            </div>
            {editable ? (
              <>
                <textarea
                  value={draft}
                  onChange={(e) => setDraft(e.target.value)}
                  spellCheck={false}
                  rows={4}
                  className="w-full resize-y whitespace-pre-wrap break-all rounded-md border border-line bg-surface px-3 py-2 font-mono text-[13px] leading-snug outline-none focus:border-brand-400"
                />
                <p className="mt-1 text-[11px] text-faint">
                  {isNett
                    ? "Edit freely — e.g. rent-based single screens or exceptions. Your formula overrides the auto GST calc for this cell."
                    : "Excel-style, e.g. =X2/1.18. Applies to this cell only."}
                </p>
                <div className="mt-3 flex gap-2">
                  <button
                    onClick={() => (isNett ? onApplyNettFormula(draft) : onApply(draft))}
                    className="rounded-md bg-brand-600 px-4 py-2 text-sm font-semibold text-white hover:bg-brand-700"
                  >
                    Apply
                  </button>
                  {isNett ? (
                    <button
                      onClick={onResetNett}
                      className="rounded-md border border-line px-4 py-2 text-sm font-medium text-body hover:bg-chip"
                    >
                      Reset to auto
                    </button>
                  ) : (
                    <button
                      onClick={() => setDraft(data.formula)}
                      className="rounded-md px-4 py-2 text-sm font-medium text-faint hover:bg-chip"
                    >
                      Reset
                    </button>
                  )}
                </div>
              </>
            ) : (
              <p className="rounded bg-chip px-2 py-1 font-mono text-[12px] text-body">
                {data.input
                  ? "Direct input — type the value in the cell."
                  : "Computed automatically — not directly editable."}
              </p>
            )}
          </div>

          {data.values ? (
            <div className="mt-4">
              <div className="mb-1 text-[11px] uppercase tracking-wide text-faint">
                Actual values
              </div>
              <p className="break-all rounded bg-chip px-2 py-1 font-mono text-[12px] text-body">
                {data.values}
                {data.value ? ` = ${data.value}` : ""}
              </p>
              <p className="mt-1 text-[11px] text-faint">
                The same formula with each cell reference replaced by its current
                number.
              </p>
            </div>
          ) : null}
        </div>
      </aside>
    </>
  );
}

// ── "How it's calculated" reference panel (documentation only) ──
const FORMULA_GROUPS: {
  group: string;
  items: { label: string; formula: string; why: string }[];
}[] = [
  {
    group: "Per screen",
    items: [
      {
        label: "Show (₹)",
        formula: "RATE × Audience",
        why: "Collection for a slot = ticket price × seats sold. Auto-fills when audience is entered.",
      },
      {
        label: "H.F.C (House-full collection)",
        formula: "Σ (class RATE × class AUD)",
        why: "Maximum possible gross if every seat sells, summed over all ticket classes.",
      },
      {
        label: "H.F.A (House-full audience)",
        formula: "Σ class AUD",
        why: "Total seating capacity of the screen (audience can't exceed this).",
      },
    ],
  },
  {
    group: "Per theatre — Today",
    items: [
      {
        label: "Shows",
        formula: "0.5 per filled Show + 0.5 per filled Aud",
        why: "How many show-halves have data — a full show (both Show & Aud) counts as 1.",
      },
      {
        label: "Today Gross",
        formula: "Σ all Show cells of the theatre",
        why: "The day's total collection across every screen and slot.",
      },
      {
        label: "Today Nett",
        formula: "Gross ÷ 1.18  (= Gross − Gross×18/118)",
        why: "Removes 18% GST from the gross to get the net collection.",
      },
      {
        label: "Today Share",
        formula: "Nett × top term %",
        why: "Distributor's cut — Nett times the current-week percentage from the Terms column (e.g. 52.5%).",
      },
      {
        label: "Today Audience",
        formula: "Σ all Aud cells of the theatre",
        why: "Total footfall for the day across all screens and slots.",
      },
    ],
  },
  {
    group: "Grand Total (across days)",
    items: [
      {
        label: "GT Gross / Nett / Share / Audience",
        formula: "Day 1 = Today;  Day N = previous day's GT + this day's Today",
        why: "Running cumulative totals carried forward across the day tabs.",
      },
    ],
  },
  {
    group: "Section total (red row)",
    items: [
      {
        label: "TOTAL row",
        formula: "SUM of each column over the block",
        why: "Grand totals for the whole multiplex block and the whole single-screen block.",
      },
    ],
  },
];

function FormulaPanel({ onClose }: { open: boolean; onClose: () => void }) {
  return (
    <>
      <aside
        className="flex h-full w-[380px] shrink-0 flex-col overflow-hidden border-l border-line"
        style={{ backgroundColor: "var(--surface)" }}
      >
        <div className="flex items-start justify-between gap-3 border-b border-line px-5 py-4">
          <div>
            <h3 className="text-base font-semibold text-strong">How it's calculated</h3>
            <p className="text-xs text-faint">
              Every derived number in the sheet and where it comes from.
            </p>
          </div>
          <button
            onClick={onClose}
            className="rounded-md p-1 text-faint hover:bg-chip hover:text-body"
          >
            <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" stroke="currentColor" strokeWidth={2}>
              <path d="M6 6l12 12M18 6L6 18" />
            </svg>
          </button>
        </div>

        <div className="flex-1 overflow-y-auto px-5 py-4">
          {FORMULA_GROUPS.map((g) => (
            <div key={g.group} className="mb-5">
              <h4 className="mb-2 text-[11px] font-semibold uppercase tracking-wide text-brand-700">
                {g.group}
              </h4>
              <div className="space-y-2">
                {g.items.map((it) => (
                  <div
                    key={it.label}
                    className="rounded-lg border border-line bg-muted/40 px-3 py-2"
                  >
                    <div className="text-sm font-semibold text-strong">{it.label}</div>
                    <div className="mt-1 rounded bg-chip px-2 py-1 font-mono text-[12px] text-body">
                      {it.formula}
                    </div>
                    <p className="mt-1.5 text-xs text-faint">{it.why}</p>
                  </div>
                ))}
              </div>
            </div>
          ))}
          <p className="text-[11px] text-faint">
            Nett strips 18% GST; Share uses the first percentage in the Terms
            column. GT accumulates across the day tabs.
          </p>
        </div>
      </aside>
    </>
  );
}

// ── Per-show entry dialog: add class prices + audience → fills Show & Aud ──
function ShowEntryDialog({
  entry,
  rows,
  setRows,
  onSave,
  onClose,
}: {
  entry: { name: string; show: string; cap: number };
  rows: { price: string; aud: string }[];
  setRows: (r: { price: string; aud: string }[]) => void;
  onSave: () => void;
  onClose: () => void;
}) {
  const num = (v: string) => (v === "" ? 0 : Number(v) || 0);
  const totalAud = rows.reduce((n, r) => n + num(r.aud), 0);
  const totalShow = rows.reduce((n, r) => n + num(r.price) * num(r.aud), 0);
  const over = entry.cap > 0 && totalAud > entry.cap;

  const setCell = (i: number, field: "price" | "aud", val: string) =>
    setRows(rows.map((r, j) => (j === i ? { ...r, [field]: val } : r)));
  const addRow = () => setRows([...rows, { price: "", aud: "" }]);
  const removeRow = (i: number) =>
    setRows(rows.length > 1 ? rows.filter((_, j) => j !== i) : rows);

  return (
    <div
      className="fixed inset-0 z-[70] flex items-center justify-center bg-black/60 p-4"
      onClick={onClose}
    >
      <div
        className="w-full max-w-md overflow-hidden rounded-xl border border-line shadow-pop"
        style={{ backgroundColor: "var(--surface)" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="border-b border-line px-5 py-4">
          <h3 className="text-base font-semibold text-strong">Add show entry</h3>
          <p className="text-xs text-faint">
            <b className="text-body">{entry.name || "Screen"}</b> · {entry.show}
            {entry.cap > 0 ? ` · capacity ${inr(entry.cap)}` : ""}
          </p>
        </div>

        <div className="px-5 py-4">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-xs text-faint">
                <th className="pb-1 font-medium">Class price (₹)</th>
                <th className="pb-1 font-medium">Audience</th>
                <th className="pb-1 text-right font-medium">Amount (₹)</th>
                <th className="w-6" />
              </tr>
            </thead>
            <tbody>
              {rows.map((r, i) => (
                <tr key={i}>
                  <td className="py-1 pr-2">
                    <input
                      type="number"
                      value={r.price}
                      onChange={(e) => setCell(i, "price", e.target.value)}
                      className="w-28 rounded border border-line bg-surface px-2 py-1 text-sm outline-none focus:border-brand-400"
                    />
                  </td>
                  <td className="py-1 pr-2">
                    <input
                      type="number"
                      value={r.aud}
                      autoFocus={i === 0}
                      onChange={(e) => setCell(i, "aud", e.target.value)}
                      className="w-24 rounded border border-line bg-surface px-2 py-1 text-sm outline-none focus:border-brand-400"
                    />
                  </td>
                  <td className="py-1 text-right text-body">
                    ₹{inr(num(r.price) * num(r.aud))}
                  </td>
                  <td className="py-1 text-right">
                    <button
                      onClick={() => removeRow(i)}
                      title="Remove class"
                      className="rounded p-1 text-faint hover:bg-chip hover:text-rose-500"
                    >
                      ✕
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>

          <button
            onClick={addRow}
            className="mt-2 rounded border border-dashed border-line px-2 py-1 text-[11px] font-medium text-body hover:border-brand-400 hover:text-brand-700"
          >
            + Add class
          </button>

          <div className="mt-4 flex items-center justify-between rounded-lg bg-muted px-3 py-2 text-sm">
            <span className="text-faint">
              Total audience{" "}
              <b className={over ? "text-rose-600" : "text-strong"}>
                {inr(totalAud)}
              </b>
              {over ? ` (over ${inr(entry.cap)})` : ""}
            </span>
            <span className="text-faint">
              Show <b className="text-strong">₹{inr(totalShow)}</b>
            </span>
          </div>
        </div>

        <div className="flex items-center justify-end gap-2 border-t border-line px-5 py-4">
          <button
            onClick={onClose}
            className="rounded-md px-4 py-2 text-sm font-medium text-faint hover:bg-chip"
          >
            Cancel
          </button>
          <button
            onClick={onSave}
            disabled={over}
            className="rounded-md bg-brand-600 px-4 py-2 text-sm font-semibold text-white hover:bg-brand-700 disabled:opacity-40"
          >
            Save
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Share dialog: invite people by email as Viewer / Editor ──
function ShareDialog({
  movieId,
  onClose,
}: {
  movieId: string;
  onClose: () => void;
}) {
  const DOMAIN = "@svf.in";
  const [owner, setOwner] = useState<string | null>(null);
  const [shares, setShares] = useState<{ email: string; role: string }[]>([]);
  const [username, setUsername] = useState(""); // part before @svf.in
  const [role, setRole] = useState<"editor" | "viewer">("editor");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [known, setKnown] = useState<{ email: string; name: string }[]>([]);

  async function load() {
    try {
      const res = await fetch(`/api/movies/${movieId}/shares`, { cache: "no-store" });
      if (!res.ok) {
        setErr(res.status === 403 ? "Only the owner can manage sharing." : "Failed to load.");
        return;
      }
      const j = await res.json();
      setOwner(j.owner ?? null);
      setShares(j.shares ?? []);
      setErr("");
    } catch {
      setErr("Failed to load.");
    }
  }
  useEffect(() => {
    load();
    // existing users → autocomplete + "recognized" check
    fetch("/api/users/suggest", { cache: "no-store" })
      .then((r) => (r.ok ? r.json() : { users: [] }))
      .then((j) => setKnown(j.users ?? []))
      .catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const emailFor = (u: string) => {
    const s = u.trim().toLowerCase();
    return s.includes("@") ? s : s.replace(/\s+/g, "") + DOMAIN;
  };
  const target = emailFor(username);
  const recognized = known.some((k) => k.email.toLowerCase() === target);

  async function add() {
    const e = target;
    if (!e.includes("@") || e.startsWith("@")) {
      setErr("Enter a username.");
      return;
    }
    setBusy(true);
    try {
      const res = await fetch(`/api/movies/${movieId}/shares`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email: e, role }),
      });
      if (!res.ok) {
        setErr((await res.json().catch(() => ({}))).error || "Failed.");
      } else {
        setUsername("");
        await load();
      }
    } finally {
      setBusy(false);
    }
  }
  async function remove(e: string) {
    await fetch(`/api/movies/${movieId}/shares?email=${encodeURIComponent(e)}`, {
      method: "DELETE",
    });
    load();
  }
  // Change a person's access (Editor ↔ Viewer). Re-POSTing upserts the role.
  async function changeRole(e: string, newRole: "editor" | "viewer") {
    setShares((prev) =>
      prev.map((s) => (s.email === e ? { ...s, role: newRole } : s))
    ); // optimistic
    await fetch(`/api/movies/${movieId}/shares`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email: e, role: newRole }),
    });
    load();
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4"
      onClick={onClose}
    >
      <div
        className="w-full max-w-md overflow-hidden rounded-xl border border-line shadow-pop"
        style={{ backgroundColor: "var(--surface)" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between border-b border-line px-5 py-4">
          <h3 className="text-base font-semibold text-strong">Share this movie</h3>
          <button
            onClick={onClose}
            className="rounded-md p-1 text-faint hover:bg-chip hover:text-body"
          >
            <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" stroke="currentColor" strokeWidth={2}>
              <path d="M6 6l12 12M18 6L6 18" />
            </svg>
          </button>
        </div>
        <div className="px-5 py-4">
          <div className="mb-2 flex gap-2">
            <div className="flex flex-1 items-stretch rounded-md border border-line bg-surface focus-within:border-brand-400">
              <input
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                placeholder="username"
                list="svf-user-suggestions"
                className="w-full rounded-l-md bg-transparent px-3 py-2 text-sm outline-none"
                onKeyDown={(e) => e.key === "Enter" && add()}
              />
              <span className="flex items-center rounded-r-md border-l border-line bg-muted px-2 text-sm text-faint">
                {DOMAIN}
              </span>
            </div>
            <datalist id="svf-user-suggestions">
              {known.map((k) => (
                <option key={k.email} value={k.email.replace(DOMAIN, "")}>
                  {k.name || k.email}
                </option>
              ))}
            </datalist>
            <select
              value={role}
              onChange={(e) => setRole(e.target.value as "editor" | "viewer")}
              className="rounded-md border border-line bg-surface px-2 py-2 text-sm text-body outline-none focus:border-brand-400"
            >
              <option value="editor">Editor</option>
              <option value="viewer">Viewer</option>
            </select>
            <button
              onClick={add}
              disabled={busy}
              className="rounded-md bg-brand-600 px-4 py-2 text-sm font-semibold text-white hover:bg-brand-700 disabled:opacity-40"
            >
              Add
            </button>
          </div>
          {/* recognize whether this username already has a login account */}
          {username.trim() && (
            <p className="mb-2 text-[11px]">
              {recognized ? (
                <span className="text-brand-600">✓ {target} is a registered user.</span>
              ) : (
                <span className="text-amber-600">
                  {target} has no login yet — create it in Admin → Users so they can sign in.
                </span>
              )}
            </p>
          )}
          {err && <p className="mb-2 text-xs text-rose-600">{err}</p>}

          <div className="rounded-lg border border-line">
            {owner && (
              <div className="flex items-center justify-between px-3 py-2 text-sm">
                <span className="text-body">{owner}</span>
                <span className="text-xs text-faint">Owner</span>
              </div>
            )}
            {shares.length === 0 && !owner && (
              <p className="px-3 py-3 text-sm text-faint">Not shared with anyone yet.</p>
            )}
            {shares.map((s) => (
              <div
                key={s.email}
                className="flex items-center justify-between border-t border-line px-3 py-2 text-sm"
              >
                <span className="text-body">{s.email}</span>
                <span className="flex items-center gap-3">
                  <select
                    value={s.role === "viewer" ? "viewer" : "editor"}
                    onChange={(e) =>
                      changeRole(s.email, e.target.value as "editor" | "viewer")
                    }
                    className="rounded-md border border-line bg-surface px-1.5 py-1 text-xs text-body outline-none focus:border-brand-400"
                  >
                    <option value="editor">Editor</option>
                    <option value="viewer">Viewer</option>
                  </select>
                  <button
                    onClick={() => remove(s.email)}
                    className="text-xs font-medium text-rose-600 hover:underline"
                  >
                    Remove
                  </button>
                </span>
              </div>
            ))}
          </div>
          <p className="mt-3 text-[11px] text-faint">
            Editors can change the sheet; Viewers can only look. Everyone with
            access sees each other&apos;s edits within a couple of seconds.
          </p>
        </div>
      </div>
    </div>
  );
}
