// Convert a parsed ExcelJS workbook into a Univer workbook snapshot (IWorkbookData
// shape), preserving values, formulas, and styling (fills, fonts, borders,
// alignment, number formats), merged cells, column widths, row heights, and
// sheet tab colors — so an uploaded .xlsx renders as a faithful replica.
/* eslint-disable @typescript-eslint/no-explicit-any */

// ARGB (e.g. "FF00FF00") or "00FF00" → "#RRGGBB"
function argbToHex(argb?: string): string | undefined {
  if (!argb) return undefined;
  let s = String(argb).replace(/^#/, "");
  if (s.length === 8) s = s.slice(2); // drop alpha
  if (s.length !== 6) return undefined;
  return "#" + s.toUpperCase();
}

// Theme palette (theme index 0-9). Default is the classic Office theme; the
// REAL palette is read from each uploaded file's xl/theme/theme1.xml and set
// via setThemePalette() so colors match the source exactly.
let THEME = [
  "FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD",
  "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646",
];

/** Parse xl/theme/theme1.xml's <a:clrScheme> into the 10-entry theme palette. */
export function parseThemePalette(xml: string): string[] | null {
  const m = xml.match(/<a:clrScheme[\s\S]*?<\/a:clrScheme>/);
  if (!m) return null;
  const block = m[0];
  // Each child (dk1, lt1, dk2, lt2, accent1..6, ...) in document order.
  const colors: string[] = [];
  const re = /<a:(dk1|lt1|dk2|lt2|accent[1-6])>\s*<a:(srgbClr|sysClr)[^>]*?(?:val|lastClr)="([0-9A-Fa-f]{6})"/g;
  let mm: RegExpExecArray | null;
  const byName: Record<string, string> = {};
  while ((mm = re.exec(block))) byName[mm[1]] = mm[3].toUpperCase();
  // Excel theme index → scheme slot: 0=lt1,1=dk1,2=lt2,3=dk2,4..9=accent1..6
  const order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
  for (const k of order) colors.push(byName[k] || "000000");
  return colors.length === 10 ? colors : null;
}

export function setThemePalette(palette?: string[] | null) {
  if (palette && palette.length === 10) THEME = palette;
}
// Legacy indexed color palette (common subset; 0-63).
const INDEXED: Record<number, string> = {
  0: "000000", 1: "FFFFFF", 2: "FF0000", 3: "00FF00", 4: "0000FF",
  5: "FFFF00", 6: "FF00FF", 7: "00FFFF", 8: "000000", 9: "FFFFFF",
  10: "FF0000", 11: "00FF00", 12: "0000FF", 13: "FFFF00", 14: "FF00FF",
  15: "00FFFF", 16: "800000", 17: "008000", 18: "000080", 19: "808000",
  20: "800080", 21: "008080", 22: "C0C0C0", 23: "808080", 40: "00CCFF",
  41: "CCFFFF", 42: "CCFFCC", 43: "FFFF99", 44: "99CCFF", 45: "FF99CC",
  46: "CC99FF", 47: "FFCC99", 50: "99CC00", 51: "FF9900", 55: "333300",
  64: "000000",
};

// Lighten/darken a hex by an Excel tint (-1..1).
function applyTint(hex: string, tint?: number): string {
  if (!tint) return hex;
  const n = parseInt(hex, 16);
  let r = (n >> 16) & 255,
    g = (n >> 8) & 255,
    b = n & 255;
  const adj = (c: number) =>
    tint < 0
      ? Math.round(c * (1 + tint))
      : Math.round(c + (255 - c) * tint);
  r = adj(r);
  g = adj(g);
  b = adj(b);
  return ((r << 16) | (g << 8) | b).toString(16).padStart(6, "0").toUpperCase();
}

// Resolve an ExcelJS color object (argb | theme+tint | indexed) → "#RRGGBB".
function resolveColor(color: any): string | undefined {
  if (!color) return undefined;
  if (color.argb) return argbToHex(color.argb);
  if (typeof color.theme === "number") {
    const base = THEME[color.theme] ?? "000000";
    return "#" + applyTint(base, color.tint);
  }
  if (typeof color.indexed === "number") {
    const base = INDEXED[color.indexed];
    if (base) return "#" + applyTint(base, color.tint);
  }
  return undefined;
}

const H_ALIGN: Record<string, number> = { left: 1, center: 2, right: 3 };
const V_ALIGN: Record<string, number> = { top: 1, middle: 2, bottom: 3 };
// Minimum row height (px). The reports use a 14pt font (~19px); a shorter row
// crams the value so lines overlap ("puffed"). 22px gives the value breathing room.
export const MIN_ROW_H = 22;
// ExcelJS border style → Univer border style enum (approx: thin=1, medium=2, thick=3, dashed=4, dotted=5, double=6)
const BORDER_STYLE: Record<string, number> = {
  thin: 1,
  hair: 1,
  medium: 2,
  thick: 3,
  dashed: 4,
  dashDot: 4,
  dashDotDot: 4,
  dotted: 5,
  double: 6,
};

function cellStyle(cell: any): any | undefined {
  const s: any = {};
  const font = cell.font;
  if (font) {
    if (font.bold) s.bl = 1;
    if (font.italic) s.it = 1;
    // Strikethrough on report cells is copy-paste cruft (these Excel files carry
    // a stray "bold+strike+underline" font). Never import the strikethrough, and
    // skip the underline that rides on the same cruft font — but KEEP a genuine
    // standalone underline (e.g. the report title, which has no strike).
    if (font.underline && !font.strike) s.ul = { s: 1 };
    if (font.size) s.fs = font.size;
    // Font family is forced to Calibri workbook-wide via defaultStyle (below),
    // so we deliberately don't copy per-cell font names here.
    const cl = resolveColor(font.color);
    if (cl) s.cl = { rgb: cl };
  }
  const fill = cell.fill;
  if (fill && fill.type === "pattern" && fill.pattern && fill.pattern !== "none") {
    // solid → fgColor is the fill; some patterns carry it in bgColor.
    const bg = resolveColor(fill.fgColor) || resolveColor(fill.bgColor);
    if (bg) s.bg = { rgb: bg };
  }
  const al = cell.alignment;
  if (al) {
    if (al.horizontal && H_ALIGN[al.horizontal]) s.ht = H_ALIGN[al.horizontal];
    if (al.vertical && V_ALIGN[al.vertical]) s.vt = V_ALIGN[al.vertical];
    if (al.wrapText) s.tb = 3; // wrap
  }
  const b = cell.border;
  if (b) {
    const side = (x: any) =>
      x && x.style
        ? { s: BORDER_STYLE[x.style] ?? 1, cl: { rgb: resolveColor(x.color) || "#000000" } }
        : undefined;
    const bd: any = {};
    if (side(b.top)) bd.t = side(b.top);
    if (side(b.bottom)) bd.b = side(b.bottom);
    if (side(b.left)) bd.l = side(b.left);
    if (side(b.right)) bd.r = side(b.right);
    if (Object.keys(bd).length) s.bd = bd;
  }
  if (cell.numFmt) s.n = { pattern: cell.numFmt };
  // Vertically align every value to the MIDDLE of the cell (centered). Univer
  // renders a cell that HAS a style id but no explicit vertical alignment at the
  // TOP of the row (values look like they "float"), and it does NOT fall back to
  // the sheet defaultStyle for those — so we bake the alignment into the style:
  // honor Excel's own vertical alignment when set, otherwise use MIDDLE.
  if (s.vt == null) s.vt = 2; // 2 = MIDDLE (centered)
  return s;
}

// A value that is only whitespace (e.g. a stray " " copied from Excel) is junk:
// it counts as text, so any +/‑/* formula that references it evaluates to
// #VALUE!. Treat such cells as empty so the arithmetic (and SUMs) still work.
function isBlankText(v: any): boolean {
  return typeof v === "string" && v.trim() === "";
}

function cellValue(cell: any): { v?: any; f?: string } {
  // Formula cell (regular OR shared) — cell.formula returns the translated
  // formula string for this cell. Keep the formula (so it shows in the formula
  // bar) AND the cached result (so the value still shows even if Univer can't
  // recompute a given function).
  if (cell.formula) {
    const raw = String(cell.formula);
    const f = raw.startsWith("=") ? raw : "=" + raw;
    const r = cell.result;
    // Numeric / real string result → keep it as the cached value. Skip blank
    // strings and error objects (#VALUE! etc.) — never bake an error into a cell.
    if (r != null && typeof r !== "object" && !isBlankText(r)) return { v: r, f };
    // Error result or no cached result → just the formula (Univer recomputes).
    return { f };
  }
  const val = cell.value;
  if (val == null) return {};
  if (typeof val !== "object")
    return isBlankText(val) ? {} : { v: val }; // number / string / boolean
  if (val instanceof Date) return { v: val.toLocaleDateString("en-GB") };
  // Object-typed values — extract the real text, never "[object Object]".
  if (val.richText) {
    const t = val.richText.map((x: any) => x.text).join("");
    return isBlankText(t) ? {} : { v: t };
  }
  if (val.text != null) return isBlankText(val.text) ? {} : { v: val.text };
  if (val.error) return {}; // don't import Excel's cached errors (#VALUE!, #REF!…)
  if (val.result != null && typeof val.result !== "object")
    return isBlankText(val.result) ? {} : { v: val.result };
  if (val.formula) return { f: "=" + val.formula }; // formula w/o cached result
  return {}; // unknown object → leave the cell empty (not "[object Object]")
}

export function excelToUniverSnapshot(wb: any, name = "Uploaded"): any {
  const sheets: any = {};
  const sheetOrder: string[] = [];
  // Deduplicate styles into a shared map (cells reference by id). This shrinks
  // the snapshot 10–100× vs. inlining the full style on every cell — essential
  // so it fits in storage and saves reliably.
  const styles: Record<string, any> = {};
  const styleIds: Record<string, string> = {};
  let styleSeq = 0;
  const styleId = (s: any): string | undefined => {
    if (!s) return undefined;
    const key = JSON.stringify(s);
    let id = styleIds[key];
    if (!id) {
      id = String(++styleSeq);
      styleIds[key] = id;
      styles[id] = s;
    }
    return id;
  };
  wb.worksheets.forEach((ws: any, i: number) => {
    const sheetId = `sheet-${i + 1}`;
    sheetOrder.push(sheetId);
    const cellData: any = {};
    let maxRow = 0;
    let maxCol = 0;
    // Columns whose header reads "Spl - N" (special shows) — collapsed by default.
    const splCols = new Set<number>();
    ws.eachRow({ includeEmpty: false }, (row: any, rowNumber: number) => {
      const r = rowNumber - 1;
      row.eachCell({ includeEmpty: false }, (cell: any, colNumber: number) => {
        const c = colNumber - 1;
        const { v, f } = cellValue(cell);
        const s = styleId(cellStyle(cell));
        if (typeof v === "string" && /^\s*spl\s*-/i.test(v)) splCols.add(c);
        if (v == null && f == null && !s) return;
        (cellData[r] ||= {})[c] = {
          ...(v != null ? { v } : {}),
          ...(f ? { f } : {}),
          ...(s ? { s } : {}),
        };
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;
      });
    });

    // merged ranges
    const mergeData: any[] = [];
    const merges = ws.model?.merges || [];
    for (const m of merges) {
      // m like "A1:B2"
      try {
        const range = typeof m === "string" ? m : m.toString();
        const [tl, br] = range.split(":");
        const dec = (a1: string) => {
          const mm = /^([A-Z]+)(\d+)$/.exec(a1);
          if (!mm) return null;
          let col = 0;
          for (const ch of mm[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
          return { r: Number(mm[2]) - 1, c: col - 1 };
        };
        const a = dec(tl);
        const b = dec(br || tl);
        if (a && b)
          mergeData.push({
            startRow: a.r,
            endRow: b.r,
            startColumn: a.c,
            endColumn: b.c,
          });
      } catch {}
    }

    // column widths / row heights
    const columnData: any = {};
    (ws.columns || []).forEach((col: any, idx: number) => {
      if (col && col.width) columnData[idx] = { w: Math.round(col.width * 7) };
      if (col && idx > maxCol) maxCol = idx;
    });
    // Preserve hidden columns. Read from the RAW parsed column ranges
    // (ws.model.cols) — the authoritative source that reliably reflects how
    // Excel stored hidden ranges — plus a fallback via ws.columns. A column with
    // width 0 is treated as hidden too (some files hide by zeroing the width).
    const hiddenCols = new Set<number>();
    const rawCols: any[] = ws.model?.cols || [];
    for (const rc of rawCols) {
      const min = (rc?.min ?? 1) - 1;
      const max = (rc?.max ?? rc?.min ?? 1) - 1;
      if (rc?.hidden === true || rc?.width === 0)
        for (let c = min; c <= max; c++) hiddenCols.add(c);
    }
    (ws.columns || []).forEach((col: any, idx: number) => {
      if (col && (col.hidden === true || col.width === 0)) hiddenCols.add(idx);
    });
    for (const c of hiddenCols)
      columnData[c] = { ...(columnData[c] || {}), hd: 1 };
    // Hide the special-show columns on load (hd = 1); a toolbar toggle re-shows.
    for (const c of splCols) columnData[c] = { ...(columnData[c] || {}), hd: 1 };
    const rowData: any = {};
    ws.eachRow({ includeEmpty: false }, (row: any, rowNumber: number) => {
      const rd: any = {};
      if (row.height) {
        // Excel row height is in POINTS; Univer wants PIXELS (× 96/72). Floor at
        // MIN_ROW_H so the 14pt report font isn't crammed into a too-short row
        // (values otherwise overlap top-to-bottom — the "puffed" look).
        rd.h = Math.max(Math.round(row.height * (96 / 72)), MIN_ROW_H);
      }
      if (row.hidden) rd.hd = 1; // preserve rows hidden in the source Excel
      if (Object.keys(rd).length) rowData[rowNumber - 1] = rd;
    });

    sheets[sheetId] = {
      id: sheetId,
      name: ws.name || `Sheet${i + 1}`,
      // Whole sheet renders in Calibri, middle-aligned unless a cell overrides
      // it (matches the per-cell default baked in by cellStyle()).
      defaultStyle: { ff: "Calibri", vt: 2 },
      tabColor: resolveColor(ws.properties?.tabColor),
      // Generous grid so blank space to the right/below fills like a real
      // spreadsheet (Excel/Google Sheets show a full grid of empty cells).
      rowCount: Math.max(maxRow + 100, 500),
      columnCount: Math.max(maxCol + 26, 52),
      defaultColumnWidth: 88,
      defaultRowHeight: 22,
      mergeData,
      cellData,
      rowData,
      columnData,
      showGridlines: 1,
    };
  });

  if (!sheetOrder.length) {
    // empty fallback
    sheets["sheet-1"] = {
      id: "sheet-1",
      name: "Sheet1",
      rowCount: 500,
      columnCount: 52,
      cellData: {},
    };
    sheetOrder.push("sheet-1");
  }

  return {
    id: "wb-" + Date.now().toString(36),
    name,
    appVersion: "1.0.0",
    locale: "enUS",
    defaultStyle: { ff: "Calibri", vt: 2 }, // Calibri + middle-aligned (centered)
    sheetOrder,
    sheets,
    styles,
    resources: [],
  };
}
