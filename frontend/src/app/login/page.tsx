"use client";

import { useEffect, useRef, useState } from "react";
import { useRouter } from "next/navigation";

const MONO = "'Roboto Mono', ui-monospace, SFMono-Regular, Menlo, monospace";
const SANS =
  "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif";

const NAV: { label?: string; s?: string; navId?: number }[] = [
  { s: "Menu" },
  { label: "Dashboard", navId: 1 },
  { label: "Theatres", navId: 0 },
  { label: "Representatives" },
  { label: "Users" },
  { label: "History" },
  { s: "Movies" },
  { label: "Toxic", navId: 2 },
];

/* Where each nav tab sits, in main-content coordinates (negative x reaches into
   the sidebar) — so the cursor can glide up and "click" a tab to switch views. */
const NAV_MAIN_POS: Record<number, [string, string]> = {
  1: ["-15%", "19%"], // Dashboard
  0: ["-15%", "26%"], // Theatres
  2: ["-15%", "58%"], // Toxic
};

type Detail = {
  code: string;
  title: string;
  tag: string;
  lines: { k: string; v: string; tint?: string }[];
  foot: { k: string; v: string };
};
type FocusStep = { hi: number; cursor: [string, string]; detail: Detail };

type TheatresScene = {
  kind: "theatres";
  nav: number;
  crumb: string;
  title: string;
  tag: string;
  cards: { k: string; v: string; sub: string; tint?: string }[];
  cols: string;
  heads: { label: string; align: string }[];
  rows: string[][];
  focusSteps: FocusStep[];
};
type DashboardScene = {
  kind: "dashboard";
  nav: number;
  title: string;
  sub: string;
  metrics: { k: string; v: string; delta: string; bars: number[] }[];
  formats: { k: string; pct: number; v: string; share: string; dot: string; color: string }[];
  centres: { name: string; gross: string; occ: string }[];
  focusSteps: FocusStep[];
};
type CollectionsScene = {
  kind: "collections";
  nav: number;
  poster: string;
  posterInit: string;
  movie: string;
  sub: string;
  kpis: { k: string; v: string; delta: string; up: boolean; bars: number[] }[];
  districts: { name: string; gross: string; pct: number; occ: string }[];
  theatres: { name: string; centre: string; v: string }[];
  focusSteps: FocusStep[];
};
type Scene = TheatresScene | DashboardScene | CollectionsScene;

const SCENES: Scene[] = [
  {
    kind: "theatres",
    nav: 0,
    crumb: "Hyderabad › Gachibowli",
    title: "AMB Cinemas",
    tag: "Multiplex",
    cards: [
      { k: "Screens", v: "7", sub: "all active" },
      { k: "Total seats", v: "1,884", sub: "91% avg occ" },
      { k: "Full house", v: "₹5,02,640", sub: "per show", tint: "#05502F" },
      { k: "Format", v: "4K Atmos", sub: "IMAX on SC-1" },
    ],
    cols: "1.5fr 0.9fr 46px 82px 1fr",
    heads: [
      { label: "Screen", align: "left" },
      { label: "Format", align: "left" },
      { label: "Seats", align: "center" },
      { label: "Full house ₹", align: "right" },
      { label: "Rate slabs", align: "left" },
    ],
    rows: [
      ["AMB SC-1", "IMAX", "402", "1,60,800", "₹400 × 402"],
      ["Screen-2", "4K Atmos", "318", "1,01,760", "₹320 × 318"],
      ["Screen-3", "4K Atmos", "286", "80,080", "₹280 × 286"],
      ["Screen-4", "—", "244", "61,000", "₹250 × 244"],
      ["Screen-5", "—", "312", "78,000", "₹250 × 312"],
    ],
    focusSteps: [
      {
        hi: 0, cursor: ["30%", "60%"],
        detail: { code: "SCREEN · AMB SC-1", title: "AMB SC-1", tag: "IMAX", lines: [{ k: "Seats", v: "402" }, { k: "Occupancy", v: "95%", tint: "#05502F" }, { k: "Rate slab", v: "₹400 × 402" }, { k: "Shows today", v: "6" }], foot: { k: "Full house", v: "₹1,60,800" } },
      },
      {
        hi: 1, cursor: ["30%", "66%"],
        detail: { code: "SCREEN · SCREEN-2", title: "Screen-2", tag: "4K Atmos", lines: [{ k: "Seats", v: "318" }, { k: "Occupancy", v: "89%", tint: "#05502F" }, { k: "Rate slab", v: "₹320 × 318" }, { k: "Shows today", v: "5" }], foot: { k: "Full house", v: "₹1,01,760" } },
      },
    ],
  },
  {
    kind: "dashboard",
    nav: 1,
    title: "Dashboard",
    sub: "Nizam territory · Week 6",
    metrics: [
      { k: "Theatres", v: "482", delta: "+9", bars: [40, 48, 44, 56, 52, 64, 62, 78] },
      { k: "Screens", v: "631", delta: "+19", bars: [30, 38, 46, 42, 58, 56, 72, 84] },
      { k: "Occupancy", v: "77%", delta: "▲ 8", bars: [46, 40, 58, 52, 66, 64, 78, 86] },
      { k: "Week gross", v: "₹5.14Cr", delta: "+15%", bars: [24, 32, 40, 50, 58, 66, 76, 90] },
    ],
    formats: [
      { k: "2K QUBE", pct: 100, v: "318", share: "50%", dot: "#EA580C", color: "linear-gradient(90deg,#F97316,#EA580C)" },
      { k: "4K Atmos", pct: 43, v: "138", share: "22%", dot: "#E8834A", color: "linear-gradient(90deg,#F6A968,#E8834A)" },
      { k: "Standard", pct: 35, v: "112", share: "18%", dot: "#E29A63", color: "linear-gradient(90deg,#F1B98A,#E29A63)" },
      { k: "IMAX", pct: 12, v: "38", share: "6%", dot: "#C2410C", color: "linear-gradient(90deg,#8A2A06,#C2410C)" },
      { k: "Dolby 3D", pct: 6, v: "20", share: "3%", dot: "#B4531E", color: "linear-gradient(90deg,#D98A55,#B4531E)" },
      { k: "4DX", pct: 2, v: "5", share: "1%", dot: "#9A3412", color: "linear-gradient(90deg,#C2764A,#9A3412)" },
    ],
    centres: [
      { name: "Hyderabad", gross: "₹1.71Cr", occ: "78%" },
      { name: "Karimnagar", gross: "₹38.4L", occ: "69%" },
      { name: "Warangal", gross: "₹35.6L", occ: "67%" },
      { name: "Khammam", gross: "₹32.9L", occ: "65%" },
    ],
    focusSteps: [
      {
        hi: 0, cursor: ["78%", "42%"],
        detail: { code: "CENTRE · HYDERABAD", title: "Hyderabad", tag: "Top centre", lines: [{ k: "Theatres", v: "152" }, { k: "Screens", v: "174" }, { k: "Occupancy", v: "78%", tint: "#05502F" }, { k: "Best format", v: "2K QUBE" }], foot: { k: "Week gross", v: "₹1.71Cr" } },
      },
      {
        hi: 1, cursor: ["78%", "54%"],
        detail: { code: "CENTRE · KARIMNAGAR", title: "Karimnagar", tag: "Rank 2", lines: [{ k: "Theatres", v: "54" }, { k: "Screens", v: "63" }, { k: "Occupancy", v: "69%", tint: "#05502F" }, { k: "Best format", v: "2K QUBE" }], foot: { k: "Week gross", v: "₹38.4L" } },
      },
    ],
  },
  {
    kind: "collections",
    nav: 2,
    poster: "linear-gradient(150deg,#3A0F05,#C2410C)",
    posterInit: "TX",
    movie: "TOXIC",
    sub: "Day 6 · Running",
    kpis: [
      { k: "Gross today", v: "₹28.6L", delta: "+9%", up: true, bars: [30, 44, 38, 58, 52, 70, 66, 78, 84] },
      { k: "Week gross", v: "₹3.04Cr", delta: "+11%", up: true, bars: [24, 32, 40, 48, 56, 66, 74, 82, 90] },
      { k: "Occupancy", v: "70%", delta: "▲ 4", up: true, bars: [40, 34, 52, 46, 62, 60, 72, 68, 78] },
    ],
    districts: [
      { name: "Hyderabad", gross: "₹15.88L", pct: 100, occ: "74%" },
      { name: "Nalgonda", gross: "₹4.12L", pct: 26, occ: "62%" },
      { name: "Karimnagar", gross: "₹3.94L", pct: 25, occ: "60%" },
      { name: "Warangal", gross: "₹3.57L", pct: 22, occ: "58%" },
      { name: "Khammam", gross: "₹3.02L", pct: 19, occ: "54%" },
      { name: "Mahabubnagar", gross: "₹2.81L", pct: 18, occ: "52%" },
    ],
    theatres: [
      { name: "Prasads Multiplex", centre: "Hyderabad", v: "₹8,04,200" },
      { name: "AMB Cinemas", centre: "Gachibowli", v: "₹6,42,700" },
      { name: "PVR Nexus", centre: "Banjara Hills", v: "₹5,28,300" },
    ],
    focusSteps: [
      {
        hi: 0, cursor: ["30%", "44%"],
        detail: { code: "DIST · HYDERABAD", title: "Hyderabad", tag: "Top grosser", lines: [{ k: "Reporting", v: "152 / 152" }, { k: "Occupancy", v: "74%", tint: "#05502F" }, { k: "Best centre", v: "Gachibowli" }, { k: "Share", v: "₹7,12,600" }], foot: { k: "Gross", v: "₹15,88,000" } },
      },
      {
        hi: 1, cursor: ["30%", "51%"],
        detail: { code: "DIST · NALGONDA", title: "Nalgonda", tag: "Rank 2", lines: [{ k: "Reporting", v: "39 / 40" }, { k: "Occupancy", v: "62%" }, { k: "Best centre", v: "Miryalaguda" }, { k: "Share", v: "₹1,18,300" }], foot: { k: "Gross", v: "₹4,12,000" } },
      },
    ],
  },
];

function Sparkline({ bars, delay }: { bars: number[]; delay: number }) {
  return (
    <div style={{ height: 20, display: "flex", alignItems: "flex-end", gap: 2.5 }}>
      {bars.map((h, b) => (
        <span key={b} style={{ flex: 1, height: `${(h * 0.2).toFixed(1)}px`, borderRadius: 2, background: b === bars.length - 1 ? "#E4581A" : "rgba(138,42,6,0.22)", transformOrigin: "bottom", animation: "svf-growY 0.6s cubic-bezier(.16,1,.3,1) both", animationDelay: `${(delay + b * 0.035).toFixed(2)}s` }} />
      ))}
    </div>
  );
}

/* Scales a fixed-size design box down to fit any container width — keeps the
   preview pixel-perfect and identical on every device. */
function ScaledStage({ dw, dh, children }: { dw: number; dh: number; children: React.ReactNode }) {
  const ref = useRef<HTMLDivElement>(null);
  const [scale, setScale] = useState(1);
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const measure = () => setScale(Math.min(1, el.clientWidth / dw));
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, [dw]);
  return (
    <div ref={ref} style={{ width: "100%", display: "flex", justifyContent: "center" }}>
      <div style={{ width: dw * scale, height: dh * scale }}>
        <div style={{ width: dw, height: dh, transform: `scale(${scale})`, transformOrigin: "top left" }}>
          {children}
        </div>
      </div>
    </div>
  );
}

/* Live preview of the actual SVF admin app — cycles through real screens,
   and inside each screen steps through several rows, popping a detail card. */
function DistributionApp() {
  const [step, setStep] = useState(0);
  const [fi, setFi] = useState(-1); // focus index: -1 = browsing
  const timers = useRef<ReturnType<typeof setTimeout>[]>([]);

  useEffect(() => {
    const BROWSE = 1400, HOLD = 1900, NAVMOVE = 1200;
    const run = (s: number) => {
      timers.current.forEach(clearTimeout);
      timers.current = [];
      setStep(s);
      setFi(-1);
      const n = SCENES[s].focusSteps.length;
      for (let i = 0; i < n; i++) {
        timers.current.push(setTimeout(() => setFi(i), BROWSE + i * HOLD));
      }
      // move the cursor up to the next tab and "click" it, then switch view
      timers.current.push(setTimeout(() => setFi(-2), BROWSE + n * HOLD));
      timers.current.push(setTimeout(() => run((s + 1) % SCENES.length), BROWSE + n * HOLD + NAVMOVE));
    };
    run(0);
    return () => timers.current.forEach(clearTimeout);
  }, []);

  const sc = SCENES[step];
  const focus = fi >= 0;
  const navving = fi === -2;
  const fstep = focus ? sc.focusSteps[fi] : null;
  const detail = fstep?.detail;
  const nextNav = SCENES[(step + 1) % SCENES.length].nav;
  const cursorPos: [string, string] | null = fstep
    ? fstep.cursor
    : navving
      ? NAV_MAIN_POS[nextNav]
      : null;
  const rowHi = (i: number): React.CSSProperties =>
    fstep && i === fstep.hi
      ? { background: "#FFEEDC", boxShadow: "inset 3px 0 0 #EA580C" }
      : {};

  return (
    <div style={{ position: "relative", width: 760 }}>
      <div style={{ position: "absolute", left: 30, right: 30, bottom: -20, height: 64, borderRadius: 44, background: "rgba(120,40,4,0.34)", filter: "blur(22px)" }} />

      <div
        style={{
          position: "relative", display: "flex", height: 476, borderRadius: 16, overflow: "hidden",
          background: "#FFFDFA", fontFamily: SANS,
          boxShadow: "0 46px 90px rgba(120,40,4,0.32), 0 0 0 1px rgba(255,255,255,0.5)",
          transformOrigin: "center 44%",
          transform: focus ? "scale(0.955)" : "scale(1)",
          transition: "transform 0.5s cubic-bezier(.16,1,.3,1)",
        }}
      >
        {/* ── sidebar ── */}
        <div style={{ width: 176, flex: "none", padding: "16px 0 18px", background: "linear-gradient(180deg,#3A0F05 0%,#6E2105 46%,#A8380F 100%)", display: "flex", flexDirection: "column", gap: 10 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "0 15px" }}>
            <span style={{ width: 26, height: 26, borderRadius: 7, background: "#fff", display: "grid", placeItems: "center", fontWeight: 700, fontSize: 10, color: "#EA580C" }}>SVF</span>
            <span style={{ fontWeight: 600, fontSize: 13.5, color: "#FFEFE2" }}>Distribution</span>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 3, padding: "0 10px" }}>
            {NAV.map((n, i) =>
              n.s ? (
                <div key={i} style={{ padding: "8px 12px 3px", fontSize: 9, letterSpacing: "0.16em", textTransform: "uppercase", color: "rgba(255,214,184,0.6)", fontWeight: 600 }}>{n.s}</div>
              ) : (
                (() => {
                  const on = n.navId === sc.nav;
                  return (
                    <div key={i} style={{ display: "flex", alignItems: "center", height: 30, padding: "0 13px", borderRadius: 999, fontSize: 12.5, whiteSpace: "nowrap", background: on ? "#FFFDFA" : "transparent", color: on ? "#C2410C" : "rgba(255,239,226,0.82)", fontWeight: on ? 600 : 400, boxShadow: on ? "0 6px 14px rgba(20,6,2,0.34)" : "none", transition: "background 0.3s, color 0.3s" }}>{n.label}</div>
                  );
                })()
              )
            )}
          </div>
        </div>

        {/* ── main ── */}
        <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column", background: "#FFFDFA", position: "relative" }}>
          {/* header */}
          <div style={{ height: 62, flex: "none", display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 18px", borderBottom: "1px solid #F0E8DC" }}>
            <div key={"h" + step} style={{ display: "flex", alignItems: "center", gap: 12, minWidth: 0, animation: "svf-viewIn 0.4s both" }}>
              {sc.kind === "collections" && (
                <span style={{ width: 38, height: 44, flex: "none", borderRadius: 6, background: sc.poster, display: "grid", placeItems: "center", fontWeight: 700, fontSize: 12, color: "rgba(255,246,238,0.92)", boxShadow: "0 3px 8px rgba(120,40,4,0.3)" }}>{sc.posterInit}</span>
              )}
              <div style={{ minWidth: 0 }}>
                {sc.kind === "theatres" && <div style={{ fontSize: 10, color: "#9A8E80", whiteSpace: "nowrap" }}>{sc.crumb}</div>}
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontSize: 16, fontWeight: 700, color: "#2A1B09", whiteSpace: "nowrap" }}>
                    {sc.kind === "theatres" ? sc.title : sc.kind === "dashboard" ? sc.title : sc.movie}
                  </span>
                  {sc.kind === "theatres" && <span style={{ padding: "2px 8px", borderRadius: 5, background: "#FFE0C7", fontSize: 9.5, fontWeight: 600, color: "#C2410C" }}>{sc.tag}</span>}
                </div>
                {sc.kind === "dashboard" && <div style={{ marginTop: 2, fontSize: 11, color: "#9A8E80" }}>{sc.sub}</div>}
                {sc.kind === "collections" && <div style={{ marginTop: 2, fontSize: 11, color: "#9A8E80" }}>{sc.sub}</div>}
              </div>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 10, flex: "none" }}>
              <span style={{ display: "flex", alignItems: "center", gap: 6, height: 24, padding: "0 10px", borderRadius: 999, background: "#EAF6EF", fontSize: 10.5, fontWeight: 600, color: "#0F6F45" }}>
                <span style={{ width: 6, height: 6, borderRadius: "50%", background: "#17915B", boxShadow: "0 0 0 3px rgba(23,145,91,0.18)" }} />Live
              </span>
              <span style={{ width: 26, height: 26, borderRadius: "50%", background: "#EA580C", display: "grid", placeItems: "center", fontSize: 9.5, fontWeight: 600, color: "#fff" }}>AD</span>
            </div>
          </div>

          <div key={"v" + step} style={{ flex: 1, minHeight: 0, padding: "14px 18px 16px", display: "flex", flexDirection: "column", gap: 12, animation: "svf-viewIn 0.45s cubic-bezier(.16,1,.3,1) both" }}>

            {/* ═══════ THEATRES ═══════ */}
            {sc.kind === "theatres" && (
              <>
                <div style={{ flex: "none", display: "flex", gap: 10 }}>
                  {sc.cards.map((c, n) => (
                    <div key={n} style={{ flex: 1, minWidth: 0, padding: "10px 13px", borderRadius: 10, border: "1px solid #EFE6D9", background: "#FFFCF6", animation: "svf-rowIn 0.45s both", animationDelay: `${(0.06 + n * 0.07).toFixed(2)}s` }}>
                      <div style={{ fontSize: 8, letterSpacing: "0.12em", textTransform: "uppercase", color: "#8A6242" }}>{c.k}</div>
                      <div style={{ marginTop: 4, fontFamily: MONO, fontSize: c.v.length > 6 ? 15 : 20, fontWeight: 600, color: c.tint ?? "#2A1B09" }}>{c.v}</div>
                      <div style={{ marginTop: 3, fontSize: 9, color: "#A0917E" }}>{c.sub}</div>
                    </div>
                  ))}
                </div>
                <div style={{ flex: 1, minHeight: 0, display: "flex", flexDirection: "column", border: "1px solid #EFE6D9", borderRadius: 12, overflow: "hidden", background: "#FFFCF6" }}>
                  <div style={{ flex: "none", display: "flex", alignItems: "center", justifyContent: "space-between", padding: "11px 14px 9px" }}>
                    <span style={{ fontSize: 13, fontWeight: 600, color: "#2A1B09" }}>Screens &amp; rates</span>
                    <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                      <span style={{ padding: "2px 8px", borderRadius: 5, background: "#FDECEC", fontSize: 9, fontWeight: 600, color: "#C0392B" }}>No rep assigned</span>
                      <span style={{ padding: "2px 8px", borderRadius: 5, background: "#EA580C", fontSize: 9, fontWeight: 600, color: "#fff" }}>Assign</span>
                    </span>
                  </div>
                  <div style={{ flex: "none", display: "grid", gridTemplateColumns: sc.cols, background: "linear-gradient(180deg,#8A2A06,#6E2105)" }}>
                    {sc.heads.map((h, n) => (
                      <div key={n} style={{ padding: "6px 10px", borderRight: n < sc.heads.length - 1 ? "1px solid rgba(255,239,226,0.16)" : "none", fontSize: 7.5, letterSpacing: "0.08em", textTransform: "uppercase", color: "#FFE6D4", textAlign: h.align as React.CSSProperties["textAlign"] }}>{h.label}</div>
                    ))}
                  </div>
                  <div style={{ flex: 1, minHeight: 0 }}>
                    {sc.rows.map((cells, n) => (
                      <div key={n} style={{ display: "grid", gridTemplateColumns: sc.cols, borderBottom: "1px solid #F5EEE3", background: n % 2 ? "#FFFAF1" : "#FFFCF6", transition: "background 0.3s, box-shadow 0.3s", animation: "svf-rowIn 0.4s both", animationDelay: `${(0.14 + n * 0.06).toFixed(2)}s`, ...rowHi(n) }}>
                        {cells.map((v, c) => (
                          <div key={c} style={{ display: "flex", alignItems: "center", justifyContent: c === 2 ? "center" : c === 3 ? "flex-end" : "flex-start", gap: 5, padding: "8px 10px", borderRight: c < cells.length - 1 ? "1px solid #F5EEE3" : "none", fontFamily: c === 2 || c === 3 ? MONO : SANS, fontSize: 10.5, color: c === 0 ? "#2A1B09" : c === 3 ? "#0F6F45" : "#5B4C3A", fontWeight: c === 0 || c === 3 ? 600 : 400, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                            {c === 0 && n === 0 && <span style={{ width: 6, height: 6, flex: "none", borderRadius: "50%", background: "#17915B" }} />}
                            {v}
                          </div>
                        ))}
                      </div>
                    ))}
                  </div>
                </div>
              </>
            )}

            {/* ═══════ DASHBOARD ═══════ */}
            {sc.kind === "dashboard" && (
              <>
                <div style={{ flex: "none", display: "flex", gap: 10 }}>
                  {sc.metrics.map((m, n) => (
                    <div key={n} style={{ flex: 1, minWidth: 0, padding: "10px 13px", borderRadius: 10, border: "1px solid #EFE6D9", background: "#FFFCF6", animation: "svf-rowIn 0.45s both", animationDelay: `${(0.06 + n * 0.07).toFixed(2)}s` }}>
                      <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 6 }}>
                        <span style={{ fontSize: 8, letterSpacing: "0.12em", textTransform: "uppercase", color: "#8A6242" }}>{m.k}</span>
                        <span style={{ fontFamily: MONO, fontSize: 9, color: "#05502F" }}>{m.delta}</span>
                      </div>
                      <div style={{ marginTop: 4, fontFamily: MONO, fontSize: 19, fontWeight: 600, color: "#2A1B09", animation: "svf-countPulse 0.6s both", animationDelay: `${(0.12 + n * 0.07).toFixed(2)}s` }}>{m.v}</div>
                      <div style={{ marginTop: 7 }}><Sparkline bars={m.bars} delay={0.16 + n * 0.07} /></div>
                    </div>
                  ))}
                </div>
                <div style={{ flex: 1, minHeight: 0, display: "flex", gap: 12 }}>
                  <div style={{ flex: 1.15, minWidth: 0, padding: "13px 15px", borderRadius: 12, border: "1px solid #EFE6D9", background: "linear-gradient(180deg,#FFFCF6,#FFF7EC)", display: "flex", flexDirection: "column" }}>
                    <div style={{ flex: "none", display: "flex", alignItems: "baseline", justifyContent: "space-between", marginBottom: 9 }}>
                      <span style={{ fontSize: 13, fontWeight: 600, color: "#2A1B09" }}>Screens by format</span>
                      <span style={{ fontSize: 10, color: "#8A7F71" }}>631 screens · 6 formats</span>
                    </div>
                    <div style={{ flex: "none", display: "grid", gridTemplateColumns: "82px 1fr 34px 34px", gap: 10, padding: "0 2px 6px", borderBottom: "1px solid #F0E4D6", fontSize: 8, letterSpacing: "0.08em", textTransform: "uppercase", color: "#A38C78" }}>
                      <span>Format</span><span>Share</span><span style={{ textAlign: "right" }}>Scr</span><span style={{ textAlign: "right" }}>%</span>
                    </div>
                    <div style={{ flex: 1, minHeight: 0, display: "flex", flexDirection: "column", justifyContent: "center", gap: 9, paddingTop: 4 }}>
                      {sc.formats.map((f, n) => (
                        <div key={n} style={{ display: "grid", gridTemplateColumns: "82px 1fr 34px 34px", gap: 10, alignItems: "center", animation: "svf-rowIn 0.45s both", animationDelay: `${(0.12 + n * 0.05).toFixed(2)}s` }}>
                          <span style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 10.5, color: "#4D4034", minWidth: 0 }}>
                            <span style={{ width: 7, height: 7, flex: "none", borderRadius: 2, background: f.dot }} />
                            <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{f.k}</span>
                          </span>
                          <div style={{ minWidth: 0, height: 11, borderRadius: 6, background: "rgba(138,42,6,0.09)", overflow: "hidden" }}>
                            <div style={{ width: `${f.pct}%`, height: "100%", borderRadius: 6, background: f.color, transformOrigin: "left", animation: "svf-growX 0.85s cubic-bezier(.16,1,.3,1) both", animationDelay: `${(0.2 + n * 0.05).toFixed(2)}s` }} />
                          </div>
                          <span style={{ textAlign: "right", fontFamily: MONO, fontSize: 10.5, fontWeight: 600, color: "#2A1B09" }}>{f.v}</span>
                          <span style={{ textAlign: "right", fontFamily: MONO, fontSize: 10, color: "#8A7F71" }}>{f.share}</span>
                        </div>
                      ))}
                    </div>
                    <div style={{ flex: "none", display: "grid", gridTemplateColumns: "82px 1fr 34px 34px", gap: 10, alignItems: "center", marginTop: 8, paddingTop: 7, borderTop: "1px solid #F0E4D6" }}>
                      <span style={{ fontSize: 10, fontWeight: 600, color: "#2A1B09" }}>Total</span>
                      <span style={{ fontSize: 9.5, color: "#8A7F71" }}>2K QUBE leads</span>
                      <span style={{ textAlign: "right", fontFamily: MONO, fontSize: 10.5, fontWeight: 700, color: "#C2410C" }}>631</span>
                      <span style={{ textAlign: "right", fontFamily: MONO, fontSize: 10, color: "#8A7F71" }}>100%</span>
                    </div>
                  </div>
                  <div style={{ flex: 1, minWidth: 0, padding: "13px 15px", borderRadius: 12, border: "1px solid #EFE6D9", background: "#FFFCF6", display: "flex", flexDirection: "column" }}>
                    <span style={{ flex: "none", fontSize: 13, fontWeight: 600, color: "#2A1B09", marginBottom: 6 }}>Top centres</span>
                    <div style={{ flex: 1, minHeight: 0 }}>
                      {sc.centres.map((c, n) => (
                        <div key={n} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 8px", borderRadius: 7, transition: "background 0.3s, box-shadow 0.3s", animation: "svf-rowIn 0.4s both", animationDelay: `${(0.14 + n * 0.06).toFixed(2)}s`, ...rowHi(n) }}>
                          <span style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 11.5, color: "#3B2A18", minWidth: 0 }}>
                            <span style={{ width: 16, height: 16, flex: "none", borderRadius: "50%", display: "grid", placeItems: "center", background: n === 0 ? "#EA580C" : "#F0E0CF", fontFamily: MONO, fontSize: 8.5, fontWeight: 600, color: n === 0 ? "#fff" : "#8A6242" }}>{n + 1}</span>
                            <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{c.name}</span>
                          </span>
                          <span style={{ display: "flex", alignItems: "baseline", gap: 8, flex: "none" }}>
                            <span style={{ fontFamily: MONO, fontSize: 10.5, fontWeight: 600, color: "#C2410C" }}>{c.gross}</span>
                            <span style={{ fontSize: 9, color: "#8A7F71" }}>{c.occ}</span>
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </>
            )}

            {/* ═══════ COLLECTIONS ═══════ */}
            {sc.kind === "collections" && (
              <>
                <div style={{ flex: "none", display: "flex", gap: 10 }}>
                  {sc.kpis.map((k, n) => (
                    <div key={n} style={{ flex: 1, minWidth: 0, padding: "11px 13px", borderRadius: 10, border: "1px solid #EFE6D9", background: "#FFFCF6", animation: "svf-rowIn 0.45s both", animationDelay: `${(0.06 + n * 0.08).toFixed(2)}s` }}>
                      <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 6 }}>
                        <span style={{ fontSize: 8.5, letterSpacing: "0.12em", textTransform: "uppercase", color: "#8A6242" }}>{k.k}</span>
                        <span style={{ fontFamily: MONO, fontSize: 9.5, color: k.up ? "#05502F" : "#C2410C" }}>{k.delta}</span>
                      </div>
                      <div style={{ marginTop: 5, fontFamily: MONO, fontSize: 20, fontWeight: 600, color: "#2A1B09", animation: "svf-countPulse 0.6s both", animationDelay: `${(0.12 + n * 0.08).toFixed(2)}s` }}>{k.v}</div>
                      <div style={{ marginTop: 8 }}><Sparkline bars={k.bars} delay={0.16 + n * 0.08} /></div>
                    </div>
                  ))}
                </div>
                <div style={{ flex: 1, minHeight: 0, display: "flex", flexDirection: "column", padding: "13px 15px", borderRadius: 12, border: "1px solid #EFE6D9", background: "linear-gradient(180deg,#FFFCF6,#FFF7EC)" }}>
                  <div style={{ flex: "none", display: "flex", alignItems: "baseline", justifyContent: "space-between", marginBottom: 10 }}>
                    <span style={{ fontSize: 13.5, fontWeight: 600, color: "#2A1B09" }}>Collections by district</span>
                    <span style={{ fontSize: 10.5, color: "#8A7F71" }}>Today · 147 reporting</span>
                  </div>
                  <div style={{ flex: 1, minHeight: 0, display: "flex", flexDirection: "column", justifyContent: "space-between", gap: 2 }}>
                    {sc.districts.map((d, n) => (
                      <div key={n} style={{ display: "flex", alignItems: "center", gap: 10, padding: "0 4px", borderRadius: 7, transition: "background 0.3s, box-shadow 0.3s", animation: "svf-rowIn 0.45s both", animationDelay: `${(0.12 + n * 0.06).toFixed(2)}s`, ...rowHi(n) }}>
                        <span style={{ width: 92, flex: "none", fontSize: 11, color: n === 0 ? "#2A1B09" : "#6B5D4C", fontWeight: n === 0 ? 600 : 400, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{d.name}</span>
                        <div style={{ flex: 1, minWidth: 0, height: 15, borderRadius: 8, background: "rgba(138,42,6,0.09)", overflow: "hidden" }}>
                          <div style={{ width: `${d.pct}%`, height: "100%", borderRadius: 8, background: n === 0 ? "linear-gradient(90deg,#F97316,#EA580C)" : "linear-gradient(90deg,#F6A968,#E8834A)", transformOrigin: "left", animation: "svf-growX 0.85s cubic-bezier(.16,1,.3,1) both", animationDelay: `${(0.2 + n * 0.06).toFixed(2)}s` }} />
                        </div>
                        <span style={{ width: 44, flex: "none", textAlign: "right", fontSize: 9.5, color: "#8A7F71" }}>{d.occ}</span>
                        <span style={{ width: 62, flex: "none", textAlign: "right", fontFamily: MONO, fontSize: 11, fontWeight: 500, color: "#2A1B09" }}>{d.gross}</span>
                      </div>
                    ))}
                  </div>
                </div>
                <div style={{ flex: "none", display: "flex", gap: 10 }}>
                  {sc.theatres.map((t, n) => (
                    <div key={n} style={{ flex: 1, minWidth: 0, padding: "9px 12px", borderRadius: 9, border: "1px solid #EFE6D9", background: "#FFFCF6", animation: "svf-rowIn 0.4s both", animationDelay: `${(0.24 + n * 0.06).toFixed(2)}s` }}>
                      <div style={{ fontSize: 11, fontWeight: 600, color: "#2A1B09", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{t.name}</div>
                      <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", marginTop: 3, gap: 6 }}>
                        <span style={{ fontSize: 9.5, color: "#9A8E80", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{t.centre}</span>
                        <span style={{ flex: "none", fontFamily: MONO, fontSize: 11, fontWeight: 600, color: "#C2410C" }}>{t.v}</span>
                      </div>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>

          {/* cursor — glides to the highlighted row (or up to the next tab), then "taps" */}
          <div style={{ position: "absolute", left: cursorPos ? cursorPos[0] : "-12%", top: cursorPos ? cursorPos[1] : "50%", width: 0, height: 0, zIndex: 12, transition: "left 0.75s cubic-bezier(.5,0,.2,1), top 0.75s cubic-bezier(.5,0,.2,1)", pointerEvents: "none" }}>
            <div style={{ position: "absolute", left: -12, top: -12, width: 24, height: 24, borderRadius: "50%", background: "rgba(228,88,26,0.5)", animation: "svf-ripple 1.6s ease-out infinite" }} />
            <svg width="19" height="23" viewBox="0 0 20 24" style={{ position: "absolute", left: -2, top: -2, filter: "drop-shadow(0 3px 5px rgba(28,10,3,0.4))", animation: "svf-cursorTap 1.6s ease-in-out infinite" }}>
              <path d="M2 1 L2 19 L7 14.6 L10.6 22 L13.6 20.6 L10 13.4 L16.6 13 Z" fill="#2A0B03" stroke="#FFFDFA" strokeWidth="1.4" />
            </svg>
          </div>
        </div>

        {/* dim scrim — the box recedes behind the card that lifts out */}
        <div style={{ position: "absolute", inset: 0, zIndex: 9, borderRadius: 16, pointerEvents: "none", background: "linear-gradient(135deg, rgba(28,10,3,0.16), rgba(28,10,3,0.34))", opacity: focus ? 1 : 0, transition: "opacity 0.5s ease" }} />
      </div>

      {/* detail card — pops out of the box and docks at the bottom-right corner.
          Re-keyed per focus step so a fresh card animates in for each row. */}
      {detail && (
        <div key={"d" + step + "-" + fi} style={{ position: "absolute", right: -18, bottom: -18, width: 238, zIndex: 20, borderRadius: 13, overflow: "hidden", background: "#FFFDFA", boxShadow: "0 26px 52px rgba(28,10,3,0.4), 0 0 0 1px rgba(43,20,3,0.06)", animation: "svf-cornerIn 0.5s cubic-bezier(.16,1,.3,1) both", fontFamily: SANS }}>
          <div style={{ padding: "11px 14px", background: "linear-gradient(96deg,#8A2A06,#C2410C)" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
              <span style={{ fontFamily: MONO, fontSize: 8.5, letterSpacing: "0.1em", color: "rgba(255,239,226,0.82)" }}>{detail.code}</span>
              <span style={{ padding: "2px 7px", borderRadius: 4, background: "rgba(255,246,238,0.18)", fontSize: 8, letterSpacing: "0.08em", textTransform: "uppercase", color: "#FFF6EE" }}>{detail.tag}</span>
            </div>
            <div style={{ marginTop: 3, fontSize: 14, fontWeight: 600, color: "#FFF6EE" }}>{detail.title}</div>
          </div>
          <div style={{ padding: "10px 14px 13px" }}>
            {detail.lines.map((l, n) => (
              <div key={n} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "5px 0", borderBottom: "1px solid #F5EEE3", animation: "svf-rowIn 0.4s both", animationDelay: `${(0.16 + n * 0.06).toFixed(2)}s` }}>
                <span style={{ fontSize: 10.5, color: "#7A6E60" }}>{l.k}</span>
                <span style={{ fontFamily: MONO, fontSize: 10.5, fontWeight: 500, color: l.tint ?? "#2A1B09" }}>{l.v}</span>
              </div>
            ))}
            <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", marginTop: 9, paddingTop: 9, borderTop: "2px solid rgba(43,20,3,0.22)" }}>
              <span style={{ fontSize: 8, letterSpacing: "0.14em", textTransform: "uppercase", color: "#8A6242" }}>{detail.foot.k}</span>
              <span style={{ fontFamily: MONO, fontSize: 17, fontWeight: 600, color: "#2A1B09" }}>{detail.foot.v}</span>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default function LoginPage() {
  const router = useRouter();
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  async function submit(e: React.FormEvent) {
    e.preventDefault();
    setError("");
    setLoading(true);
    try {
      const res = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email, password }),
      });
      if (res.ok) {
        const data = await res.json().catch(() => ({}));
        const dest = data.redirect || "/admin/theatres";
        if (data.role === "rep") {
          window.location.href = dest;
        } else {
          router.replace(dest);
          router.refresh();
        }
      } else {
        const data = await res.json().catch(() => ({}));
        setError(data.error || "Login failed");
      }
    } catch {
      setError("Something went wrong. Try again.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="flex min-h-screen flex-col lg:flex-row">
      {/* ── Left: live SVF app preview ── */}
      <aside
        className="relative flex w-full flex-col overflow-hidden lg:w-[54%]"
        style={{ backgroundImage: "linear-gradient(135deg,#7c2d12 0%,#ea580c 48%,#f59e0b 100%)" }}
      >
        <div className="pointer-events-none absolute inset-0" style={{ background: "radial-gradient(66% 52% at 50% 40%, rgba(255,226,193,0.35), transparent 78%)" }} />
        <div className="pointer-events-none absolute inset-0" style={{ background: "linear-gradient(198deg, rgba(120,40,4,0.18), transparent 46%, rgba(120,40,4,0.32))" }} />

        <div className="relative z-10 flex items-center gap-3 px-6 pt-8 sm:px-12 sm:pt-10">
          <span className="grid h-10 w-10 place-items-center rounded-lg bg-white/15 text-sm font-bold tracking-tight text-white ring-1 ring-white/30 backdrop-blur">SVF</span>
          <div className="leading-tight">
            <p className="text-base font-semibold text-white">SVF Distribution</p>
            <p className="text-xs text-white/70">Nizam Territory</p>
          </div>
        </div>

        <div className="relative z-10 flex flex-1 items-center justify-center px-4 py-8 sm:px-10">
          <div className="w-full max-w-[780px]">
            <ScaledStage dw={782} dh={524}>
              <DistributionApp />
            </ScaledStage>
          </div>
        </div>

        <div className="relative z-10 max-w-lg px-6 pb-8 sm:px-12 sm:pb-10">
          <h2 className="text-xl font-bold leading-snug tracking-tight text-[#3A0F05] sm:text-2xl">
            Every screen. Every collection. One dashboard.
          </h2>
          <p className="mt-2.5 text-sm leading-relaxed text-[#3A0F05]/80">
            Theatres, rate cards and live box-office across the Nizam territory — managed end to end.
          </p>
        </div>
      </aside>

      {/* ── Right: sign-in form ── */}
      <main className="flex w-full flex-1 items-center justify-center p-6 lg:w-[46%]" style={{ backgroundColor: "#fff8f1" }}>
        <div className="w-full max-w-sm">
          <div className="mb-8 flex items-center gap-3">
            <span className="grid h-10 w-10 place-items-center rounded-lg text-sm font-bold tracking-tight text-white shadow-sm" style={{ backgroundImage: "linear-gradient(135deg,#fb923c,#ea580c)" }}>SVF</span>
            <div className="leading-tight">
              <p className="text-base font-semibold" style={{ color: "#3d2415" }}>SVF Distribution</p>
              <p className="text-xs" style={{ color: "#a07a60" }}>Nizam Territory</p>
            </div>
          </div>

          <h1 className="text-2xl font-bold tracking-tight" style={{ color: "#3d2415" }}>Welcome back</h1>
          <p className="mb-6 mt-1 text-sm" style={{ color: "#8a6a53" }}>Sign in to continue to Cine Distribution.</p>

          <form onSubmit={submit} className="space-y-4">
            <div>
              <label className="mb-1.5 block text-xs font-medium" style={{ color: "#6b4a35" }}>Email Address</label>
              <input
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="you@svf.in"
                autoComplete="username"
                className="w-full rounded-lg border bg-white px-3.5 py-2.5 text-sm text-[#3d2415] outline-none transition placeholder:text-[#c4a68f] focus:border-[#ea580c] focus:ring-2 focus:ring-[#f97316]/25"
                style={{ borderColor: "#f0dcc9" }}
              />
            </div>

            <div>
              <label className="mb-1.5 block text-xs font-medium" style={{ color: "#6b4a35" }}>Password</label>
              <div className="relative">
                <input
                  type={showPw ? "text" : "password"}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••"
                  autoComplete="current-password"
                  className="w-full rounded-lg border bg-white px-3.5 py-2.5 pr-10 text-sm text-[#3d2415] outline-none transition placeholder:text-[#c4a68f] focus:border-[#ea580c] focus:ring-2 focus:ring-[#f97316]/25"
                  style={{ borderColor: "#f0dcc9" }}
                />
                <button
                  type="button"
                  onClick={() => setShowPw((v) => !v)}
                  tabIndex={-1}
                  title={showPw ? "Hide password" : "Show password"}
                  className="absolute inset-y-0 right-0 grid w-10 place-items-center text-[#b58e73] hover:text-[#ea580c]"
                >
                  {showPw ? (
                    <svg viewBox="0 0 24 24" className="h-4 w-4" fill="none" stroke="currentColor" strokeWidth={1.8} strokeLinecap="round" strokeLinejoin="round">
                      <path d="M3 3l18 18M10.6 10.6a2 2 0 0 0 2.8 2.8" />
                      <path d="M9.4 5.2A9.5 9.5 0 0 1 12 5c5 0 9 4.5 10 7-.4 1-.9 1.9-1.6 2.7M6.1 6.1C3.9 7.4 2.4 9.5 2 12c1 2.5 5 7 10 7 1.5 0 2.9-.4 4.1-1" />
                    </svg>
                  ) : (
                    <svg viewBox="0 0 24 24" className="h-4 w-4" fill="none" stroke="currentColor" strokeWidth={1.8} strokeLinecap="round" strokeLinejoin="round">
                      <path d="M2 12s4-7 10-7 10 7 10 7-4 7-10 7-10-7-10-7Z" />
                      <circle cx="12" cy="12" r="3" />
                    </svg>
                  )}
                </button>
              </div>
            </div>

            {error && (
              <p className="rounded-lg bg-rose-50 px-3 py-2 text-xs font-medium text-rose-600">{error}</p>
            )}

            <button
              type="submit"
              disabled={loading}
              className="w-full rounded-lg py-2.5 text-sm font-semibold text-white shadow-md transition hover:brightness-95 active:brightness-90 disabled:cursor-wait disabled:opacity-80"
              style={{ backgroundImage: "linear-gradient(135deg,#f97316,#ea580c)" }}
            >
              {loading ? "Signing in…" : "Sign In"}
            </button>
          </form>

          <div className="mt-10 border-t pt-5 text-center" style={{ borderColor: "#f0dcc9" }}>
            <p className="text-[10px] font-semibold uppercase tracking-[0.18em]" style={{ color: "#b89a86" }}>Powered by</p>
            <img src="/lorven-logo.png" alt="Lorven AI Studio" className="mx-auto mt-2 h-8 w-auto" />
            <p className="mt-4 text-[11px]" style={{ color: "#a07a60" }}>© 2026 Lorven AI Mediavision LLP</p>
          </div>
        </div>
      </main>
    </div>
  );
}
