"use client";

import { useMemo, useState } from "react";
import { Theatre, DistrictSummary, districtsOf, inr } from "@/data/theatres";
import { useReps, repMap } from "@/lib/reps";
import { useAssignments } from "@/lib/assignments";
import {
  useRateOverrides,
  screenKey,
  setScreenSlabs,
  resetScreen,
  slabsHfc,
  slabsSeats,
  toNum,
} from "@/lib/rates";
import RepAvatar from "./RepAvatar";
import AssignRepModal from "./AssignRepModal";

type CentreNode = { name: string; theatres: Theatre[] };
type DistrictNode = { name: string; total: number; centres: CentreNode[] };

function buildTree(
  theatres: Theatre[],
  districts: DistrictSummary[]
): DistrictNode[] {
  return districts.map((d) => {
    const rows = theatres.filter((t) => t.district === d.name);
    const byCentre = new Map<string, Theatre[]>();
    for (const t of rows) {
      const arr = byCentre.get(t.centre) || [];
      arr.push(t);
      byCentre.set(t.centre, arr);
    }
    return {
      name: d.name,
      total: rows.length,
      centres: [...byCentre.entries()].map(([name, list]) => ({
        name,
        theatres: list,
      })),
    };
  });
}

export default function TheatreExplorer({
  theatres,
}: {
  theatres: Theatre[];
}) {
  const districts = useMemo(() => districtsOf(theatres), [theatres]);
  const tree = useMemo(
    () => buildTree(theatres, districts),
    [theatres, districts]
  );
  const assignments = useAssignments();
  const reps = useReps();
  const byId = useMemo(() => repMap(reps), [reps]);

  const [query, setQuery] = useState("");
  const [openDistricts, setOpenDistricts] = useState<Set<string>>(
    () => new Set([tree[0]?.name])
  );
  const [openCentres, setOpenCentres] = useState<Set<string>>(new Set());
  const [selectedId, setSelectedId] = useState<number | null>(null);
  const [limits, setLimits] = useState<Record<string, number>>({});

  const PAGE = 10;
  const limitFor = (d: string) => limits[d] ?? PAGE;

  const q = query.trim().toLowerCase();
  const selected = selectedId
    ? theatres.find((t) => t.id === selectedId) ?? null
    : null;

  function toggle(set: Set<string>, key: string, setter: (s: Set<string>) => void) {
    const n = new Set(set);
    n.has(key) ? n.delete(key) : n.add(key);
    setter(n);
  }

  function matches(t: Theatre) {
    if (!q) return true;
    return (
      t.theatre.toLowerCase().includes(q) ||
      t.centre.toLowerCase().includes(q) ||
      t.district.toLowerCase().includes(q)
    );
  }

  return (
    <div className="flex h-[calc(100vh-150px)] overflow-hidden rounded-lg border border-line bg-surface">
      {/* ---- Tree ---- */}
      <div className="flex w-[340px] shrink-0 flex-col border-r border-line">
        <div className="border-b border-line p-3">
          <input
            value={query}
            onChange={(e) => {
              setQuery(e.target.value);
              if (e.target.value) {
                const od = new Set<string>();
                const oc = new Set<string>();
                for (const d of tree) {
                  for (const c of d.centres) {
                    if (c.theatres.some(matches)) {
                      od.add(d.name);
                      oc.add(`${d.name}//${c.name}`);
                    }
                  }
                }
                setOpenDistricts(od);
                setOpenCentres(oc);
              }
            }}
            placeholder="Search district, centre or theatre…"
            className="w-full rounded-md border border-line px-3 py-2 text-sm outline-none focus:border-brand-400"
          />
        </div>

        <div className="flex-1 overflow-y-auto py-1">
          {tree.map((d) => {
            const districtCentres = q
              ? d.centres.filter((c) => c.theatres.some(matches))
              : d.centres;
            if (q && districtCentres.length === 0) return null;
            const dOpen = openDistricts.has(d.name);
            const assigned = theatres.filter(
              (t) => t.district === d.name && assignments[t.id]
            ).length;

            return (
              <div key={d.name}>
                <button
                  onClick={() => toggle(openDistricts, d.name, setOpenDistricts)}
                  className="flex w-full items-center gap-1.5 px-2 py-1.5 text-left hover:bg-muted"
                >
                  <Chevron open={dOpen} />
                  <span className="flex-1 text-sm font-semibold text-strong">
                    {d.name || "Unassigned district"}
                  </span>
                  <span className="rounded bg-chip px-1.5 py-0.5 text-[10px] font-medium text-faint">
                    {assigned}/{d.total}
                  </span>
                </button>

                {dOpen && (
                  <div className="ml-[15px] border-l border-line">
                    {(q
                      ? districtCentres
                      : districtCentres.slice(0, limitFor(d.name))
                    ).map((c) => {
                      const key = `${d.name}//${c.name}`;
                      const cOpen = openCentres.has(key) || !!q;
                      const centreTheatres = q
                        ? c.theatres.filter(matches)
                        : c.theatres;
                      return (
                        <div key={key}>
                          <button
                            onClick={() => toggle(openCentres, key, setOpenCentres)}
                            className="flex w-full items-center gap-1.5 py-1 pl-3 pr-2 text-left hover:bg-muted"
                          >
                            <Chevron open={cOpen} small />
                            <span className="flex-1 truncate text-[13px] font-medium text-body">
                              {c.name || "Unassigned centre"}
                            </span>
                            <span className="text-[10px] text-faint">
                              {c.theatres.length}
                            </span>
                          </button>

                          {cOpen && (
                            <div className="ml-[15px] border-l border-line">
                              {centreTheatres.map((t) => {
                                const rep = assignments[t.id]
                                  ? byId[assignments[t.id]]
                                  : undefined;
                                const active = selectedId === t.id;
                                return (
                                  <button
                                    key={t.id}
                                    onClick={() => setSelectedId(t.id)}
                                    className={`flex w-full items-center gap-2 py-1 pl-4 pr-2 text-left text-[13px] ${
                                      active
                                        ? "bg-brand-50 text-brand-700"
                                        : "text-body hover:bg-muted"
                                    }`}
                                  >
                                    <TypeDot type={t.type} />
                                    <span className="flex-1 truncate">
                                      {t.theatre}
                                    </span>
                                    {rep ? (
                                      <RepAvatar
                                        name={rep.name}
                                        color={rep.color}
                                        size={16}
                                      />
                                    ) : (
                                      <span className="h-1.5 w-1.5 rounded-full bg-neutral-300" />
                                    )}
                                  </button>
                                );
                              })}
                            </div>
                          )}
                        </div>
                      );
                    })}

                    {!q && districtCentres.length > limitFor(d.name) && (
                      <button
                        onClick={() =>
                          setLimits((l) => ({
                            ...l,
                            [d.name]: limitFor(d.name) + PAGE,
                          }))
                        }
                        className="w-full py-1.5 pl-3 pr-2 text-left text-xs font-medium text-brand-700 hover:bg-muted"
                      >
                        Load more · {districtCentres.length - limitFor(d.name)}{" "}
                        remaining
                      </button>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>

      {/* ---- Detail ---- */}
      <div className="flex-1 overflow-y-auto">
        {selected ? (
          <TheatreDetail
            t={selected}
            repId={assignments[selected.id]}
            rep={assignments[selected.id] ? byId[assignments[selected.id]] : undefined}
          />
        ) : (
          <div className="flex h-full flex-col items-center justify-center gap-2 p-10 text-center">
            <div className="grid h-12 w-12 place-items-center rounded-full bg-chip text-faint">
              <svg viewBox="0 0 24 24" className="h-6 w-6" fill="none" stroke="currentColor" strokeWidth={1.6}>
                <path d="M4 4h16v4H4V4Zm0 6h16v10H4V10Zm3 3v4m5-4v4m5-4v4" />
              </svg>
            </div>
            <p className="text-sm font-medium text-body">Select a theatre</p>
            <p className="max-w-xs text-xs text-faint">
              Browse the tree by district and centre, then pick a theatre to see its
              screens, rates and representative.
            </p>
          </div>
        )}
      </div>
    </div>
  );
}

function TheatreDetail({
  t,
  repId,
  rep,
}: {
  t: Theatre;
  repId?: string;
  rep?: { name: string; color: string; region: string };
}) {
  const [modal, setModal] = useState(false);
  const overrides = useRateOverrides();

  // effective per-screen slabs (override or default) + recomputed totals
  const eff = t.screens.map((s, i) => {
    const slabs = overrides[screenKey(t.id, i)] ?? s.slabs;
    return { slabs, hfc: slabsHfc(slabs), seats: slabsSeats(slabs) };
  });
  const totalSeats = eff.reduce((n, e) => n + e.seats, 0);
  const totalHfc = eff.reduce((n, e) => n + e.hfc, 0);

  return (
    <div className="p-6">
      <p className="text-xs text-faint">
        {t.district} <span className="mx-1">/</span> {t.centre}
      </p>
      <div className="mt-1 flex flex-wrap items-center gap-3">
        <h3 className="text-xl font-semibold tracking-tight text-strong">
          {t.theatre}
        </h3>
        <span
          className={`rounded px-2 py-0.5 text-[11px] font-medium capitalize ${
            t.type === "multiplex"
              ? "bg-brand-50 text-brand-700"
              : "bg-sky-50 text-sky-700 dark:bg-sky-500/15 dark:text-sky-300"
          }`}
        >
          {t.type}
        </span>
      </div>

      <div className="mt-4 grid grid-cols-2 gap-3 sm:grid-cols-4">
        <MiniStat
          label="Screens"
          value={t.screen_count}
          icon={<path d="M3 4h18v12H3zM8 20h8M12 16v4" />}
        />
        <MiniStat
          label="Total seats"
          value={inr(totalSeats)}
          icon={<path d="M4 18v-5a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v5M6 11V7a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v4M5 18v2m14-2v2" />}
        />
        <MiniStat
          label="Full house"
          value={`₹${inr(totalHfc)}`}
          valueClass="text-emerald-600 dark:text-emerald-400"
          icon={<><circle cx="12" cy="12" r="8" /><path d="M12 8v8m-2-5.5a2 2 0 0 1 4 0M10 13.5a2 2 0 0 0 4 0" /></>}
        />
        <MiniStat
          label="Format"
          value={t.format || "—"}
          icon={<><rect x="3" y="4" width="18" height="16" rx="2" /><path d="M7 4v16M17 4v16M3 9h4m10 0h4M3 15h4m10 0h4" /></>}
        />
      </div>

      <div className="mt-5 flex items-center justify-between rounded-lg border border-line bg-muted px-4 py-3">
        <div className="flex items-center gap-3">
          {rep ? (
            <>
              <RepAvatar name={rep.name} color={rep.color} size={34} />
              <div className="leading-tight">
                <p className="text-sm font-medium text-strong">{rep.name}</p>
                <p className="text-xs text-faint">{rep.region}</p>
              </div>
            </>
          ) : (
            <div className="flex items-center gap-3">
              <span className="grid h-[34px] w-[34px] place-items-center rounded-full bg-chip text-faint">
                ?
              </span>
              <p className="text-sm text-faint">No representative assigned</p>
            </div>
          )}
        </div>
        <button
          onClick={() => setModal(true)}
          className="inline-flex items-center rounded-md bg-brand-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-brand-700"
        >
          {rep ? "Change representative" : "Assign representative"}
        </button>
      </div>

      <div className="mb-2 mt-6 flex items-center justify-between">
        <h4 className="text-sm font-semibold text-strong">Screens &amp; rates</h4>
        <span className="text-xs text-faint">
          Click a screen&apos;s rate slabs to edit — full house recalculates.
        </span>
      </div>
      <ScreensTable t={t} />

      {modal && (
        <AssignRepModal
          theatreId={t.id}
          theatreName={`${t.theatre} · ${t.centre}`}
          currentRepId={repId}
          onClose={() => setModal(false)}
        />
      )}
    </div>
  );
}

function ScreensTable({ t }: { t: Theatre }) {
  const overrides = useRateOverrides();
  const [editing, setEditing] = useState<number | null>(null);
  const [draft, setDraft] = useState<{ rate: string; aud: string }[]>([]);

  const effSlabs = (i: number) =>
    overrides[screenKey(t.id, i)] ?? t.screens[i].slabs;
  const isOverridden = (i: number) => !!overrides[screenKey(t.id, i)];

  function start(i: number) {
    setEditing(i);
    setDraft(
      effSlabs(i).map((s) => ({
        rate: s.rate == null ? "" : String(s.rate),
        aud: s.aud == null ? "" : String(s.aud),
      }))
    );
  }
  function cancel() {
    setEditing(null);
    setDraft([]);
  }
  function save(i: number) {
    const parsed = draft
      .map((d) => ({ rate: toNum(d.rate), aud: toNum(d.aud) }))
      .filter((d) => d.rate || d.aud);
    setScreenSlabs(t.id, i, parsed);
    cancel();
  }

  const draftHfc = draft.reduce((n, d) => n + toNum(d.rate) * toNum(d.aud), 0);
  const draftSeats = draft.reduce((n, d) => n + toNum(d.aud), 0);

  return (
    <div className="overflow-x-auto rounded-lg border border-line">
      <table className="w-full min-w-[560px] text-sm">
        <thead>
          <tr className="border-b border-line bg-muted text-left text-xs uppercase tracking-wide text-faint [&>th]:border-l [&>th]:border-line [&>th:first-child]:border-l-0">
            <th className="px-3 py-2.5 font-medium">Screen</th>
            <th className="px-3 py-2.5 font-medium">Format</th>
            <th className="px-3 py-2.5 text-right font-medium">Seats</th>
            <th className="px-3 py-2.5 text-right font-medium">Full House (₹)</th>
            <th className="px-3 py-2.5 font-medium">Rate slabs</th>
          </tr>
        </thead>
        <tbody className="[&>tr]:border-b [&>tr]:border-line [&>tr:last-child]:border-0 [&_td]:border-l [&_td]:border-line [&_td:first-child]:border-l-0">
          {t.screens.map((s, i) => {
            const editingThis = editing === i;
            const slabs = effSlabs(i);
            const seats = editingThis ? draftSeats : slabsSeats(slabs);
            const hfc = editingThis ? draftHfc : slabsHfc(slabs);
            return (
              <tr
                key={i}
                className={
                  editingThis
                    ? "bg-brand-50/60"
                    : i % 2
                    ? "bg-muted/40"
                    : "bg-transparent"
                }
              >
                <td className="px-3 py-2.5 align-top font-medium text-strong">
                  {s.name || `Screen ${i + 1}`}
                  {isOverridden(i) && (
                    <span
                      title="Edited"
                      className="ml-1.5 inline-block h-1.5 w-1.5 rounded-full bg-brand-500 align-middle"
                    />
                  )}
                </td>
                <td className="px-3 py-2.5 align-top">
                  {s.format ? (
                    <span className="rounded bg-chip px-1.5 py-0.5 text-[11px] font-medium text-body">
                      {s.format}
                    </span>
                  ) : (
                    <span className="text-faint">—</span>
                  )}
                </td>
                <td className="px-3 py-2.5 text-right align-top text-body">
                  {inr(seats)}
                </td>
                <td className="px-3 py-2.5 text-right align-top font-medium text-emerald-600 dark:text-emerald-400">
                  {inr(hfc)}
                </td>
                <td className="px-3 py-2 align-top">
                  {editingThis ? (
                    <div className="space-y-1.5">
                      {draft.map((d, di) => (
                        <div key={di} className="flex items-center gap-1.5">
                          <span className="text-faint">₹</span>
                          <input
                            type="number"
                            value={d.rate}
                            onChange={(e) =>
                              setDraft((dr) =>
                                dr.map((x, xi) =>
                                  xi === di ? { ...x, rate: e.target.value } : x
                                )
                              )
                            }
                            className="w-20 rounded border border-line bg-surface px-1.5 py-1 text-xs outline-none focus:border-brand-400"
                          />
                          <span className="text-faint">×</span>
                          <input
                            type="number"
                            value={d.aud}
                            onChange={(e) =>
                              setDraft((dr) =>
                                dr.map((x, xi) =>
                                  xi === di ? { ...x, aud: e.target.value } : x
                                )
                              )
                            }
                            className="w-20 rounded border border-line bg-surface px-1.5 py-1 text-xs outline-none focus:border-brand-400"
                          />
                          <button
                            title="Remove slab"
                            onClick={() =>
                              setDraft((dr) => dr.filter((_, xi) => xi !== di))
                            }
                            className="rounded p-1 text-faint hover:bg-chip hover:text-rose-500"
                          >
                            <svg viewBox="0 0 24 24" className="h-3.5 w-3.5" fill="none" stroke="currentColor" strokeWidth={2}>
                              <path d="M6 6l12 12M18 6L6 18" />
                            </svg>
                          </button>
                        </div>
                      ))}
                      <div className="flex items-center gap-2 pt-1">
                        <button
                          onClick={() =>
                            setDraft((dr) => [...dr, { rate: "", aud: "" }])
                          }
                          className="rounded border border-dashed border-line px-2 py-1 text-[11px] font-medium text-body hover:border-brand-400 hover:text-brand-700"
                        >
                          + Add slab
                        </button>
                        <span className="flex-1" />
                        {isOverridden(i) && (
                          <button
                            onClick={() => {
                              resetScreen(t.id, i);
                              cancel();
                            }}
                            className="rounded px-2 py-1 text-[11px] font-medium text-faint hover:bg-chip"
                          >
                            Reset
                          </button>
                        )}
                        <button
                          onClick={cancel}
                          className="rounded px-2 py-1 text-[11px] font-medium text-faint hover:bg-chip"
                        >
                          Cancel
                        </button>
                        <button
                          onClick={() => save(i)}
                          className="rounded bg-brand-600 px-2.5 py-1 text-[11px] font-semibold text-white hover:bg-brand-700"
                        >
                          Save
                        </button>
                      </div>
                    </div>
                  ) : (
                    <button
                      onClick={() => start(i)}
                      className="group flex w-full items-center gap-2 text-left text-xs text-body"
                    >
                      <span>
                        {slabs.length
                          ? slabs
                              .map((sl) => `₹${sl.rate} × ${sl.aud}`)
                              .join("  ·  ")
                          : "—"}
                      </span>
                      <svg
                        viewBox="0 0 24 24"
                        className="h-3.5 w-3.5 shrink-0 text-faint opacity-0 transition group-hover:opacity-100"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth={1.8}
                      >
                        <path d="M4 20h4L18 10l-4-4L4 16v4Zm11-13 3 3" />
                      </svg>
                    </button>
                  )}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function MiniStat({
  label,
  value,
  icon,
  valueClass,
}: {
  label: string;
  value: string | number;
  icon?: React.ReactNode;
  valueClass?: string;
}) {
  return (
    <div className="rounded-lg bg-muted px-3 py-2.5">
      <p className="flex items-center gap-1.5 text-[11px] uppercase tracking-wide text-faint">
        {icon && (
          <svg
            viewBox="0 0 24 24"
            className="h-3.5 w-3.5"
            fill="none"
            stroke="currentColor"
            strokeWidth={1.7}
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            {icon}
          </svg>
        )}
        {label}
      </p>
      <p className={`mt-0.5 text-base font-semibold ${valueClass || "text-strong"}`}>
        {value}
      </p>
    </div>
  );
}

function Chevron({ open, small }: { open: boolean; small?: boolean }) {
  return (
    <svg
      viewBox="0 0 24 24"
      className={`${small ? "h-3.5 w-3.5" : "h-4 w-4"} shrink-0 text-faint transition-transform ${
        open ? "rotate-90" : ""
      }`}
      fill="none"
      stroke="currentColor"
      strokeWidth={2.2}
    >
      <path d="M9 6l6 6-6 6" />
    </svg>
  );
}

function TypeDot({ type }: { type: "multiplex" | "single" }) {
  return (
    <span
      className={`h-2 w-2 shrink-0 rounded-[3px] ${
        type === "multiplex" ? "bg-brand-500" : "bg-neutral-300"
      }`}
      title={type}
    />
  );
}
