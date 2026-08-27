"use client";

import { useEffect, useMemo, useState } from "react";

type HistoryRow = {
  id: string;
  email: string;
  role: string;
  ip: string;
  device: string;
  location: string;
  createdAt: number;
};

const ROLE_STYLE: Record<string, string> = {
  admin: "bg-brand-50 text-brand-700",
  user: "bg-sky-50 text-sky-700",
  rep: "bg-amber-50 text-amber-700",
};

function roleClass(role: string) {
  return ROLE_STYLE[role] || "bg-chip text-faint";
}

function relativeTime(ts: number) {
  const diff = Date.now() - ts;
  const s = Math.round(diff / 1000);
  if (s < 60) return "just now";
  const m = Math.round(s / 60);
  if (m < 60) return `${m}m ago`;
  const h = Math.round(m / 60);
  if (h < 24) return `${h}h ago`;
  const d = Math.round(h / 24);
  if (d < 30) return `${d}d ago`;
  return new Date(ts).toLocaleDateString();
}

export default function HistoryPage() {
  const [rows, setRows] = useState<HistoryRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");
  const [query, setQuery] = useState("");
  const [role, setRole] = useState<"all" | "admin" | "user" | "rep">("all");

  async function load() {
    setLoading(true);
    setErr("");
    try {
      const res = await fetch("/api/history", { cache: "no-store" });
      if (!res.ok) {
        setErr((await res.json().catch(() => ({})))?.error || "Could not load history.");
        setRows([]);
      } else {
        setRows((await res.json()).history ?? []);
      }
    } catch {
      setErr("Could not load history.");
    } finally {
      setLoading(false);
    }
  }
  useEffect(() => {
    load();
  }, []);

  async function clearAll() {
    if (!confirm("Clear the entire login history? This cannot be undone.")) return;
    await fetch("/api/history", { method: "DELETE" });
    load();
  }

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return rows.filter((r) => {
      if (role !== "all" && r.role !== role) return false;
      if (!q) return true;
      return (
        r.email.toLowerCase().includes(q) ||
        r.device.toLowerCase().includes(q) ||
        r.location.toLowerCase().includes(q) ||
        r.ip.toLowerCase().includes(q)
      );
    });
  }, [rows, query, role]);

  const counts = useMemo(() => {
    const c = { admin: 0, user: 0, rep: 0 };
    for (const r of rows) if (r.role in c) c[r.role as keyof typeof c]++;
    return c;
  }, [rows]);

  return (
    <div className="space-y-6">
      <div className="flex items-start justify-between gap-4">
        <div>
          <h2 className="text-xl font-semibold tracking-tight text-strong">
            Login history
          </h2>
          <p className="mt-0.5 text-sm text-faint">
            Every successful sign-in — who logged in, their role, the device and
            browser used, where they connected from, and when.
          </p>
        </div>
        <div className="flex shrink-0 gap-2">
          <button
            onClick={load}
            className="rounded-md border border-line px-3 py-2 text-sm font-medium text-body hover:bg-chip"
          >
            ↻ Refresh
          </button>
          <button
            onClick={clearAll}
            disabled={rows.length === 0}
            className="rounded-md border border-line px-3 py-2 text-sm font-medium text-rose-600 hover:bg-rose-50 disabled:opacity-40"
          >
            Clear history
          </button>
        </div>
      </div>

      {/* Summary + filters */}
      <div className="flex flex-wrap items-center gap-3">
        <div className="flex items-center gap-1.5 rounded-lg border border-line bg-surface p-1 text-sm">
          {(
            [
              ["all", `All ${rows.length}`],
              ["admin", `Admin ${counts.admin}`],
              ["user", `Users ${counts.user}`],
              ["rep", `Reps ${counts.rep}`],
            ] as const
          ).map(([key, label]) => (
            <button
              key={key}
              onClick={() => setRole(key)}
              className={`rounded-md px-3 py-1.5 font-medium transition ${
                role === key
                  ? "bg-brand-600 text-white"
                  : "text-body hover:bg-chip"
              }`}
            >
              {label}
            </button>
          ))}
        </div>
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Search email, device, location or IP…"
          className="w-full max-w-xs rounded-md border border-line bg-surface px-3 py-2 text-sm outline-none focus:border-brand-400"
        />
      </div>

      {err && (
        <p className="rounded-md border border-rose-200 bg-rose-50 px-4 py-2.5 text-sm text-rose-700">
          {err}
        </p>
      )}

      {/* Table */}
      <div className="overflow-x-auto rounded-xl border border-line bg-surface">
        <table className="w-full min-w-[820px] text-sm">
          <thead>
            <tr className="border-b border-line bg-muted text-left text-xs uppercase tracking-wide text-faint">
              <th className="px-4 py-2.5 font-medium">When</th>
              <th className="px-4 py-2.5 font-medium">User</th>
              <th className="px-4 py-2.5 font-medium">Role</th>
              <th className="px-4 py-2.5 font-medium">Device</th>
              <th className="px-4 py-2.5 font-medium">Location</th>
              <th className="px-4 py-2.5 font-medium">IP</th>
            </tr>
          </thead>
          <tbody className="[&>tr]:border-b [&>tr]:border-line [&>tr:last-child]:border-0">
            {loading ? (
              <tr>
                <td colSpan={6} className="px-4 py-8 text-center text-sm text-faint">
                  Loading history…
                </td>
              </tr>
            ) : filtered.length === 0 ? (
              <tr>
                <td colSpan={6} className="px-4 py-8 text-center text-sm text-faint">
                  {rows.length === 0
                    ? "No logins recorded yet. History starts filling as people sign in."
                    : "No entries match your filters."}
                </td>
              </tr>
            ) : (
              filtered.map((r) => (
                <tr key={r.id} className="hover:bg-muted/40">
                  <td className="whitespace-nowrap px-4 py-2.5">
                    <div className="font-medium text-strong">
                      {relativeTime(r.createdAt)}
                    </div>
                    <div className="text-[11px] text-faint">
                      {new Date(r.createdAt).toLocaleString()}
                    </div>
                  </td>
                  <td className="px-4 py-2.5 font-medium text-strong">{r.email}</td>
                  <td className="px-4 py-2.5">
                    <span
                      className={`rounded px-2 py-0.5 text-xs font-medium capitalize ${roleClass(
                        r.role
                      )}`}
                    >
                      {r.role}
                    </span>
                  </td>
                  <td className="px-4 py-2.5 text-body">{r.device || "—"}</td>
                  <td className="px-4 py-2.5 text-body">{r.location || "—"}</td>
                  <td className="px-4 py-2.5 font-mono text-xs text-faint">
                    {r.ip || "—"}
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      <p className="text-xs text-faint">
        Location is detected from the network address at sign-in. On a local
        network or VPN it may be approximate or unavailable.
      </p>
    </div>
  );
}
