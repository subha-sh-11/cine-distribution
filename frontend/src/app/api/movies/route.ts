import { NextResponse } from "next/server";
import crypto from "crypto";
import { pool } from "@/lib/db";
import { currentEmail } from "@/lib/auth";

const ADMIN = (process.env.ADMIN_EMAIL || "admin@svf.in").toLowerCase();

// Movies live in the shared Postgres DB (not per-browser localStorage) so every
// admin user sees the same list — someone adding a movie is visible to all.

let ensured = false;
async function ensureTable() {
  if (ensured) return;
  await pool.query(`
    CREATE TABLE IF NOT EXISTS movies (
      id          TEXT PRIMARY KEY,
      name        TEXT NOT NULL,
      release     TEXT DEFAULT '',
      created_at  BIGINT NOT NULL,
      owner_email TEXT
    )
  `);
  await pool
    .query("ALTER TABLE movies ADD COLUMN IF NOT EXISTS owner_email TEXT")
    .catch(() => {});
  ensured = true;
}

type Row = {
  id: string;
  name: string;
  release: string | null;
  created_at: string;
  owned?: boolean;
  owner_email?: string | null;
  share_role?: string | null;
  shared_by?: string | null;
};
const toMovie = (r: Row) => {
  // "Shared by" should be whoever actually shared the movie (recorded per share).
  // Fall back to the uploader only for legacy shares that predate that tracking.
  // The internal admin/system account is never surfaced as a sharer — end users
  // should only ever see a real person's share, so admin is collapsed to null
  // (the UI then shows a neutral "another user").
  const rawSharer = r.shared_by ?? r.owner_email ?? null;
  const sharedBy =
    rawSharer && rawSharer.toLowerCase() !== ADMIN ? rawSharer : null;
  return {
    id: r.id,
    name: r.name,
    release: r.release ?? "",
    createdAt: Number(r.created_at),
    owned: r.owned !== false, // admin/owner → true; shared-with-me → false
    ownerEmail: r.owner_email ?? null,
    sharedBy,
    role: r.share_role ?? (r.owned === false ? "viewer" : "editor"),
  };
};

export async function GET() {
  const email = await currentEmail();
  if (!email)
    return NextResponse.json({ error: "Not signed in." }, { status: 401 });
  await ensureTable();
  // Admin sees every movie; a user sees only the ones they own or are shared.
  if (email.toLowerCase() === ADMIN) {
    const { rows } = await pool.query<Row>(
      "SELECT id, name, release, created_at, TRUE AS owned, owner_email FROM movies ORDER BY created_at DESC"
    );
    return NextResponse.json({ movies: rows.map(toMovie) });
  }
  await pool
    .query(
      `CREATE TABLE IF NOT EXISTS movie_shares (
         movie_id TEXT NOT NULL, email TEXT NOT NULL,
         role TEXT NOT NULL DEFAULT 'editor', created_at BIGINT NOT NULL,
         PRIMARY KEY (movie_id, email))`
    )
    .catch(() => {});
  await pool
    .query("ALTER TABLE movie_shares ADD COLUMN IF NOT EXISTS shared_by TEXT")
    .catch(() => {});
  // Join only THIS user's share row so we can tell owned vs shared-with-me and
  // surface the role + who shared it. A user sees a movie if they own it OR it's
  // shared to them.
  const { rows } = await pool.query<Row>(
    `SELECT DISTINCT m.id, m.name, m.release, m.created_at,
            (lower(m.owner_email) = lower($1)) AS owned,
            m.owner_email,
            s.role AS share_role,
            s.shared_by
       FROM movies m
       LEFT JOIN movie_shares s
         ON s.movie_id = m.id AND lower(s.email) = lower($1)
      WHERE lower(m.owner_email) = lower($1) OR s.email IS NOT NULL
      ORDER BY m.created_at DESC`,
    [email]
  );
  return NextResponse.json({ movies: rows.map(toMovie) });
}

export async function POST(req: Request) {
  const owner = await currentEmail();
  if (!owner)
    return NextResponse.json({ error: "Not signed in." }, { status: 401 });
  await ensureTable();
  const b = (await req.json().catch(() => ({}))) as Record<string, unknown>;
  const name = String(b.name ?? "").trim();
  if (!name) return NextResponse.json({ error: "Name required." }, { status: 400 });
  const release = String(b.release ?? "").trim();
  const id = "mov-" + crypto.randomBytes(5).toString("hex");
  const { rows } = await pool.query<Row>(
    "INSERT INTO movies (id, name, release, created_at, owner_email) VALUES ($1,$2,$3,$4,$5) RETURNING id, name, release, created_at",
    [id, name, release, Date.now(), owner]
  );
  return NextResponse.json({ movie: toMovie(rows[0]) });
}
