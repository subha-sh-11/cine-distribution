import { NextResponse } from "next/server";
import { pool } from "@/lib/db";
import { currentEmail } from "@/lib/auth";

// Manage who a movie is shared with (email + role: editor|viewer).
// Only an editor (admin or owner) may view/change the share list.

async function ensure() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS movie_shares (
      movie_id   TEXT NOT NULL,
      email      TEXT NOT NULL,
      role       TEXT NOT NULL DEFAULT 'editor',
      created_at BIGINT NOT NULL,
      PRIMARY KEY (movie_id, email)
    )
  `);
  await pool
    .query("ALTER TABLE movies ADD COLUMN IF NOT EXISTS owner_email TEXT")
    .catch(() => {});
  await pool
    .query("ALTER TABLE movie_shares ADD COLUMN IF NOT EXISTS shared_by TEXT")
    .catch(() => {});
}

const whoami = currentEmail; // admin or user session

async function isEditor(movieId: string, email: string): Promise<boolean> {
  const admin = (process.env.ADMIN_EMAIL || "admin@svf.in").toLowerCase();
  if (email.toLowerCase() === admin) return true;
  const owner = await pool.query("SELECT owner_email FROM movies WHERE id = $1", [
    movieId,
  ]);
  if (owner.rows[0]?.owner_email?.toLowerCase() === email.toLowerCase())
    return true;
  const shr = await pool.query(
    "SELECT role FROM movie_shares WHERE movie_id=$1 AND LOWER(email)=LOWER($2)",
    [movieId, email]
  );
  return shr.rows[0]?.role === "editor";
}

export async function GET(
  _req: Request,
  { params }: { params: Promise<{ id: string }> }
) {
  const email = await whoami();
  if (!email) return NextResponse.json({ error: "Not signed in." }, { status: 401 });
  await ensure();
  const { id } = await params;
  if (!(await isEditor(id, email)))
    return NextResponse.json({ error: "No access." }, { status: 403 });
  const owner = await pool.query("SELECT owner_email FROM movies WHERE id=$1", [id]);
  const { rows } = await pool.query(
    "SELECT email, role FROM movie_shares WHERE movie_id=$1 ORDER BY created_at",
    [id]
  );
  return NextResponse.json({ owner: owner.rows[0]?.owner_email ?? null, shares: rows });
}

export async function POST(
  req: Request,
  { params }: { params: Promise<{ id: string }> }
) {
  const email = await whoami();
  if (!email) return NextResponse.json({ error: "Not signed in." }, { status: 401 });
  await ensure();
  const { id } = await params;
  if (!(await isEditor(id, email)))
    return NextResponse.json({ error: "No access." }, { status: 403 });
  const b = (await req.json().catch(() => ({}))) as { email?: string; role?: string };
  const target = String(b.email ?? "").trim().toLowerCase();
  const role = b.role === "viewer" ? "viewer" : "editor";
  if (!target || !target.includes("@"))
    return NextResponse.json({ error: "Valid email required." }, { status: 400 });
  await pool.query(
    `INSERT INTO movie_shares (movie_id, email, role, created_at, shared_by)
       VALUES ($1,$2,$3,$4,$5)
     ON CONFLICT (movie_id, email)
       DO UPDATE SET role = EXCLUDED.role, shared_by = EXCLUDED.shared_by`,
    [id, target, role, Date.now(), email]
  );
  return NextResponse.json({ ok: true });
}

export async function DELETE(
  req: Request,
  { params }: { params: Promise<{ id: string }> }
) {
  const email = await whoami();
  if (!email) return NextResponse.json({ error: "Not signed in." }, { status: 401 });
  await ensure();
  const { id } = await params;
  if (!(await isEditor(id, email)))
    return NextResponse.json({ error: "No access." }, { status: 403 });
  const target = new URL(req.url).searchParams.get("email");
  if (!target)
    return NextResponse.json({ error: "email required." }, { status: 400 });
  await pool.query(
    "DELETE FROM movie_shares WHERE movie_id=$1 AND LOWER(email)=LOWER($2)",
    [id, target]
  );
  return NextResponse.json({ ok: true });
}
