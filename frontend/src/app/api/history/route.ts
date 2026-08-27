import { NextResponse } from "next/server";
import { cookies } from "next/headers";
import { SESSION_COOKIE, verifySession } from "@/lib/auth";
import {
  listHistory,
  clearHistory,
  enrichMissingLocations,
} from "@/lib/login-history";

// Login / access history — admin only.

async function isAdmin() {
  const token = (await cookies()).get(SESSION_COOKIE)?.value;
  return verifySession(token);
}

export async function GET() {
  if (!(await isAdmin()))
    return NextResponse.json({ error: "Admins only." }, { status: 403 });
  const rows = await listHistory();
  await enrichMissingLocations(rows); // best-effort; fills location when possible
  return NextResponse.json({ history: rows });
}

export async function DELETE() {
  if (!(await isAdmin()))
    return NextResponse.json({ error: "Admins only." }, { status: 403 });
  await clearHistory();
  return NextResponse.json({ ok: true });
}
