import { NextResponse } from "next/server";
import { pool } from "@/lib/db";
import { currentEmail } from "@/lib/auth";

// Lightweight list of app-user emails/names for share autocomplete. Any signed-in
// user can read it (email + name only — no secrets).
export async function GET() {
  if (!(await currentEmail()))
    return NextResponse.json({ error: "Not signed in." }, { status: 401 });
  try {
    const { rows } = await pool.query(
      "SELECT email, name FROM app_users ORDER BY email"
    );
    return NextResponse.json({ users: rows });
  } catch {
    return NextResponse.json({ users: [] });
  }
}
     