import { NextResponse } from "next/server";
import { pool } from "@/lib/db";
import {
  ADMIN_EMAIL,
  checkCredentials,
  checkUserCredentials,
  signSession,
  SESSION_COOKIE,
  SESSION_MAX_AGE,
  USER_EMAIL,
  USER_SESSION_COOKIE,
} from "@/lib/auth";
import {
  REP_SESSION_COOKIE,
  REP_SESSION_MAX_AGE,
  signRepSession,
  verifyPassword,
} from "@/lib/rep-auth";
import { recordLogin } from "@/lib/login-history";

const cookieOpts = {
  httpOnly: true as const,
  sameSite: "lax" as const,
  secure: process.env.NODE_ENV === "production",
  path: "/",
};

export async function POST(req: Request) {
  let email = "";
  let password = "";
  try {
    const b = await req.json();
    email = String(b.email ?? "");
    password = String(b.password ?? "");
  } catch {
    return NextResponse.json({ error: "Bad request" }, { status: 400 });
  }

  // 1) Admin
  if (checkCredentials(email, password)) {
    const res = NextResponse.json({ role: "admin", redirect: "/admin/theatres" });
    res.cookies.set(SESSION_COOKIE, signSession(ADMIN_EMAIL), {
      ...cookieOpts,
      maxAge: SESSION_MAX_AGE,
    });
    await recordLogin(req, ADMIN_EMAIL, "admin");
    return res;
  }

  // 2) User (movies viewer)
  if (checkUserCredentials(email, password)) {
    const res = NextResponse.json({ role: "user", redirect: "/movies" });
    res.cookies.set(USER_SESSION_COOKIE, signSession(USER_EMAIL), {
      ...cookieOpts,
      maxAge: SESSION_MAX_AGE,
    });
    await recordLogin(req, USER_EMAIL, "user");
    return res;
  }

  // 3) Representative (validated against the DB)
  const em = email.trim().toLowerCase();
  const { rows } = await pool.query(
    `SELECT * FROM reps WHERE lower(email) = $1 AND status = 'approved' LIMIT 1`,
    [em]
  );
  const rep = rows[0];
  if (rep && verifyPassword(password, rep.password_hash)) {
    const res = NextResponse.json({ role: "rep", redirect: "/rep" });
    res.cookies.set(REP_SESSION_COOKIE, signRepSession(rep.id), {
      ...cookieOpts,
      maxAge: REP_SESSION_MAX_AGE,
    });
    await recordLogin(req, rep.email ?? em, "rep");
    return res;
  }

  // 4) App user (admin-created account) — logs in with their own email so
  //    owned + shared movies show up in their profile.
  try {
    const u = await pool.query(
      `SELECT email, password_hash FROM app_users WHERE lower(email) = $1 LIMIT 1`,
      [em]
    );
    const appUser = u.rows[0];
    if (appUser && verifyPassword(password, appUser.password_hash)) {
      const res = NextResponse.json({ role: "user", redirect: "/movies" });
      res.cookies.set(USER_SESSION_COOKIE, signSession(appUser.email), {
        ...cookieOpts,
        maxAge: SESSION_MAX_AGE,
      });
      await recordLogin(req, appUser.email, "user");
      return res;
    }
  } catch {
    /* app_users table may not exist yet — ignore */
  }

  return NextResponse.json({ error: "Invalid email or password" }, { status: 401 });
}
