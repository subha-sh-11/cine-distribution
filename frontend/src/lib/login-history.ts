import { pool } from "@/lib/db";

// Records every successful login (admin / user / rep) with the device and
// location it came from, so the admin can review access history.

let ensured = false;
async function ensure() {
  if (ensured) return;
  await pool.query(`
    CREATE TABLE IF NOT EXISTS login_history (
      id          BIGSERIAL PRIMARY KEY,
      email       TEXT NOT NULL,
      role        TEXT NOT NULL,
      ip          TEXT DEFAULT '',
      device      TEXT DEFAULT '',
      user_agent  TEXT DEFAULT '',
      location    TEXT DEFAULT '',
      created_at  BIGINT NOT NULL
    )
  `);
  ensured = true;
}

/** Best-effort client IP from the usual proxy headers (Vercel, CF, nginx). */
export function getClientIp(req: Request): string {
  const xff = req.headers.get("x-forwarded-for");
  if (xff) return xff.split(",")[0].trim();
  return (
    req.headers.get("x-real-ip") ||
    req.headers.get("cf-connecting-ip") ||
    ""
  );
}

/** Turn a User-Agent string into a friendly "Chrome on Windows · Desktop". */
export function parseUserAgent(ua: string): string {
  if (!ua) return "Unknown device";

  let os = "Unknown OS";
  if (/windows/i.test(ua)) os = "Windows";
  else if (/iphone|ipod/i.test(ua)) os = "iOS";
  else if (/ipad/i.test(ua)) os = "iPadOS";
  else if (/android/i.test(ua)) os = "Android";
  else if (/mac os x|macintosh/i.test(ua)) os = "macOS";
  else if (/cros/i.test(ua)) os = "ChromeOS";
  else if (/linux/i.test(ua)) os = "Linux";

  let browser = "Unknown browser";
  if (/edg\//i.test(ua)) browser = "Edge";
  else if (/opr\/|opera/i.test(ua)) browser = "Opera";
  else if (/samsungbrowser/i.test(ua)) browser = "Samsung Internet";
  else if (/firefox\/|fxios/i.test(ua)) browser = "Firefox";
  else if (/chrome\/|crios/i.test(ua)) browser = "Chrome";
  else if (/safari\//i.test(ua)) browser = "Safari";

  let type = "Desktop";
  if (/ipad|tablet/i.test(ua)) type = "Tablet";
  else if (/mobi|iphone|android/i.test(ua)) type = "Mobile";

  return `${browser} on ${os} · ${type}`;
}

/** Location from Vercel's edge geo headers (present automatically in prod). */
export function getGeoFromHeaders(req: Request): string {
  const dec = (s: string | null) => {
    if (!s) return "";
    try {
      return decodeURIComponent(s);
    } catch {
      return s;
    }
  };
  const city = dec(req.headers.get("x-vercel-ip-city"));
  const region = dec(req.headers.get("x-vercel-ip-country-region"));
  const country = dec(req.headers.get("x-vercel-ip-country"));
  return [city, region, country].filter(Boolean).join(", ");
}

/** Record a successful login. Never throws — history must not block sign-in. */
export async function recordLogin(req: Request, email: string, role: string) {
  try {
    await ensure();
    const ip = getClientIp(req);
    const ua = req.headers.get("user-agent") || "";
    const device = parseUserAgent(ua);
    const location = getGeoFromHeaders(req);
    await pool.query(
      `INSERT INTO login_history (email, role, ip, device, user_agent, location, created_at)
         VALUES ($1,$2,$3,$4,$5,$6,$7)`,
      [email, role, ip, device, ua, location, Date.now()]
    );
  } catch {
    /* swallow — a failed history write must never fail the login */
  }
}

export type HistoryRow = {
  id: string;
  email: string;
  role: string;
  ip: string;
  device: string;
  location: string;
  createdAt: number;
};

export async function listHistory(limit = 500): Promise<HistoryRow[]> {
  await ensure();
  const { rows } = await pool.query(
    `SELECT id, email, role, ip, device, location, created_at
       FROM login_history ORDER BY created_at DESC LIMIT $1`,
    [limit]
  );
  return rows.map((r) => ({
    id: String(r.id),
    email: r.email,
    role: r.role,
    ip: r.ip ?? "",
    device: r.device ?? "",
    location: r.location ?? "",
    createdAt: Number(r.created_at),
  }));
}

export async function clearHistory() {
  await ensure();
  await pool.query("DELETE FROM login_history");
}

/** True for loopback / private LAN IPs we can't geolocate. */
function isPrivateIp(ip: string): boolean {
  if (!ip) return true;
  const v = ip.replace(/^::ffff:/i, "");
  return (
    v === "::1" ||
    v === "127.0.0.1" ||
    /^127\./.test(v) ||
    /^10\./.test(v) ||
    /^192\.168\./.test(v) ||
    /^172\.(1[6-9]|2\d|3[01])\./.test(v) ||
    /^169\.254\./.test(v) ||
    /^fc00:|^fd00:|^fe80:/i.test(v)
  );
}

/**
 * Fill in `location` for rows that don't have one yet (e.g. deployments not on
 * Vercel, where geo headers are absent) using a free IP lookup, and cache the
 * result back into the table. Best-effort and time-boxed; never throws.
 */
export async function enrichMissingLocations(rows: HistoryRow[]): Promise<void> {
  const ips = Array.from(
    new Set(
      rows
        .filter((r) => !r.location && !isPrivateIp(r.ip))
        .map((r) => r.ip)
    )
  ).slice(0, 20);
  if (ips.length === 0) return;

  const results = await Promise.allSettled(
    ips.map(async (ip) => {
      const ctrl = new AbortController();
      const t = setTimeout(() => ctrl.abort(), 2500);
      try {
        const res = await fetch(
          `http://ip-api.com/json/${encodeURIComponent(
            ip
          )}?fields=status,country,regionName,city`,
          { signal: ctrl.signal }
        );
        const data = (await res.json()) as {
          status?: string;
          country?: string;
          regionName?: string;
          city?: string;
        };
        if (data.status !== "success") return { ip, loc: "" };
        const loc = [data.city, data.regionName, data.country]
          .filter(Boolean)
          .join(", ");
        return { ip, loc };
      } catch {
        return { ip, loc: "" };
      } finally {
        clearTimeout(t);
      }
    })
  );

  const map = new Map<string, string>();
  for (const r of results) {
    if (r.status === "fulfilled" && r.value.loc) map.set(r.value.ip, r.value.loc);
  }
  if (map.size === 0) return;

  // Cache back and reflect in the returned rows.
  await Promise.allSettled(
    Array.from(map.entries()).map(([ip, loc]) =>
      pool.query(
        `UPDATE login_history SET location=$1
           WHERE ip=$2 AND (location IS NULL OR location='')`,
        [loc, ip]
      )
    )
  );
  for (const row of rows) {
    if (!row.location && map.has(row.ip)) row.location = map.get(row.ip)!;
  }
}
