/**
 * ClickShift Community Cloud ML Sync Worker (Cloudflare Workers + D1 Free Tier).
 *
 * Free tier capacity:
 *   - Workers: 100,000 requests/day
 *   - D1 SQLite: 100,000 rows written/day, 5,000,000 rows read/day
 *   (Unlike Cloudflare Workers KV, which caps free writes at 1,000/day.)
 *
 * Setup (once):
 *   1. Create a D1 database: `wrangler d1 create clickshift-ml`
 *   2. Run schema:
 *      CREATE TABLE IF NOT EXISTS contrib (
 *        client_id TEXT PRIMARY KEY,
 *        payload   TEXT NOT NULL,
 *        updated   INTEGER NOT NULL
 *      );
 *      CREATE TABLE IF NOT EXISTS merged (
 *        id      INTEGER PRIMARY KEY CHECK (id = 1),
 *        payload TEXT NOT NULL,
 *        updated INTEGER NOT NULL
 *      );
 *   3. Configure D1 binding `DB` in wrangler.toml and deploy `worker.js`.
 *   4. Users then point ClickShift at `https://<your-worker>.workers.dev` via:
 *      `ClickShift.exe config cloud_url=https://<your-worker>.workers.dev cloud_sync=on`
 */

const MAX_BODY_BYTES = 128 * 1024;
const MAX_ARMS = 512;
const MAX_SAMPLES_PER_ARM = 1_000_000;
const MAX_LATENCY_MS = 86_400_000;

function validStateShape(obj) {
  if (!obj || typeof obj !== "object") return false;
  if (typeof obj.bandit !== "object" || obj.bandit === null) return false;
  const arms = Object.entries(obj.bandit);
  if (arms.length > MAX_ARMS) return false;
  for (const [, arm] of arms) {
    if (!arm || typeof arm !== "object") return false;
    const n = Number(arm.n || 0);
    if (!Number.isFinite(n) || n < 0 || n > MAX_SAMPLES_PER_ARM) return false;
    if (!Array.isArray(arm.A) || !Array.isArray(arm.b)) return false;
  }
  if (obj.median_ms && typeof obj.median_ms === "object") {
    for (const ms of Object.values(obj.median_ms)) {
      const v = Number(ms);
      if (!Number.isFinite(v) || v <= 0 || v > MAX_LATENCY_MS) return false;
    }
  }
  return true;
}

function mergeStates(rows) {
  let maxClock = 0;
  const bandit = {};
  const median_ms = {};
  const counts = {};

  for (const row of rows) {
    let st;
    try {
      st = JSON.parse(row.payload);
    } catch {
      continue;
    }
    if (!validStateShape(st)) continue;
    maxClock = Math.max(maxClock, Number(st.clock || 0));
    for (const [armName, arm] of Object.entries(st.bandit || {})) {
      const n = Math.min(Number(arm.n || 0), MAX_SAMPLES_PER_ARM);
      if (!n) continue;
      if (!bandit[armName]) {
        bandit[armName] = { A: arm.A, b: arm.b, n };
        counts[armName] = 1;
      } else {
        const cur = bandit[armName];
        const k = counts[armName] + 1;
        for (let i = 0; i < Math.min(cur.b.length, arm.b.length); i++) {
          cur.b[i] += (Number(arm.b[i]) - cur.b[i]) / k;
        }
        for (let i = 0; i < Math.min(cur.A.length, arm.A.length); i++) {
          for (let j = 0; j < Math.min(cur.A[i].length, arm.A[i].length); j++) {
            cur.A[i][j] += (Number(arm.A[i][j]) - cur.A[i][j]) / k;
          }
        }
        cur.n += n;
        counts[armName] = k;
      }
    }
    for (const [pair, ms] of Object.entries(st.median_ms || {})) {
      const v = Number(ms);
      if (!Number.isFinite(v) || v <= 0 || v > MAX_LATENCY_MS) continue;
      median_ms[pair] = median_ms[pair] ? Math.min(median_ms[pair], v) : v;
    }
  }
  return { version: 1, clock: maxClock + 1, updated: Math.floor(Date.now() / 1000), bandit, median_ms };
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (request.method === "GET" && url.pathname === "/state") {
      const row = await env.DB.prepare("SELECT payload FROM merged WHERE id = 1").first();
      const body = row ? row.payload : JSON.stringify({ version: 1, clock: 0, bandit: {}, median_ms: {} });
      return new Response(body, { headers: { "Content-Type": "application/json", "Cache-Control": "public, max-age=60" } });
    }

    if (request.method === "PUT" && url.pathname === "/state") {
      const text = await request.text();
      if (text.length > MAX_BODY_BYTES) {
        return new Response("Payload too large", { status: 413 });
      }
      let parsed;
      try {
        parsed = JSON.parse(text);
      } catch {
        return new Response("Invalid JSON", { status: 400 });
      }
      if (!validStateShape(parsed)) {
        return new Response("Invalid state schema", { status: 400 });
      }
      const clientHash = await crypto.subtle.digest(
        "SHA-256",
        new TextEncoder().encode(request.headers.get("CF-Connecting-IP") || "anon")
      );
      const clientId = Array.from(new Uint8Array(clientHash)).slice(0, 12).map((b) => b.toString(16).padStart(2, "0")).join("");
      const now = Math.floor(Date.now() / 1000);

      await env.DB.prepare(
        "INSERT INTO contrib (client_id, payload, updated) VALUES (?1, ?2, ?3) ON CONFLICT(client_id) DO UPDATE SET payload = excluded.payload, updated = excluded.updated"
      ).bind(clientId, text, now).run();

      const allRows = await env.DB.prepare("SELECT payload FROM contrib ORDER BY updated DESC LIMIT 500").all();
      const merged = mergeStates(allRows.results || []);
      await env.DB.prepare(
        "INSERT INTO merged (id, payload, updated) VALUES (1, ?1, ?2) ON CONFLICT(id) DO UPDATE SET payload = excluded.payload, updated = excluded.updated"
      ).bind(JSON.stringify(merged), now).run();

      return new Response(JSON.stringify({ ok: true, clock: merged.clock }), {
        headers: { "Content-Type": "application/json" },
      });
    }

    return new Response("Not Found", { status: 404 });
  },
};
