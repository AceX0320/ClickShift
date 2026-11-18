/**
 * ClickShift Shared ML Cloud Backend (Cloudflare Worker + SQLite Durable Object).
 *
 * Implements:
 *   GET  /health -> service health & configuration status
 *   GET  /state  -> pooled global LinUCB bandit & median latencies JSON
 *   PUT  /state  -> validates & merges one client's local bandit state atomically
 */

const DIM = 10;
const MAX_BODY_BYTES = 262_144; // 256 KiB cap on client PUT
const MAX_ARMS = 512;
const MAX_MEDIAN_PAIRS = 2048;
const MAX_SAMPLES_PER_ARM = 1_000_000;
const MAX_LATENCY_MS = 86_400_000; // 24 h
const MAX_CONTRIBUTORS = 400;
const RATE_WINDOW_MS = 10 * 60 * 1000; // 10 min
const RATE_LIMIT_MAX_PUTS = 30;
const ARM_NAME_RE = /^[a-z0-9._-]{1,20}>[a-z0-9._-]{1,20}\|[a-z0-9._-]{1,64}$/i;
const PAIR_KEY_RE = /^[a-z0-9._-]{1,20}>[a-z0-9._-]{1,20}$/i;
const CLIENT_ID_RE = /^[a-zA-Z0-9_-]{8,64}$/;

function jsonResponse(data, status = 200, extraHeaders = {}) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "X-Content-Type-Options": "nosniff",
      ...extraHeaders,
    },
  });
}

function isSymmetricPSD(A) {
  if (!Array.isArray(A) || A.length !== DIM) return false;
  for (let i = 0; i < DIM; i++) {
    const row = A[i];
    if (!Array.isArray(row) || row.length !== DIM) return false;
    for (let j = 0; j < DIM; j++) {
      const v = row[j];
      if (typeof v !== "number" || !Number.isFinite(v) || Math.abs(v) > 1e9) return false;
      if (j < i && Math.abs(v - A[j][i]) > 1e-4) return false;
    }
    if (A[i][i] <= 0) return false;
  }
  // Cholesky check (L * L^T = A) verifies strict positive-definiteness
  const L = Array.from({ length: DIM }, () => new Float64Array(DIM));
  for (let i = 0; i < DIM; i++) {
    for (let j = 0; j <= i; j++) {
      let sum = A[i][j];
      for (let k = 0; k < j; k++) sum -= L[i][k] * L[j][k];
      if (i === j) {
        if (sum <= 1e-7) return false;
        L[i][j] = Math.sqrt(sum);
      } else {
        L[i][j] = sum / L[j][j];
      }
    }
  }
  return true;
}

export function validateClientState(raw, fallbackClientId = "anon_default") {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  if (raw.version !== undefined && raw.version !== 1) return null;

  const clientId =
    typeof raw.client_id === "string" && CLIENT_ID_RE.test(raw.client_id)
      ? raw.client_id
      : CLIENT_ID_RE.test(fallbackClientId)
        ? fallbackClientId
        : "anon_default";

  const clockRaw = Number(raw.clock ?? 0);
  const clock = Number.isFinite(clockRaw) ? Math.max(0, Math.min(1e9, Math.trunc(clockRaw))) : 0;

  const cleanBandit = {};
  const rawBandit = raw.bandit;
  if (!rawBandit || typeof rawBandit !== "object" || Array.isArray(rawBandit)) return null;

  let armCount = 0;
  for (const [arm, entry] of Object.entries(rawBandit)) {
    if (armCount >= MAX_ARMS) break;
    if (typeof arm !== "string" || arm.length > 128 || !ARM_NAME_RE.test(arm)) continue;
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) continue;

    const n = Number(entry.n);
    if (!Number.isInteger(n) || n < 0 || n > MAX_SAMPLES_PER_ARM) continue;
    if (!isSymmetricPSD(entry.A)) continue;

    const b = entry.b;
    if (!Array.isArray(b) || b.length !== DIM) continue;
    const bound = 2.25 * Math.max(n, 1);
    const cleanB = [];
    let bOk = true;
    for (let i = 0; i < DIM; i++) {
      const val = b[i];
      if (typeof val !== "number" || !Number.isFinite(val)) {
        bOk = false;
        break;
      }
      cleanB.push(Math.max(-bound, Math.min(bound, val)));
    }
    if (!bOk) continue;

    cleanBandit[arm] = {
      A: entry.A.map((row) => row.map((v) => Number(v))),
      b: cleanB,
      n,
    };
    armCount++;
  }

  const cleanMedians = {};
  const rawMedians = raw.median_ms;
  if (rawMedians && typeof rawMedians === "object" && !Array.isArray(rawMedians)) {
    let pairCount = 0;
    for (const [pair, msVal] of Object.entries(rawMedians)) {
      if (pairCount >= MAX_MEDIAN_PAIRS) break;
      if (typeof pair !== "string" || !PAIR_KEY_RE.test(pair)) continue;
      const ms = Number(msVal);
      if (!Number.isFinite(ms) || ms <= 0 || ms > MAX_LATENCY_MS) continue;
      cleanMedians[pair] = ms;
      pairCount++;
    }
  }

  if (armCount === 0 && Object.keys(cleanMedians).length === 0) return null;
  return {
    clientId,
    state: { version: 1, clock, bandit: cleanBandit, median_ms: cleanMedians },
  };
}

export function mergeStates(states, prevClock = 0) {
  const accum = new Map();
  const medians = {};
  let maxClock = Math.max(0, Number(prevClock) || 0);

  for (const st of states) {
    if (!st || typeof st !== "object") continue;
    if (Number.isFinite(Number(st.clock))) {
      maxClock = Math.max(maxClock, Math.min(1e9, Math.trunc(Number(st.clock))));
    }
    const bandit = st.bandit || {};
    for (const [arm, armData] of Object.entries(bandit)) {
      if (!armData || typeof armData !== "object") continue;
      const n = Number(armData.n || 0);
      if (!Number.isFinite(n) || n < 0 || n > MAX_SAMPLES_PER_ARM) continue;
      let acc = accum.get(arm);
      if (!acc) {
        acc = {
          ASum: Array.from({ length: DIM }, () => new Float64Array(DIM)),
          bSum: new Float64Array(DIM),
          nA: 0,
          n: 0,
        };
        accum.set(arm, acc);
      }
      acc.n = Math.min(MAX_SAMPLES_PER_ARM, acc.n + Math.trunc(n));
      if (n > 0 && Array.isArray(armData.A) && Array.isArray(armData.b)) {
        for (let i = 0; i < DIM; i++) {
          acc.bSum[i] += Number(armData.b[i]);
          for (let j = 0; j < DIM; j++) {
            acc.ASum[i][j] += Number(armData.A[i][j]);
          }
        }
        acc.nA++;
      }
    }
    const pairs = st.median_ms || {};
    for (const [pair, msVal] of Object.entries(pairs)) {
      const ms = Number(msVal);
      if (!Number.isFinite(ms) || ms <= 0 || ms > MAX_LATENCY_MS) continue;
      if (medians[pair] === undefined || ms < medians[pair]) medians[pair] = ms;
    }
  }

  const outBandit = {};
  for (const [arm, acc] of accum.entries()) {
    if (acc.nA === 0) continue;
    const denom = Math.max(1, acc.nA);
    outBandit[arm] = {
      A: acc.ASum.map((row) => Array.from(row, (v) => v / denom)),
      b: Array.from(acc.bSum, (v) => v / denom),
      n: acc.n,
    };
  }

  return {
    version: 1,
    clock: maxClock + 1,
    updated: Date.now() / 1000,
    bandit: outBandit,
    median_ms: medians,
  };
}

export class CloudPoolDO {
  constructor(ctx, env) {
    this.ctx = ctx;
    this.env = env;
    this.rateMap = new Map();
    this.schemaReady = false;
  }

  _ensureSchema() {
    if (this.schemaReady) return;
    this.ctx.storage.sql.exec(`
      CREATE TABLE IF NOT EXISTS contributors (
        client_id TEXT PRIMARY KEY,
        updated_at INTEGER NOT NULL,
        payload TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS pool_state (
        id INTEGER PRIMARY KEY CHECK (id = 1),
        clock INTEGER NOT NULL,
        merged_json TEXT NOT NULL
      );
    `);
    this.schemaReady = true;
  }

  _rateLimited(ip) {
    const key = ip || "unknown";
    const now = Date.now();
    const cur = this.rateMap.get(key);
    if (!cur || now - cur.start >= RATE_WINDOW_MS) {
      this.rateMap.set(key, { start: now, count: 1 });
      if (this.rateMap.size > 5000) this.rateMap.clear();
      return false;
    }
    cur.count += 1;
    return cur.count > RATE_LIMIT_MAX_PUTS;
  }

  _currentPool() {
    this._ensureSchema();
    const rows = this.ctx.storage.sql
      .exec("SELECT clock, merged_json FROM pool_state WHERE id = 1")
      .toArray();
    if (rows.length > 0) {
      return { clock: Number(rows[0].clock), mergedJson: String(rows[0].merged_json) };
    }
    const empty = JSON.stringify({
      version: 1,
      clock: 0,
      updated: Date.now() / 1000,
      bandit: {},
      median_ms: {},
    });
    return { clock: 0, mergedJson: empty };
  }

  async fetch(request) {
    const url = new URL(request.url);
    if (request.method === "GET" && url.pathname === "/state") {
      const { mergedJson } = this._currentPool();
      return new Response(mergedJson, {
        status: 200,
        headers: {
          "Content-Type": "application/json; charset=utf-8",
          "Cache-Control": "public, max-age=30",
          "X-Content-Type-Options": "nosniff",
        },
      });
    }

    if (request.method === "PUT" && url.pathname === "/state") {
      const ip = request.headers.get("CF-Connecting-IP") || "unknown";
      if (this._rateLimited(ip)) {
        return jsonResponse({ error: "rate limit exceeded" }, 429);
      }

      const buf = await request.arrayBuffer();
      if (buf.byteLength > MAX_BODY_BYTES) {
        return jsonResponse({ error: "payload too large" }, 413);
      }

      let rawJson;
      try {
        rawJson = JSON.parse(new TextDecoder().decode(buf));
      } catch {
        return jsonResponse({ error: "invalid json" }, 400);
      }

      const fallbackId = `ip_${ip.replace(/[^a-zA-Z0-9]/g, "_").slice(0, 24)}`;
      const validated = validateClientState(rawJson, fallbackId);
      if (!validated) {
        return jsonResponse({ error: "invalid or empty bandit payload" }, 400);
      }

      this._ensureSchema();
      const now = Date.now();
      this.ctx.storage.sql.exec(
        `INSERT INTO contributors (client_id, updated_at, payload)
         VALUES (?, ?, ?)
         ON CONFLICT(client_id) DO UPDATE SET
           updated_at = excluded.updated_at,
           payload = excluded.payload`,
        validated.clientId,
        now,
        JSON.stringify(validated.state)
      );

      this.ctx.storage.sql.exec(
        `DELETE FROM contributors
         WHERE client_id NOT IN (
           SELECT client_id FROM contributors ORDER BY updated_at DESC LIMIT ?
         )`,
        MAX_CONTRIBUTORS
      );

      const rows = this.ctx.storage.sql
        .exec("SELECT payload FROM contributors ORDER BY updated_at DESC")
        .toArray();
      const contribs = [];
      for (const r of rows) {
        try {
          contribs.push(JSON.parse(String(r.payload)));
        } catch {
          // ignore corrupt historical row
        }
      }

      const { clock: prevClock } = this._currentPool();
      const merged = mergeStates(contribs, prevClock);
      const mergedJson = JSON.stringify(merged);
      this.ctx.storage.sql.exec(
        `INSERT INTO pool_state (id, clock, merged_json)
         VALUES (1, ?, ?)
         ON CONFLICT(id) DO UPDATE SET
           clock = excluded.clock,
           merged_json = excluded.merged_json`,
        merged.clock,
        mergedJson
      );

      return jsonResponse({
        ok: true,
        clock: merged.clock,
        arms: Object.keys(merged.bandit).length,
        contributors: contribs.length,
      });
    }

    return jsonResponse({ error: "not found" }, 404);
  }
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (request.method === "GET" && (url.pathname === "/" || url.pathname === "/health")) {
      return jsonResponse({ ok: true, service: "clickshift-cloud", version: 1 });
    }
    if (url.pathname !== "/state" || (request.method !== "GET" && request.method !== "PUT")) {
      return jsonResponse({ error: "not found" }, 404);
    }
    if (request.method === "PUT") {
      const clen = Number(request.headers.get("content-length") || 0);
      if (Number.isFinite(clen) && clen > MAX_BODY_BYTES) {
        return jsonResponse({ error: "payload too large" }, 413);
      }
    }
    const stub = env.CLOUD_POOL.get(env.CLOUD_POOL.idFromName("global"));
    return stub.fetch(request);
  },
};
