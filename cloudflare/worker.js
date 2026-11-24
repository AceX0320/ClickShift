/**
 * ClickShift Community Cloud ML Sync Worker (Cloudflare Workers + D1 Free Tier).
 *
 * Includes Download Client ID Protection:
 *   - Only client_ids registered in the `downloads` table (from official
 *     ClickShift downloads / attested binary registrations) are permitted to
 *     contribute ML data via `PUT /state`.
 *   - Any `PUT /state` whose `client_id` does not match a registered, non-revoked
 *     download ID in `downloads` is rejected with HTTP 403 Forbidden.
 *
 * Setup (once):
 *   1. Create a D1 database: `wrangler d1 create clickshift-ml`
 *   2. Configure D1 binding `DB` in wrangler.toml (and optional `ADMIN_TOKEN` secret)
 *      and deploy `worker.js` (tables are auto-created on first request).
 */

const MAX_BODY_BYTES = 128 * 1024;
const MAX_ARMS = 512;
const MAX_SAMPLES_PER_ARM = 1_000_000;
const MAX_LATENCY_MS = 86_400_000;
const LATEST_EXE_DOWNLOAD_URL = "https://github.com/AceX0320/ClickShift/releases/latest/download/ClickShift.exe";
const CLIENT_ID_RE = /^[0-9a-f]{12,64}$/;
const SEED_DOWNLOAD_CLIENT_IDS = ["a66c09d196954a1c"];
const DEFAULT_ATTEST_CONTEXT =
  "ClickShift-Download-Attestation-v1:RUNTMSAAAAA51GhTNj8aBWoIBxWojEWaMNYmLZBR+9rwzScwqaFE/nMIqrMDoTuPrnN/VFzAM8Biv7QwpRHdLnn6z9Op3OhU";

let schemaInitialized = false;

async function ensureSchema(db) {
  if (schemaInitialized || !db) return;
  await db
    .prepare(
      "CREATE TABLE IF NOT EXISTS downloads (client_id TEXT PRIMARY KEY, download_proof TEXT NOT NULL, app_version TEXT NOT NULL, exe_sha256 TEXT NOT NULL, registered_at INTEGER NOT NULL, last_seen INTEGER NOT NULL, revoked INTEGER NOT NULL DEFAULT 0)"
    )
    .run();
  await db
    .prepare(
      "CREATE TABLE IF NOT EXISTS contrib (client_id TEXT PRIMARY KEY, payload TEXT NOT NULL, updated INTEGER NOT NULL)"
    )
    .run();
  await db
    .prepare(
      "CREATE TABLE IF NOT EXISTS merged (id INTEGER PRIMARY KEY CHECK (id = 1), payload TEXT NOT NULL, updated INTEGER NOT NULL)"
    )
    .run();
  const now = Math.floor(Date.now() / 1000);
  for (const seedId of SEED_DOWNLOAD_CLIENT_IDS) {
    await db
      .prepare(
        "INSERT OR IGNORE INTO downloads (client_id, download_proof, app_version, exe_sha256, registered_at, last_seen, revoked) VALUES (?1, 'seed', '1.0.27', 'seed', ?2, ?2, 0)"
      )
      .bind(seedId, now)
      .run();
  }
  schemaInitialized = true;
}

function isValidClientId(cid) {
  return typeof cid === "string" && CLIENT_ID_RE.test(cid.trim().toLowerCase());
}

async function computeDownloadProof(clientId, appVersion, exeSha256, secretContext) {
  const cid = String(clientId || "").trim().toLowerCase();
  const ver = String(appVersion || "").trim();
  const sha = String(exeSha256 || "").trim().toLowerCase();
  const keyBytes = new TextEncoder().encode(secretContext || DEFAULT_ATTEST_CONTEXT);
  const msgBytes = new TextEncoder().encode(`clickshift-download-v1|${cid}|${ver}|${sha}`);
  const cryptoKey = await crypto.subtle.importKey(
    "raw",
    keyBytes,
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const sig = await crypto.subtle.sign("HMAC", cryptoKey, msgBytes);
  return Array.from(new Uint8Array(sig))
    .slice(0, 16)
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

function timingSafeEqualHex(a, b) {
  if (typeof a !== "string" || typeof b !== "string" || a.length !== b.length || a.length === 0) {
    return false;
  }
  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return diff === 0;
}

async function verifyDownloadProof(clientId, proof, appVersion, exeSha256, env) {
  if (!isValidClientId(clientId) || typeof proof !== "string" || proof.length < 16) {
    return false;
  }
  const secret = (env && env.CLICKSHIFT_DOWNLOAD_KEY) || DEFAULT_ATTEST_CONTEXT;
  const expected = await computeDownloadProof(clientId, appVersion, exeSha256, secret);
  return timingSafeEqualHex(proof.trim().toLowerCase(), expected.toLowerCase());
}

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

async function recomputeMerged(db, now) {
  const allRows = await db
    .prepare(
      "SELECT c.payload FROM contrib c INNER JOIN downloads d ON c.client_id = d.client_id WHERE d.revoked = 0 ORDER BY c.updated DESC LIMIT 500"
    )
    .all();
  const merged = mergeStates(allRows.results || []);
  await db
    .prepare(
      "INSERT INTO merged (id, payload, updated) VALUES (1, ?1, ?2) ON CONFLICT(id) DO UPDATE SET payload = excluded.payload, updated = excluded.updated"
    )
    .bind(JSON.stringify(merged), now)
    .run();
  return merged;
}

export default {
  async fetch(request, env) {
    await ensureSchema(env.DB);
    const url = new URL(request.url);
    const now = Math.floor(Date.now() / 1000);

    // 1. Tracked Download Endpoint: registers a new download client_id in D1 and redirects to GitHub Release EXE
    if (request.method === "GET" && url.pathname === "/download") {
      const reqCid = (url.searchParams.get("client_id") || "").trim().toLowerCase();
      let clientId = isValidClientId(reqCid) ? reqCid : "";
      if (!clientId) {
        const rand = new Uint8Array(8);
        crypto.getRandomValues(rand);
        clientId = Array.from(rand)
          .map((b) => b.toString(16).padStart(2, "0"))
          .join("");
      }
      const ver = (url.searchParams.get("version") || "latest").slice(0, 32);
      const proof = await computeDownloadProof(
        clientId,
        ver,
        "download",
        (env && env.CLICKSHIFT_DOWNLOAD_KEY) || DEFAULT_ATTEST_CONTEXT
      );
      await env.DB.prepare(
        "INSERT INTO downloads (client_id, download_proof, app_version, exe_sha256, registered_at, last_seen, revoked) VALUES (?1, ?2, ?3, 'download', ?4, ?4, 0) ON CONFLICT(client_id) DO UPDATE SET last_seen = excluded.last_seen"
      )
        .bind(clientId, proof, ver, now)
        .run();

      if (url.searchParams.get("json") === "1") {
        return new Response(
          JSON.stringify({ ok: true, client_id: clientId, download_proof: proof, download_url: LATEST_EXE_DOWNLOAD_URL }),
          { headers: { "Content-Type": "application/json" } }
        );
      }
      return Response.redirect(LATEST_EXE_DOWNLOAD_URL, 302);
    }

    // 2. Register a downloaded ClickShift installation's client_id
    if (request.method === "POST" && url.pathname === "/downloads/register") {
      let body;
      try {
        body = await request.json();
      } catch {
        return new Response(JSON.stringify({ ok: false, error: "Invalid JSON" }), {
          status: 400,
          headers: { "Content-Type": "application/json" },
        });
      }
      const cid = String(body.client_id || "").trim().toLowerCase();
      const ver = String(body.app_version || body.version || "").trim().slice(0, 32);
      const sha = String(body.exe_sha256 || "").trim().toLowerCase().slice(0, 64);
      const proof = String(body.download_proof || "").trim().toLowerCase();
      if (!isValidClientId(cid)) {
        return new Response(JSON.stringify({ ok: false, error: "Invalid client_id format" }), {
          status: 400,
          headers: { "Content-Type": "application/json" },
        });
      }
      const proofOk = await verifyDownloadProof(cid, proof, ver, sha, env);
      if (!proofOk) {
        return new Response(JSON.stringify({ ok: false, error: "Invalid download attestation proof" }), {
          status: 403,
          headers: { "Content-Type": "application/json" },
        });
      }
      const existing = await env.DB.prepare("SELECT revoked FROM downloads WHERE client_id = ?1").bind(cid).first();
      if (existing && Number(existing.revoked) === 1) {
        return new Response(JSON.stringify({ ok: false, error: "Client ID is revoked" }), {
          status: 403,
          headers: { "Content-Type": "application/json" },
        });
      }
      if (!existing && env && env.LOCK_DOWNLOAD_REGISTRY === "1") {
        return new Response(JSON.stringify({ ok: false, error: "Download registry is locked to pre-registered IDs" }), {
          status: 403,
          headers: { "Content-Type": "application/json" },
        });
      }
      await env.DB.prepare(
        "INSERT INTO downloads (client_id, download_proof, app_version, exe_sha256, registered_at, last_seen, revoked) VALUES (?1, ?2, ?3, ?4, ?5, ?5, 0) ON CONFLICT(client_id) DO UPDATE SET download_proof = excluded.download_proof, app_version = excluded.app_version, exe_sha256 = excluded.exe_sha256, last_seen = excluded.last_seen"
      )
        .bind(cid, proof, ver, sha, now)
        .run();
      return new Response(JSON.stringify({ ok: true, registered: true, client_id: cid }), {
        headers: { "Content-Type": "application/json" },
      });
    }

    // 3. Verify whether a client_id is in the registered downloads table
    if (request.method === "GET" && url.pathname === "/downloads/verify") {
      const cid = String(url.searchParams.get("client_id") || "").trim().toLowerCase();
      if (!isValidClientId(cid)) {
        return new Response(JSON.stringify({ ok: false, registered: false }), {
          status: 400,
          headers: { "Content-Type": "application/json" },
        });
      }
      const row = await env.DB.prepare("SELECT client_id, revoked FROM downloads WHERE client_id = ?1").bind(cid).first();
      const registered = Boolean(row && Number(row.revoked) === 0);
      return new Response(JSON.stringify({ ok: true, client_id: cid, registered }), {
        headers: { "Content-Type": "application/json" },
      });
    }

    // 4. Admin management endpoint for download client IDs (allow / revoke / list)
    if (request.method === "POST" && url.pathname === "/downloads/manage") {
      const adminToken = env && env.ADMIN_TOKEN ? String(env.ADMIN_TOKEN).trim() : "";
      const authHeader = (request.headers.get("Authorization") || "").trim();
      if (!adminToken || authHeader !== `Bearer ${adminToken}`) {
        return new Response(JSON.stringify({ ok: false, error: "Unauthorized" }), {
          status: 401,
          headers: { "Content-Type": "application/json" },
        });
      }
      const cmd = await request.json();
      const action = String(cmd.action || "").toLowerCase();
      const ids = Array.isArray(cmd.client_ids)
        ? cmd.client_ids
        : cmd.client_id
        ? [cmd.client_id]
        : [];
      if (action === "allow") {
        let added = 0;
        for (const rawId of ids) {
          const cid = String(rawId || "").trim().toLowerCase();
          if (!isValidClientId(cid)) continue;
          await env.DB.prepare(
            "INSERT INTO downloads (client_id, download_proof, app_version, exe_sha256, registered_at, last_seen, revoked) VALUES (?1, 'admin', 'admin', 'admin', ?2, ?2, 0) ON CONFLICT(client_id) DO UPDATE SET revoked = 0, last_seen = excluded.last_seen"
          )
            .bind(cid, now)
            .run();
          added++;
        }
        return new Response(JSON.stringify({ ok: true, action: "allow", count: added }), {
          headers: { "Content-Type": "application/json" },
        });
      }
      if (action === "revoke") {
        let revokedCount = 0;
        for (const rawId of ids) {
          const cid = String(rawId || "").trim().toLowerCase();
          if (!isValidClientId(cid)) continue;
          await env.DB.prepare(
            "INSERT INTO downloads (client_id, download_proof, app_version, exe_sha256, registered_at, last_seen, revoked) VALUES (?1, 'revoked', 'revoked', 'revoked', ?2, ?2, 1) ON CONFLICT(client_id) DO UPDATE SET revoked = 1, last_seen = excluded.last_seen"
          )
            .bind(cid, now)
            .run();
          await env.DB.prepare("DELETE FROM contrib WHERE client_id = ?1").bind(cid).run();
          revokedCount++;
        }
        await recomputeMerged(env.DB, now);
        return new Response(JSON.stringify({ ok: true, action: "revoke", count: revokedCount }), {
          headers: { "Content-Type": "application/json" },
        });
      }
      const rows = await env.DB.prepare(
        "SELECT client_id, app_version, registered_at, last_seen, revoked FROM downloads ORDER BY last_seen DESC LIMIT 500"
      ).all();
      return new Response(JSON.stringify({ ok: true, downloads: rows.results || [] }), {
        headers: { "Content-Type": "application/json" },
      });
    }

    // 5. Pull merged community ML state
    if (request.method === "GET" && url.pathname === "/state") {
      const row = await env.DB.prepare("SELECT payload FROM merged WHERE id = 1").first();
      const body = row ? row.payload : JSON.stringify({ version: 1, clock: 0, bandit: {}, median_ms: {} });
      return new Response(body, {
        headers: { "Content-Type": "application/json", "Cache-Control": "public, max-age=60" },
      });
    }

    // 6. Push contributor ML state — strictly gated by registered download client_id
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

      const cid = String(
        parsed.client_id || request.headers.get("X-ClickShift-Client-Id") || ""
      )
        .trim()
        .toLowerCase();
      if (!isValidClientId(cid)) {
        return new Response(
          JSON.stringify({ ok: false, error: "Missing or invalid download client_id" }),
          { status: 403, headers: { "Content-Type": "application/json" } }
        );
      }

      let dlRow = await env.DB.prepare("SELECT client_id, revoked FROM downloads WHERE client_id = ?1")
        .bind(cid)
        .first();

      if (dlRow && Number(dlRow.revoked) === 1) {
        return new Response(
          JSON.stringify({ ok: false, error: "Download client_id has been revoked" }),
          { status: 403, headers: { "Content-Type": "application/json" } }
        );
      }

      // If not yet in `downloads`, only accept if the client provides a valid download attestation proof
      // and the registry is not locked by admin
      if (!dlRow) {
        const proof = String(
          parsed.download_proof || request.headers.get("X-ClickShift-Download-Proof") || ""
        )
          .trim()
          .toLowerCase();
        const ver = String(parsed.app_version || "").trim().slice(0, 32);
        const sha = String(parsed.exe_sha256 || "").trim().toLowerCase().slice(0, 64);
        const proofValid = await verifyDownloadProof(cid, proof, ver, sha, env);
        if (!proofValid || (env && env.LOCK_DOWNLOAD_REGISTRY === "1")) {
          return new Response(
            JSON.stringify({
              ok: false,
              error: "Client ID does not match registered download IDs — cloud update rejected",
            }),
            { status: 403, headers: { "Content-Type": "application/json" } }
          );
        }
        await env.DB.prepare(
          "INSERT INTO downloads (client_id, download_proof, app_version, exe_sha256, registered_at, last_seen, revoked) VALUES (?1, ?2, ?3, ?4, ?5, ?5, 0)"
        )
          .bind(cid, proof, ver, sha, now)
          .run();
      } else {
        await env.DB.prepare("UPDATE downloads SET last_seen = ?1 WHERE client_id = ?2")
          .bind(now, cid)
          .run();
      }

      await env.DB.prepare(
        "INSERT INTO contrib (client_id, payload, updated) VALUES (?1, ?2, ?3) ON CONFLICT(client_id) DO UPDATE SET payload = excluded.payload, updated = excluded.updated"
      )
        .bind(cid, text, now)
        .run();

      const merged = await recomputeMerged(env.DB, now);
      return new Response(JSON.stringify({ ok: true, clock: merged.clock, client_id: cid }), {
        headers: { "Content-Type": "application/json" },
      });
    }

    return new Response("Not Found", { status: 404 });
  },
};
