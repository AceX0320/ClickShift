import assert from "node:assert/strict";
import { CloudPoolDO, mergeStates, validateClientState } from "./src/index.js";

const DIM = 10;
const eye = (scale = 8.0) =>
  Array.from({ length: DIM }, (_, i) =>
    Array.from({ length: DIM }, (_, j) => (i === j ? scale : 0.0))
  );

// 1. Validate rejects NaN, wrong dims, non-PSD, bad arm names, huge latencies
assert.equal(validateClientState(null), null);
assert.equal(
  validateClientState({
    version: 1,
    client_id: "user_alpha_01",
    bandit: { "bad_arm_no_sep": { A: eye(), b: new Array(DIM).fill(0), n: 1 } },
  }),
  null
);
const negEig = eye(-2.0);
assert.equal(
  validateClientState({
    version: 1,
    client_id: "user_alpha_01",
    bandit: { "pdf>docx|libreoffice": { A: negEig, b: new Array(DIM).fill(0), n: 2 } },
  }),
  null
);
const asym = eye(8.0);
asym[0][1] = 3.0;
assert.equal(
  validateClientState({
    version: 1,
    client_id: "user_alpha_01",
    bandit: { "pdf>docx|libreoffice": { A: asym, b: new Array(DIM).fill(0), n: 2 } },
  }),
  null
);

// 2. Valid payloads validate and merge properly across clients
const A1 = eye(9.0);
const b1 = new Array(DIM).fill(1.0);
const A2 = eye(11.0);
const b2 = new Array(DIM).fill(3.0);

const v1 = validateClientState({
  version: 1,
  client_id: "client_11111111",
  clock: 1,
  bandit: { "pptx>pdf|office-com": { A: A1, b: b1, n: 4 } },
  median_ms: { "pptx>pdf": 240.0, "bad>lat": -5 },
});
const v2 = validateClientState({
  version: 1,
  client_id: "client_22222222",
  clock: 2,
  bandit: { "pptx>pdf|office-com": { A: A2, b: b2, n: 6 } },
  median_ms: { "pptx>pdf": 180.0 },
});
assert.ok(v1 && v2);
assert.equal(v1.state.median_ms["bad>lat"], undefined);

const merged = mergeStates([v1.state, v2.state], 0);
assert.equal(merged.bandit["pptx>pdf|office-com"].n, 10);
assert.equal(merged.bandit["pptx>pdf|office-com"].A[0][0], 10.0);
assert.equal(merged.bandit["pptx>pdf|office-com"].b[0], 2.0);
assert.equal(merged.median_ms["pptx>pdf"], 180.0);
assert.equal(merged.clock, 3);

// 3. Simulate Durable Object SQLite + GET / PUT lifecycle cleanly (idempotent client updates)
function makeFakeDO() {
  const contributors = new Map();
  let poolRow = null;
  const fakeSql = {
    exec(query, ...params) {
      const q = query.trim().toUpperCase();
      if (q.startsWith("CREATE TABLE")) return { toArray: () => [] };
      if (q.startsWith("SELECT CLOCK, MERGED_JSON")) {
        return { toArray: () => (poolRow ? [poolRow] : []) };
      }
      if (q.startsWith("INSERT INTO CONTRIBUTORS")) {
        const [cid, updatedAt, payload] = params;
        contributors.set(cid, { client_id: cid, updated_at: updatedAt, payload });
        return { toArray: () => [] };
      }
      if (q.startsWith("DELETE FROM CONTRIBUTORS")) {
        const limit = params[0];
        const sorted = [...contributors.values()].sort((a, b) => b.updated_at - a.updated_at);
        const keep = new Set(sorted.slice(0, limit).map((r) => r.client_id));
        for (const k of contributors.keys()) if (!keep.has(k)) contributors.delete(k);
        return { toArray: () => [] };
      }
      if (q.startsWith("SELECT PAYLOAD FROM CONTRIBUTORS")) {
        return {
          toArray: () =>
            [...contributors.values()]
              .sort((a, b) => b.updated_at - a.updated_at)
              .map((r) => ({ payload: r.payload })),
        };
      }
      if (q.startsWith("INSERT INTO POOL_STATE")) {
        poolRow = { clock: params[0], merged_json: params[1] };
        return { toArray: () => [] };
      }
      throw new Error("Unhandled fake SQL: " + query);
    },
  };
  return new CloudPoolDO({ storage: { sql: fakeSql } }, {});
}

const doObj = makeFakeDO();
const get0 = await doObj.fetch(new Request("https://cloud.example/state", { method: "GET" }));
assert.equal(get0.status, 200);
assert.deepEqual((await get0.json()).bandit, {});

const put1 = await doObj.fetch(
  new Request("https://cloud.example/state", {
    method: "PUT",
    headers: { "Content-Type": "application/json", "CF-Connecting-IP": "198.51.100.1" },
    body: JSON.stringify({
      version: 1,
      client_id: "user_aaaa1111",
      clock: 1,
      bandit: { "pptx>pdf|office-com": { A: A1, b: b1, n: 4 } },
      median_ms: { "pptx>pdf": 240.0 },
    }),
  })
);
assert.equal(put1.status, 200);

// Re-putting updated state for SAME client replaces rather than double-counting
const put1b = await doObj.fetch(
  new Request("https://cloud.example/state", {
    method: "PUT",
    headers: { "Content-Type": "application/json", "CF-Connecting-IP": "198.51.100.1" },
    body: JSON.stringify({
      version: 1,
      client_id: "user_aaaa1111",
      clock: 2,
      bandit: { "pptx>pdf|office-com": { A: A1, b: b1, n: 5 } },
      median_ms: { "pptx>pdf": 210.0 },
    }),
  })
);
assert.equal(put1b.status, 200);

const put2 = await doObj.fetch(
  new Request("https://cloud.example/state", {
    method: "PUT",
    headers: { "Content-Type": "application/json", "CF-Connecting-IP": "198.51.100.2" },
    body: JSON.stringify({
      version: 1,
      client_id: "user_bbbb2222",
      clock: 1,
      bandit: { "pptx>pdf|office-com": { A: A2, b: b2, n: 6 } },
      median_ms: { "pptx>pdf": 180.0 },
    }),
  })
);
assert.equal(put2.status, 200);

const get1 = await doObj.fetch(new Request("https://cloud.example/state", { method: "GET" }));
const finalPool = await get1.json();
assert.equal(finalPool.bandit["pptx>pdf|office-com"].n, 11); // 5 + 6, no double counting
assert.equal(finalPool.median_ms["pptx>pdf"], 180.0);

console.log("cloud_worker unit tests passed");
