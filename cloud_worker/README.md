# ClickShift Global ML Cloud Worker (Free Cloudflare Workers + Durable Objects)

This directory contains a zero-dependency Cloudflare Worker + SQLite Durable Object
backend (`GET /state`, `PUT /state`, `GET /health`) so **every ClickShift user** can share
and download aggregate bandit learnings without needing personal GitHub tokens.

## Deploy (takes ~2 minutes, free Cloudflare plan)

1. From `cloud_worker/`, authenticate once with Cloudflare:
   ```powershell
   npx wrangler login
   ```
2. Deploy the Worker + Durable Object:
   ```powershell
   npx wrangler deploy
   ```
   Wrangler will print your live HTTPS endpoint URL, e.g.:
   `https://clickshift-cloud.<your-subdomain>.workers.dev`
3. Verify it responds:
   ```powershell
   curl https://clickshift-cloud.<your-subdomain>.workers.dev/health
   ```
4. To enable it for **all ClickShift users** by default:
   - Set `DEFAULT_CLOUD_URL = "https://clickshift-cloud.<your-subdomain>.workers.dev"` in `src/clickshift/cloud.py`
   - Update `DEFAULT_CONFIG` in `src/clickshift/monitor.py` (`cloud_mode: "url"`, `cloud_url: "<url>"`, `cloud_enabled: True`)
   - Bump, commit, tag, and push to build the updated `ClickShift.exe`.

Users who prefer offline mode can opt out anytime with:
```bat
ClickShift.exe config cloud_sync=off
```

## Run local Worker validation tests

```powershell
node cloud_worker/test_worker.mjs
```
