# ClickShift Publishing, Code Signing & Cloud ML Guide

This guide covers the external configuration steps outside this repository that complete the full security & Windows trust stack.

---

## 1. Enable Offline ECDSA Release Signing (Protects All Auto-Updates)

1. Generate your ECDSA P-256 release keypair once on your own machine:
   ```powershell
   python scripts/sign_release.py keygen
   ```
2. Paste the printed **PUBLIC** key into `TRUSTED_ECDSA_PUBKEYS_B64` in [`src/clickshift/updater.py`](../src/clickshift/updater.py) (keep an old key alongside a new key whenever rotating).
3. On GitHub (**Settings -> Environments -> New environment -> `release`**):
   * Enable **Required reviewers** (add yourself) so no branch/PR or rogue workflow run can ever sign a release without manual approval.
   * Add an Environment Secret named **`CLICKSHIFT_SIGNING_KEY_B64`** containing the **PRIVATE** key printed by `keygen` (and store a backup offline in a password manager).
4. Whenever you push a tag (`git tag v1.0.19 && git push origin v1.0.19`), `release.yml` builds the EXE, generates `SHA256SUMS.txt`, signs `ClickShift-v*.exe.sig` after your Environment approval, and attaches all assets cleanly.

---

## 2. GitHub Repository Security Settings

In GitHub **Settings**:
1. Enforce **Two-Factor Authentication (2FA)** on your GitHub account.
2. Add a **Ruleset / Branch Protection** for `main`: require PRs or restrict force pushes/deletions.
3. Add a **Tag Ruleset** for `v*`: restrict tag creation/deletion strictly to repo admins so third-party contributors or Actions cannot create fake version tags.

---

## 3. Windows SmartScreen & Zero-UAC Win11 Menus

ClickShift already supports both distribution models seamlessly (`is_packaged_store_app()` automatically disables GitHub self-update when packaged in the Store per Store Policy 10.2.5):

### Option A: Microsoft Store (`$0` Individual Developer Account — Works Right Now)
1. Register for free at `storedeveloper.microsoft.com` (government ID + selfie verification).
2. Copy the Partner Center **Package Identity Name** and **Publisher DN**, then build the Store MSIX using the bundled packager:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\build_store_msix.ps1 -IdentityName "<PartnerCenterName>" -Publisher "<PartnerCenterPublisher>"
   ```
3. Submit `dist/ClickShift-Store-v*.msix` to Partner Center — Microsoft **hosts and signs the MSIX for free**, eliminating SmartScreen prompts, Defender unknown-publisher warnings, `LocalMachine\Root`, and UAC prompts for Store users (ClickShift detects its Store package identity at runtime and skips the standalone self-updater and UAC shell installer cleanly).

### Option B: SignPath Foundation (`$0` Free Open-Source Authenticode Signing for Direct GitHub Releases)
Once the public repository establishes contributor/star traction:
1. Keep the included MIT `LICENSE`.
2. Apply at `signpath.org` linking your GitHub repository.
3. Update `Publisher` in `src/shell_ext/AppxManifest.xml` to match your assigned SignPath certificate subject and add `signpath/github-action` to `.github/workflows/release.yml` inside the `release` environment.

---

## 4. Optional Community Cloud ML Sync (`cloud_sync`)

ClickShift defaults `cloud_enabled` to `False`. Never distribute a GitHub Gist PAT to end users (`gist` scope grants write access to all your gists).

If you want public community ML pooling:
1. Deploy [`cloudflare/worker.js`](../cloudflare/worker.js) on Cloudflare's free plan using **Cloudflare D1** (`100,000` row writes/day free).
2. Users point ClickShift at the HTTPS endpoint with zero tokens:
   ```bat
   ClickShift.exe config cloud_url=https://your-worker.workers.dev cloud_sync=on
   ```
