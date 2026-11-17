# Security Policy

## Reporting a Vulnerability
Please report security issues privately through GitHub Security Advisories (**Security -> Advisories -> Report a vulnerability**) rather than opening a public issue. Include steps to reproduce and affected ClickShift version.

## Auto-Updater Verification
ClickShift validates every downloaded update before applying it:
1. Only HTTPS requests to `api.github.com`, `github.com/AceX0320/ClickShift/...`, `objects.githubusercontent.com`, and `release-assets.githubusercontent.com` are permitted.
2. Downloaded executables are staged inside `%LOCALAPPDATA%\ClickShift\staging\` with current-user NTFS permissions (never shared `%TEMP%`).
3. When `SHA256SUMS.txt` is present on a release, its SHA-256 digest is verified atomically before swapping the binary.
4. When `TRUSTED_ECDSA_PUBKEYS_B64` is configured in [`src/clickshift/updater.py`](src/clickshift/updater.py), ClickShift requires and verifies a matching ECDSA P-256 `.sig` signature natively via Windows `bcrypt.dll` (`CNG`). Two key slots are supported for seamless rotation.

## Maintainer Setup Checklist (To Complete Full Protection)
See [`docs/PUBLISHING_AND_CLOUD.md`](docs/PUBLISHING_AND_CLOUD.md) for configuring:
- **Release Signing Keys:** `python scripts/sign_release.py keygen` + MFA-gated `release` GitHub Environment (`CLICKSHIFT_SIGNING_KEY_B64`).
- **Windows Trust & SmartScreen:** Submitting MSIX to the Microsoft Store ($0 individual developer onboarding, automatic Microsoft code signing) and/or enrolling in SignPath Foundation once eligible.
- **Community Cloud ML Sync (Optional):** Free-tier Cloudflare Worker + Cloudflare D1 (`cloudflare/worker.js`) so end users never hold GitHub tokens.
