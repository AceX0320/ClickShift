# ClickShift ⚡ — ultra-fast local document converter for Windows

**One EXE. Right-click any document → converted. Learns to get faster.**

> **📥 Direct EXE Download (Latest — v1.0.18):** [**Download `ClickShift-v1.0.18.exe`**](https://github.com/AceX0320/ClickShift/releases/download/v1.0.18/ClickShift-v1.0.18.exe) · [**Latest Release Page**](https://github.com/AceX0320/ClickShift/releases/latest)  
> *(Or grab the latest CI artifact from [GitHub Actions](https://github.com/AceX0320/ClickShift/actions/workflows/build.yml) / browse [Releases](https://github.com/AceX0320/ClickShift/releases))*

ClickShift is a 100 % local, offline document conversion tool:

* **Both Windows 11 & Classic right-click menus** in Explorer — *"Convert with ClickShift"* cascading submenu + *"Open with"* quick target picker (`--choose`) exposing **all compatible conversion targets** on every document/image type.
* **Always running** — after the first double-click it registers itself in **Task Manager → Startup apps** and lives in the system tray forever, delegating right-click conversions to a warm local IPC daemon (`127.0.0.1:45829`) for zero cold-start overhead.
* **Latest-generation output only** — `.doc` never produces legacy Word 97; you get `.docx` (OOXML 2007+), `.xlsx`, `.pptx`, modern image codecs (WebP/AVIF/JPEG XL), UTF-8 text…
* **Adaptive ML routing** *(the "trick")* — a lightweight contextual bandit (LinUCB) learns from every conversion which backend is fastest/most reliable **on your machine for files like yours**, and re-orders the fallback chain accordingly. Speed improves with use; correctness can never degrade (fallbacks are always kept). If you don't want it: `clickshift config ml=off` or `CLICKSHIFT_DISABLE_ML=1`.

---

## How it works (speed & perfection)

| Conversion | Fastest backend (auto-picked) | Fallbacks |
|---|---|---|
| DOCX/XLSX/PPTX/DOC/XLS/PPT/ODT/RTF → PDF | **Microsoft Office COM** (`SaveAs2` FileFormat 17 / 32 / `ExportAsFixedFormat`) | LibreOffice headless → pure-Python **python-docx / python-pptx + PyMuPDF** (`docx`/`pptx`) |
| PDF → DOCX | **pdf2docx** (layout/table/image preserving) | **Microsoft Office Word COM** (PDF reflow) → **PyMuPDF (`fitz`) + python-docx** (with 200 DPI scanned-page rendering) → **LibreOffice Writer** (`--infilter=writer_pdf_import`, `MS Word 2007 XML`) → **pypdf + python-docx** |
| PDF / DOCX → PPTX | **PyMuPDF + python-pptx** (editable textboxes, images, 200 DPI scanned fallback) / **python-docx + python-pptx** (headings, bullets, tables, images) | LibreOffice Impress (`--infilter=impress_pdf_import` for PDF) |
| PPTX → DOCX | **python-pptx + python-docx** (titles, bullets, tables, images, speaker notes) | LibreOffice |
| Any paged document (`PDF`/`DOCX`/`PPTX`/`XLSX`/`EPUB`…) → Images (`PNG`/`JPG`/`WEBP`/…) | **PyMuPDF (`fitz`) + Pillow** (renders **every page/slide** at 200 DPI into a `<document_name>/` folder as `page_001.<ext>`…) | — |
| Legacy .doc/.xls/.ppt → modern | Office COM FileFormat **12 / 51 / 24** (= OOXML 2007+) | LibreOffice `MS * 2007 XML` filters |
| Images ↔ images / PDF / DOCX / PPTX | **Pillow** + **python-docx / python-pptx** (selecting multiple images combines them in selection order into one PDF/DOCX/PPTX) | LibreOffice |
| TXT/MD/CSV/HTML → DOCX | **python-docx** (real headings/lists/bold/code styles) | Office COM → LibreOffice |
| DOCX → TXT/MD/HTML | **python-docx** direct read | LibreOffice |
| EPUB → TXT/DOCX/HTML/MD | stdlib zip + spine-order extraction | — |
| PDF → TXT/MD/HTML/CSV | **PyMuPDF (`fitz`) / pypdf** (fast text extraction) | Office COM → LibreOffice |

Notes on *perfection*:
* LibreOffice filters used are explicitly the **2007 XML** generation ones (plus `writer_pdf_import` when reading PDFs) — never `MS Word 97`.
* Text outputs are written as UTF-8 (with BOM where Windows consumers expect it).
* Every pair has an ordered fallback chain; if all fail you get a clear error popup (and opt-in central diagnostics inside `~/.clickshift/debug/` if enabled).
* Persistent worker pool + warm local IPC daemon means no cold-start per conversion and persistent COM/bandit caches.

## The ML part (`src/clickshift/learning.py`)

For each `(source→target, backend)` arm the engine records latency + success + output quality sanity.
A LinUCB linear bandit over a dense **10-dimensional content-aware feature vector** (`[bias, norm_log_bytes, is_large, depth, sin_h, cos_h, structure_density, media_ratio, complexity_proxy, interaction]`, extracted via fast 32 KB header probing) picks the next attempt order using $O(d^2)$ Sherman–Morrison rank-1 inverse updates (`A_inv`) and exponential discounting ($\gamma = 0.995$) for non-stationary adaptation. State persists in `%USERPROFILE%\.clickshift\bandit.json` (debounced in daemon mode); every run is logged to `history.jsonl` and visible via:

```bat
ClickShift.exe stats
```

Cold start falls back to static weights; exploration bonus guarantees the model keeps testing alternatives, but a failing arm is demoted instantly and the engine transparently retries the next backend — so results stay correct even mid-learning.

## Shared learnings (optional cloud sync) — "send the improvements somewhere"

ClickShift can **pool everyone's ML learnings** so that a person who downloads the
tool later starts with an already-trained model instead of a cold one. This is
**opt-in and privacy-safe**: only aggregate numbers are exchanged (learned
weight matrices + median latencies per format pair). *Never* file names, paths
or contents. The conversion itself always stays 100% local.

### Recommended setup: a community GitHub Gist (zero backend needed)

1. **Owner (you), once:** create a secret gist containing one small JSON file
   (`merged.json` = `{"version":1,"clock":0,"bandit":{},"median_ms":{}}`),
   note its ID from the URL, then in repo settings add two GitHub Secrets:
   `CLICKSHIFT_GIST_ID` and `CLICKSHIFT_CLOUD_TOKEN` (fine-grained PAT with `gist`
   scope). The included workflow `.github/workflows/cloud-merge-cron.yml`
   runs every 6 h and merges all `contrib-*.json` files into `merged.json`.
2. Ship the EXE with defaults baked in, or let each user configure it:

```bat
ClickShift.exe config cloud_gist=<the-gist-id>
setx CLICKSHIFT_CLOUD_TOKEN <your-pat-with-gist-scope>
ClickShift.exe config cloud_sync=on
```

That's it. Every conversion now (debounced to ≤1 upload/minute) pushes your
aggregate learning as `contrib-<clientid>.json`, and periodically pulls the
pooled `merged.json` back into your local bandit — so you also benefit from
everyone else's machines. Check state any time with `ClickShift.exe stats`
("Cloud learnings sync: ON/OFF, last pull/push").

Alternative: run any tiny REST endpoint speaking `GET /state` + `PUT /state`
(Cloudflare Worker/KV, S3, nginx+file…) and use `config cloud_url=https://...`
instead of the gist.

Merge rules (`src/clickshift/cloud.py::merge_states`): sample counts sum, both
covariance matrices `A` and reward vectors `b` average across contributors
(preserving the $\theta = A^{-1} b$ scale and keeping the ridge prior
well-conditioned), latency medians take the minimum, and a Lamport clock guards
ordering — so concurrent users can never corrupt the pool.

### Security model (defence in depth)

* **Where it lives:** only two endpoints are ever contacted —
  `https://api.github.com` (gist mode) or *your own* `https://…` URL. Plaintext
  `http://`, `file://`, `ftp://` and malformed URLs are hard-refused by an
  allow-list (`_check_url`); TLS certificate verification is always on and
  cannot be disabled from config.
* **Least-privilege credentials:** the token is a fine-grained GitHub PAT with
  the `gist` scope *only* — it cannot touch repos, code, packages, issues or
  any account data. It is stored locally (`~/.clickshift/cloud_token`, created via
  `os.open(..., 0o600)`; on Windows use `setx` so it stays in your user
  profile) and is sent exclusively to `api.github.com`. Use **your own** token;
  nothing of yours is ever accessible through someone else's.
* **Nothing identifying leaves the machine:** payload = weight matrices +
  per-pair median latencies + counters. No file names, paths, document content,
  usernames or hardware IDs (client id is a random UUID). Verified by test.
* **Hostile-server protection (model poisoning):** downloaded state is treated
  as untrusted input — 2 MiB response cap, ≤512 arms, ≤1M samples/arm, every
  matrix validated (shape/finiteness/symmetry/positive-semi-definite), rewards
  clamped proportionally to sample count (`[-2.25 * n, 2.25 * n]`), latencies
  bounded to sane ranges, sha256 of accepted state recorded for tamper
  visibility. Garbage arms are dropped one by one.
* **ML can never break correctness:** even a perfectly poisoned pooled model
  can only reorder *which backend tries first*; the curated fallback chains
  still run, so conversions stay correct offline, online, or under attack.
* **Fail-open locally:** any sync error is logged and ignored; the local bandit
  keeps learning normally. Turning it off is one command:
  `ClickShift.exe config cloud_sync=off`.

---

## Download the EXE (no clone, no build)

The direct permanent download link for the latest release page is:

```
https://github.com/AceX0320/ClickShift/releases/latest
```

* **Direct Download (Latest Release Page):** [https://github.com/AceX0320/ClickShift/releases/latest](https://github.com/AceX0320/ClickShift/releases/latest)
* **Direct Download (v1.0.18 Pinned):** [https://github.com/AceX0320/ClickShift/releases/download/v1.0.18/ClickShift-v1.0.18.exe](https://github.com/AceX0320/ClickShift/releases/download/v1.0.18/ClickShift-v1.0.18.exe)
* **All Releases:** [https://github.com/AceX0320/ClickShift/releases](https://github.com/AceX0320/ClickShift/releases)
* **CI Build Artifacts (`ClickShift-windows-exe`):** [https://github.com/AceX0320/ClickShift/actions/workflows/build.yml](https://github.com/AceX0320/ClickShift/actions/workflows/build.yml)
* `/releases/latest` **always points to the newest published versioned EXE** (`ClickShift-v<version>.exe`). ClickShift checks for new GitHub releases automatically every **12 hours** in the background, updates silently, and leaves a Windows notification showing the old → new version when it restarts.

Publishing happens automatically on every push to `main` (via `.github/workflows/build.yml`) as well as on version tags (via `.github/workflows/release.yml`):

1. Push to `main` (or tag and push a version) → CI runs the full test suite, builds `ClickShift-v<version>.exe` on a Windows runner, uploads the `ClickShift-windows-exe` artifact, and publishes `ClickShift-v<version>.exe` to the latest GitHub Release:
   ```bat
   git tag v1.0.18
   git push origin v1.0.18
   ```
2. Share the link above. Anyone who downloads and double-clicks the
   EXE gets the self-install (both Windows right-click menus + Task Manager Startup entry).

**"Each EXE downloaded with the latest improvements":** ML learnings are
deliberately *not* baked into the binary (that would force a re-download every
time anyone converts a file). They live in the shared cloud pool (section
"Shared learnings"): every user pulls the pooled state on launch, so a
brand-new install of any release starts warm — improvements arrive
continuously, far faster than any rebuild could ship them. The EXE itself only
needs re-releasing when code changes.

## Build the EXE yourself

> A Windows `.exe` must be produced on Windows (PyInstaller cannot cross-compile). Two options, both ~3 minutes:

### Option A — one command on your PC
```powershell
git clone https://github.com/AceX0320/ClickShift.git
cd ClickShift
powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1
```
The script creates a venv, installs deps, runs `build.py`, copies `dist\ClickShift-v<version>.exe` to `%LOCALAPPDATA%\ClickShift`, performs the install (startup + right-click menus) and launches the tray daemon.

### Option B — free GitHub CI
Push to GitHub; the included workflow **`.github/workflows/build.yml`** runs tests on Linux and builds `ClickShift-v<version>.exe` on `windows-latest`, publishing it both as the downloadable artifact **`ClickShift-windows-exe`** and to the latest GitHub Release.

Manual build if you prefer:
```bat
py -3.12 -m venv .venv && .venv\Scripts\activate
pip install -r requirements.txt pyinstaller pywin32 pystray
python build.py        REM -> dist\ClickShift-v<version>.exe
```

## Use

1. **Double-click `ClickShift.exe` once.** It self-installs:
   * adds **"Convert with ClickShift"** cascading flyout menus (with all compatible target formats) + **"Open with → ClickShift — Convert (pick format…)"** across both the **Windows 11 primary right-click menu** and the **Classic right-click menu** for pdf/doc/docx/xls/xlsx/ppt/pptx/odt/ods/odp/rtf/txt/md/csv/html/epub/png/jpg/webp/avif/jxl/… (`HKCU` shell verbs — no admin needed);
   * writes `HKCU\Software\Microsoft\Windows\CurrentVersion\Run\ClickShift` → it now appears in **Task Manager → Startup apps** and starts with Windows automatically;
   * stays running in the tray (watch folder + warm local IPC daemon + 12-hour silent background auto-updater).
   * Explorer is notified automatically (`SHChangeNotify`); if needed, sign out/in once so the Windows 11 menu mode takes effect.
2. **Right-click a document → Convert with ClickShift → pick any target format** (or select multiple images and convert to `PDF`, `DOCX`, or `PPTX` in selection order).
3. **Clean output folders (opt-in diagnostic logs):** By default, ClickShift writes zero `*.clickshift-debug.txt` log files when converting. If needed for troubleshooting, run `ClickShift.exe config debug=onerror` or `debug=always` to save logs inside `%USERPROFILE%\.clickshift\debug\`.
4. Bonus: drop files into **`%USERPROFILE%\ClickShiftDrop`** → auto-converted.

CLI (same binary):
```bat
ClickShift.exe convert report.pdf                     REM -> report.docx (latest gen)
ClickShift.exe convert report.pdf --choose            REM native cursor popup with all valid targets
ClickShift.exe convert img1.png img2.jpg -t pdf       REM combine ordered images into one PDF
ClickShift.exe convert a.doc b.xls -t pdf             REM batch, parallel
ClickShift.exe status                                 REM detected backends / registration
ClickShift.exe stats                                  REM learned per-backend performance
ClickShift.exe config auto_target=docx ml=on auto_update=on update_interval_s=43200
ClickShift.exe uninstall                              REM removes startup entry + menus
```

## Repository layout
```
src/clickshift/engine.py     conversion engine + 5-stage PDF->DOCX & newest-gen backend routing
src/clickshift/learning.py   10-dim content-aware LinUCB bandit (Sherman-Morrison + discounting)
src/clickshift/cloud.py      opt-in privacy-safe pooled learnings sync (GitHub Gist / HTTPS)
src/clickshift/monitor.py    tray daemon, warm IPC, 12h auto-update loop, watch folder, menu installer
src/clickshift/cli.py        command line / right-click entry point + native format picker popup
src/clickshift/updater.py    GitHub silent background & interactive updater + update notifications
src/clickshift/debuglog.py   opt-in per-conversion diagnostic report writer
entry_point.py               frozen-EXE entry (double-click = install + run forever)
build.py / clickshift.spec   PyInstaller single-file EXE build configuration
scripts/install.ps1          one-shot Windows setup
scripts/make_icon.py         generates assets/clickshift.ico
tests/                       cross-platform tests (engine, ML, cloud sync, daemon, updater, debuglog)
.github/workflows/           CI: build.yml (tests + EXE artifact + latest release) & release.yml
```

## Requirements
* Windows 10/11, Python 3.10+ only to *build* (not to run the EXE).
* For best fidelity/speed install either **Microsoft Office** (COM backend) or **LibreOffice** (headless backend). `Pillow`, `python-docx`, `python-pptx`, `pdf2docx`, `PyMuPDF` (`fitz`), `pypdf`, `chardet`, and `numpy` are bundled inside the EXE so PDF/DOCX/PPTX, multi-page image export, text, and image conversions work out of the box even without Office installed.
* Pure image/text/markup/PDF conversions work with nothing else installed.


