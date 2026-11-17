"""ClickShift adaptive learning (lightweight ML) for backend selection.

Why this exists
---------------
For a given conversion pair (e.g. ``pdf -> docx``) several backends may be
available (``pdf2docx``, ``libreoffice``, ...).  Which one is *fastest* and
*most reliable* depends on the machine, the installed software versions and
even on characteristics of the file (size, page count proxy...).

Instead of hard-coding an order we learn it online with a **contextual
multi-armed bandit** (LinUCB-style linear reward model, pure Python/numpy,
zero heavy dependencies so it fits inside a small frozen EXE):

    context features x  = [1, log(bytes), log(path depth), src_ext one-hot,
                           dst_ext one-hot, hour-of-day sin/cos]
    per-backend linear model theta_b updated with ridge regression:
        theta_b <- (A_b + lambda I)^-1 b_b          (Woodbury / closed form)
    optimistic score = x^T theta_b + alpha * sqrt(x^T A_b^-1 x)

Reward combines speed and success:

    r = success(0/1) * 1.0  -  penalty for latency  =  success - 0.5*log(ms/median_ms)

so the selector converges to the backend that is both reliable and fast.
Every completed conversion is appended to ``~/.clickshift/history.jsonl`` which
doubles as an audit log (visible via ``clickshift stats``).

Cold start: until enough samples exist, static engine weights are used; the
bandit only *re-orders* candidates, it never removes fallbacks, therefore
correctness can never degrade - worst case the first attempt fails and the
engine transparently retries the next backend.

Set environment variable ``CLICKSHIFT_DISABLE_ML=1`` to switch this off entirely
(the "scratch that part" option - everything still works).
"""

# pyright: reportMissingImports=false, reportMissingModuleSource=false
from __future__ import annotations

import hashlib
import importlib
import json
import math
import os
import sys
import threading
import time
from pathlib import Path
from typing import Any, Optional

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    __package__ = "clickshift"

def _opt_import(name: str) -> Any:
    try:
        return importlib.import_module(name)
    except Exception:  # pragma: no cover - numpy is bundled by PyInstaller anyway
        return None


np: Any = _opt_import("numpy")
_HAVE_NUMPY = np is not None


# --------------------------------------------------------------------------- #
#  Feature extraction (compact, content-aware, pair-local)
# --------------------------------------------------------------------------- #

DIM = 10  # [bias, norm_log_bytes, is_large, depth, sin_h, cos_h,
          #  structure_density, media_ratio, complexity_proxy, interaction]

_HEADER_READ_BYTES = 32_768


def _probe_content_features(src: Path, ext: str, size: int, lb: float) -> tuple[float, float, float]:
    """Fast (<0.5 ms) header probe returning (structure_density, media_ratio, complexity).

    Distinguishes e.g. scanned image-heavy PDFs from text-centric PDFs,
    or complex multi-part OOXML/EPUB archives from simple documents.
    All values are bounded in [0.0, 1.0].
    """
    struct_density = min(lb / 18.0, 1.0)
    media_ratio = 0.0
    complexity = 0.0
    if size <= 0:
        return struct_density, media_ratio, complexity
    try:
        with open(src, "rb") as fh:
            head = fh.read(_HEADER_READ_BYTES)
    except OSError:
        return struct_density, media_ratio, complexity

    if not head:
        return struct_density, media_ratio, complexity

    if ext == "pdf" or head.startswith(b"%PDF"):
        # Count page, font, and image/XObject stream markers in the header slice
        pages = head.count(b"/Type /Page") + head.count(b"/Type/Page")
        fonts = head.count(b"/Font")
        images = head.count(b"/Subtype /Image") + head.count(b"/Subtype/Image") + head.count(b"/Image")
        tables_or_forms = head.count(b"/AcroForm") + head.count(b"/StructTreeRoot")
        struct_density = min((pages * 0.15) + (lb / 20.0), 1.0)
        total_markers = max(fonts + images, 1)
        media_ratio = min(images / total_markers, 1.0) if images else 0.0
        complexity = min((fonts * 0.05) + (tables_or_forms * 0.25) + (0.3 if b"/Encrypt" in head else 0.0), 1.0)
    elif head.startswith(b"PK\x03\x04"):  # OOXML (.docx/.xlsx/.pptx/.odt) or .epub
        # Count local file headers and media entries visible in the first 32 KB
        entries = head.count(b"PK\x03\x04")
        media_hits = head.count(b"media/") + head.count(b"Pictures/") + head.count(b".png") + head.count(b".jpg")
        xml_hits = head.count(b".xml") + head.count(b".xhtml")
        struct_density = min(entries / 25.0, 1.0)
        media_ratio = min(media_hits / max(entries, 1), 1.0)
        complexity = min((xml_hits / 20.0) + (0.25 if b"embeddings" in head or b"charts" in head else 0.0), 1.0)
    elif ext in ("png", "jpg", "jpeg", "webp", "avif", "jxl", "tif", "tiff", "gif", "bmp"):
        # Image: size proxy correlates with resolution & bit depth
        struct_density = min(lb / 16.0, 1.0)
        media_ratio = 1.0
        complexity = 1.0 if size > 2_000_000 else (size / 2_000_000.0)
    else:
        # Text / Markdown / HTML / CSV
        lines = head.count(b"\n")
        md_or_tags = head.count(b"<table") + head.count(b"```") + head.count(b"|") + head.count(b"# ")
        struct_density = min(lines / 500.0, 1.0)
        non_ascii = sum(1 for b in head[:4096] if b > 127)
        media_ratio = min(non_ascii / max(len(head[:4096]), 1), 1.0)
        complexity = min(md_or_tags / 40.0, 1.0)

    return float(struct_density), float(media_ratio), float(complexity)


def _features(src: Path, dst_ext: str) -> list[float]:
    try:
        size = src.stat().st_size
    except OSError:
        size = 0
    lb = math.log1p(max(size, 0))
    norm_lb = min(lb / 20.0, 1.5)
    is_large = 1.0 if size > 5 * 1024 * 1024 else 0.0
    depth = min(len(src.parts), 12) / 12.0
    lt = time.localtime()
    h = lt.tm_hour + lt.tm_min / 60.0
    sin_h, cos_h = math.sin(2 * math.pi * h / 24), math.cos(2 * math.pi * h / 24)
    s = src.suffix.lstrip(".").lower()
    struct_density, media_ratio, complexity = _probe_content_features(src, s, size, lb)
    interaction = struct_density * media_ratio
    return [
        1.0,
        norm_lb,
        is_large,
        depth,
        sin_h,
        cos_h,
        struct_density,
        media_ratio,
        complexity,
        interaction,
    ]


# --------------------------------------------------------------------------- #
#  LinUCB selector (with cached Sherman-Morrison inverse & discounting)
# --------------------------------------------------------------------------- #

class _LinearBandit:
    """Per-arm ridge-regression UCB bandit with O(d^2) Sherman-Morrison inverse."""

    def __init__(self, dim: int = DIM, alpha: float = 1.2, ridge: float = 8.0,
                 gamma: float = 0.995) -> None:
        self.dim = dim
        self.alpha = alpha
        self.ridge = ridge
        self.gamma = gamma
        self.A: dict[str, Any] = {}      # arm -> covariance accumulator
        self.A_inv: dict[str, Any] = {}  # arm -> cached inverse of A
        self.b: dict[str, Any] = {}      # arm -> reward vector
        self.n: dict[str, int] = {}
        self._lock = threading.Lock()

    def _ensure(self, arm: str) -> None:
        if arm not in self.A:
            self.A[arm] = np.eye(self.dim) * self.ridge
            self.A_inv[arm] = np.eye(self.dim) * (1.0 / self.ridge)
            self.b[arm] = np.zeros(self.dim)
            self.n[arm] = 0
        elif arm not in self.A_inv:
            self.A_inv[arm] = np.linalg.inv(self.A[arm])

    def score(self, arm: str, x: Any) -> float:
        with self._lock:
            self._ensure(arm)
            A_inv = self.A_inv[arm]
            theta = A_inv @ self.b[arm]
            mean = float(theta @ x)
            ucb = self.alpha * float(math.sqrt(max(float(x @ A_inv @ x), 0.0)))
            return mean + ucb

    def update(self, arm: str, x: Any, reward: float) -> None:
        with self._lock:
            self._ensure(arm)
            # Apply gentle exponential forgetting on mature arms so the model
            # adapts if the user installs Office or updates LibreOffice later.
            if self.n[arm] >= 50 and 0.9 < self.gamma < 1.0:
                g = self.gamma
                self.A[arm] = g * self.A[arm] + (1.0 - g) * (np.eye(self.dim) * self.ridge) + np.outer(x, x)
                self.b[arm] = g * self.b[arm] + reward * x
                self.A_inv[arm] = np.linalg.inv(self.A[arm])
            else:
                self.A[arm] += np.outer(x, x)
                self.b[arm] += reward * x
                # Sherman-Morrison rank-1 update: O(d^2) instead of O(d^3)
                A_inv = self.A_inv[arm]
                Ax = A_inv @ x
                denom = 1.0 + float(x @ Ax)
                if denom > 1e-9:
                    self.A_inv[arm] = A_inv - np.outer(Ax, Ax) / denom
                else:
                    self.A_inv[arm] = np.linalg.inv(self.A[arm])
            self.n[arm] += 1

    def export(self) -> dict:
        with self._lock:
            return {a: {"A": self.A[a].tolist(), "b": self.b[a].tolist(), "n": self.n[a]}
                    for a in self.A}

    MAX_SAMPLES_PER_ARM = 1_000_000   # hard cap: blocks "poison-by-count" attacks

    def import_state(self, state: dict) -> None:
        """Import remote arm state with strict validation (security boundary).

        Anything malformed, oversized, non-finite or structurally wrong is
        rejected arm-by-arm - the local bandit can never be corrupted by a
        hostile/misbehaving cloud endpoint.
        """
        if not isinstance(state, dict):
            return
        with self._lock:
            for a, s in list(state.items())[:512]:          # bound #arms from remote
                try:
                    if not isinstance(a, str) or len(a) > 128 or not isinstance(s, dict):
                        continue
                    A = np.array(s["A"], dtype=float)
                    b = np.array(s["b"], dtype=float)
                    n = int(s.get("n", 0))
                    d = self.dim
                    ok = (A.shape == (d, d) and b.shape == (d,)
                          and np.all(np.isfinite(A)) and np.all(np.isfinite(b))
                          and 0 <= n <= self.MAX_SAMPLES_PER_ARM
                          and np.allclose(A, A.T)                 # covariance must be symmetric
                          and np.linalg.eigvalsh(A).min() > -1e-6)  # ...and PSD
                    if not ok:
                        continue
                    self.A[a] = A
                    try:
                        self.A_inv[a] = np.linalg.inv(A)
                    except np.linalg.LinAlgError:
                        self.A_inv[a] = np.linalg.pinv(A)
                    self.b[a] = b
                    self.n[a] = min(n, self.MAX_SAMPLES_PER_ARM)
                except Exception:
                    continue                                # skip garbage arms


# --------------------------------------------------------------------------- #
#  Public selector used by the engine
# --------------------------------------------------------------------------- #

class BanditSelector:
    """Re-orders candidate routes using learned latencies/success rates."""

    MIN_SAMPLES_TO_TRUST = 6      # per arm, before scores beat static weights

    def __init__(self, history_file: Optional[Path] = None) -> None:
        if not _HAVE_NUMPY:
            raise RuntimeError("numpy required for ML selector")
        self.data_dir = (history_file or Path.home() / ".clickshift" / "history.jsonl").parent
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.history_file = self.data_dir / "history.jsonl"
        self.state_file = self.data_dir / "bandit.json"
        self.median_ms: dict[str, float] = {}       # "src>dst|arm" -> running median-ish
        self.bandit = _LinearBandit()
        self.cloud: Any = None                      # optional CloudSync (see cloud.py)
        self._io_lock = threading.Lock()
        self._load_state()

    def attach_cloud(self, cloud: Any) -> None:
        """Wire a clickshift.cloud.CloudSync instance into this selector."""
        self.cloud = cloud

    # -- persistence -------------------------------------------------------- #
    def _load_state(self) -> None:
        try:
            st = json.loads(self.state_file.read_text(encoding="utf-8"))
            self.bandit.import_state(st.get("bandit", {}))
            self.median_ms = st.get("median_ms", {})
        except FileNotFoundError:
            pass
        except Exception:
            pass

    def _save_state(self) -> None:
        with self._io_lock:
            tmp = self.state_file.with_suffix(f".{threading.get_ident()}.tmp")
            try:
                tmp.write_text(json.dumps({"bandit": self.bandit.export(),
                                           "median_ms": dict(self.median_ms)}),
                               encoding="utf-8")
                tmp.replace(self.state_file)
            except Exception:
                try:
                    tmp.unlink(missing_ok=True)
                except Exception:
                    pass

    # -- API used by engine -------------------------------------------------- #
    def rank(self, src: Path, dst_ext: str, routes: list) -> list:
        ranked, _ = self.rank_debug(src, dst_ext, routes)
        return ranked

    def rank_debug(self, src: Path, dst_ext: str, routes: list) -> tuple[list, dict]:
        """Return (routes sorted best-first, debug info). Stable, never drops fallbacks."""
        if len(routes) <= 1:
            return routes, {"enabled": True, "scores": {},
                            "order_after": [r.name for r in routes],
                            "history_samples": sum(self.bandit.n.values())}
        x = np.array(_features(src, dst_ext), dtype=float)
        pair_key = f"{src.suffix.lstrip('.').lower()}>{dst_ext.lstrip('.').lower()}"

        scored = []
        scores: dict[str, float] = {}
        for pos, r in enumerate(routes):
            arm = f"{pair_key}|{r.name}"
            n = self.bandit.n.get(arm, 0)
            if n >= self.MIN_SAMPLES_TO_TRUST:
                s = self.bandit.score(arm, x)
            else:
                # cold start: blend static weight into a neutral prior so that
                # exploration still happens but sensible defaults are respected
                s = (r.weight / 100.0) * 0.35 + (self.bandit.score(arm, x) if n else 0.0)
                s += 0.05 * (self.MIN_SAMPLES_TO_TRUST - n) / self.MIN_SAMPLES_TO_TRUST  # explore bonus
            scores[r.name] = s
            scored.append((-s, -r.weight, pos, r))
        scored.sort(key=lambda t: (t[0], t[1], t[2]))
        ranked = [t[3] for t in scored]
        info = {"enabled": True, "scores": scores,
                "order_after": [r.name for r in ranked],
                "history_samples": sum(self.bandit.n.values())}
        return ranked, info

    def record(self, src: Path, dst_ext: str, backend: str, elapsed_ms: float, ok: bool) -> None:
        pair_key = f"{src.suffix.lstrip('.').lower()}>{dst_ext.lstrip('.').lower()}"
        arm = f"{pair_key}|{backend}"
        x = np.array(_features(src, dst_ext), dtype=float)

        key = pair_key
        med = self.median_ms.get(key)
        ref = med if med and med > 0 else max(elapsed_ms, 1.0)
        speed_reward = 1.0 - 0.5 * abs(math.log(max(elapsed_ms, 1.0) / ref))
        reward = (1.0 if ok else -1.0) + max(min(speed_reward, 1.0), -1.0) * 0.5

        self.bandit.update(arm, x, float(reward))
        # exponential moving average of successful latency per pair
        if ok:
            with self._io_lock:
                prev = self.median_ms.get(key)
                self.median_ms[key] = elapsed_ms if prev is None else prev * 0.7 + elapsed_ms * 0.3

        # opportunistic cloud sync (fire-and-forget; only when enabled)
        try:
            if self.cloud is not None and self.cloud.enabled():
                self.cloud._schedule_now()
        except Exception:
            pass

        file_tag = f"{hashlib.sha256(src.name.encode('utf-8', 'replace')).hexdigest()[:12]}{src.suffix.lower()}"
        entry = {
            "pair": pair_key, "arm": arm, "ms": round(elapsed_ms, 1),
            "ok": ok, "reward": round(reward, 3), "file": file_tag, "t": time.time(),
        }
        with self._io_lock:
            try:
                with open(self.history_file, "a", encoding="utf-8") as fh:
                    fh.write(json.dumps(entry) + "\n")
                # Keep history bounded on disk so it never grows without limit
                if self.history_file.stat().st_size > 1_000_000:
                    recent = self.history_file.read_text(encoding="utf-8", errors="ignore").splitlines()[-5000:]
                    self.history_file.write_text("\n".join(recent) + ("\n" if recent else ""), encoding="utf-8")
            except Exception:
                pass
        self._save_state()

    # -- introspection ------------------------------------------------------- #
    def stats(self) -> dict[str, Any]:
        arms_info: dict[str, Any] = {}
        for arm, n in list(self.bandit.n.items()):
            self.bandit._ensure(arm)
            A_inv = self.bandit.A_inv[arm]
            theta = A_inv @ self.bandit.b[arm]
            arms_info[arm] = {"samples": n, "mean_reward_at_bias_ctx": round(float(theta[0]), 3)}
        return {"arms": arms_info, "pairs_ms": dict(self.median_ms)}
