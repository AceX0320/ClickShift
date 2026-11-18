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
import itertools
import json
import math
import os
import sys
import threading
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator, Optional

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
_ASCII_BYTES = bytes(range(128))


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
        head4k = head[:4096]
        non_ascii = len(head4k.translate(None, _ASCII_BYTES))
        media_ratio = min(non_ascii / max(len(head4k), 1), 1.0)
        complexity = min(md_or_tags / 40.0, 1.0)

    return float(struct_density), float(media_ratio), float(complexity)


def _features(src: Path, dst_ext: str, src_size: Optional[int] = None) -> list[float]:
    if src_size is None:
        try:
            size = src.stat().st_size
        except OSError:
            size = 0
    else:
        size = src_size
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
        self._total_samples = 0
        self._ridge_I = np.eye(self.dim) * self.ridge
        self._inv_ridge_I = np.eye(self.dim) * (1.0 / self.ridge)
        self._mature_prior_I = (1.0 - self.gamma) * self._ridge_I
        self._lock = threading.Lock()

    def _ensure(self, arm: str) -> None:
        if arm not in self.A:
            self.A[arm] = self._ridge_I.copy()
            self.A_inv[arm] = self._inv_ridge_I.copy()
            self.b[arm] = np.zeros(self.dim)
            self.n[arm] = 0
        elif arm not in self.A_inv:
            self.A_inv[arm] = np.linalg.inv(self.A[arm])

    def _score_unlocked(self, arm: str, x: Any) -> float:
        self._ensure(arm)
        A_inv = self.A_inv[arm]
        theta = A_inv @ self.b[arm]
        mean = float(theta @ x)
        ucb = self.alpha * float(math.sqrt(max(float(x @ A_inv @ x), 0.0)))
        return mean + ucb

    def score(self, arm: str, x: Any) -> float:
        with self._lock:
            return self._score_unlocked(arm, x)

    def total_samples(self) -> int:
        with self._lock:
            return self._total_samples

    def rank_routes(self, pair_key: str, x: Any, routes: list, min_trust: int) -> tuple[list, dict[str, float], int]:
        scored = []
        scores: dict[str, float] = {}
        with self._lock:
            total = self._total_samples
            for pos, r in enumerate(routes):
                arm = f"{pair_key}|{r.name}"
                n = self.n.get(arm, 0)
                if n >= min_trust:
                    s = self._score_unlocked(arm, x)
                else:
                    s = (r.weight / 100.0) * 0.35 + (self._score_unlocked(arm, x) if n else 0.0)
                    s += 0.05 * (min_trust - n) / min_trust
                scores[r.name] = s
                scored.append((-s, -r.weight, pos, r))
        scored.sort(key=lambda t: (t[0], t[1], t[2]))
        return [t[3] for t in scored], scores, total

    def update(self, arm: str, x: Any, reward: float) -> None:
        with self._lock:
            self._ensure(arm)
            # Apply gentle exponential forgetting on mature arms so the model
            # adapts if the user installs Office or updates LibreOffice later.
            if self.n[arm] >= 50 and 0.9 < self.gamma < 1.0:
                g = self.gamma
                self.A[arm] = g * self.A[arm] + self._mature_prior_I + np.outer(x, x)
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
            self._total_samples += 1

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
            for a, s in itertools.islice(state.items(), 512):          # bound #arms from remote
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
            self._total_samples = sum(self.n.values())


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
        self.local_median_ms: dict[str, float] = {} # local-only latencies uploaded to cloud
        self.bandit = _LinearBandit()
        self.local_bandit = _LinearBandit()         # local-only observations (prevents echo)
        self.cloud: Any = None                      # optional CloudSync (see cloud.py)
        self._io_lock = threading.Lock()
        self._defer_save_depth = 0
        self._dirty = False
        self._load_state()

    @contextmanager
    def batch_persistence(self) -> Iterator[None]:
        """Coalesce bandit.json writes across a multi-file batch, flushing on exit."""
        with self._io_lock:
            self._defer_save_depth += 1
        try:
            yield
        finally:
            should_save = False
            with self._io_lock:
                self._defer_save_depth = max(0, self._defer_save_depth - 1)
                if self._defer_save_depth == 0 and self._dirty:
                    should_save = True
            if should_save:
                self._save_state()

    def attach_cloud(self, cloud: Any) -> None:
        """Wire a clickshift.cloud.CloudSync instance into this selector."""
        self.cloud = cloud

    # -- persistence -------------------------------------------------------- #
    def _load_state(self) -> None:
        try:
            st = json.loads(self.state_file.read_text(encoding="utf-8"))
            self.bandit.import_state(st.get("bandit", {}))
            self.median_ms = st.get("median_ms", {})
            if isinstance(st.get("local_bandit"), dict):
                self.local_bandit.import_state(st["local_bandit"])
                self.local_median_ms = st.get("local_median_ms", {})
            else:
                self.local_bandit.import_state(st.get("bandit", {}))
                self.local_median_ms = dict(self.median_ms)
        except FileNotFoundError:
            pass
        except Exception:
            pass

    def _save_state(self) -> None:
        with self._io_lock:
            tmp = self.state_file.with_suffix(f".{threading.get_ident()}.tmp")
            try:
                tmp.write_text(
                    json.dumps({
                        "bandit": self.bandit.export(),
                        "median_ms": dict(self.median_ms),
                        "local_bandit": self.local_bandit.export(),
                        "local_median_ms": dict(self.local_median_ms),
                    }),
                    encoding="utf-8",
                )
                tmp.replace(self.state_file)
                self._dirty = False
            except Exception:
                try:
                    tmp.unlink(missing_ok=True)
                except Exception:
                    pass

    def export_local_payload(self) -> tuple[dict[str, Any], dict[str, float]]:
        """Return only locally observed (bandit_arms, median_ms) so remote priors never echo back."""
        return self.local_bandit.export(), dict(self.local_median_ms)

    def merge_remote_state(
        self,
        remote_bandit: dict[str, Any],
        remote_medians: Optional[dict[str, Any]] = None,
        max_prior_samples: int = 24,
    ) -> int:
        """Blend validated remote pooled priors with local observations (preserving local adaptation)."""
        if not isinstance(remote_bandit, dict):
            return 0
        tmp = _LinearBandit(
            dim=self.bandit.dim,
            alpha=self.bandit.alpha,
            ridge=self.bandit.ridge,
            gamma=self.bandit.gamma,
        )
        tmp.import_state(remote_bandit)
        ridge_I = self.bandit._ridge_I
        combined: dict[str, Any] = {}
        with self.local_bandit._lock:
            loc_A = {k: v.copy() for k, v in self.local_bandit.A.items()}
            loc_b = {k: v.copy() for k, v in self.local_bandit.b.items()}
            loc_n = dict(self.local_bandit.n)
        for arm in set(tmp.A) | set(loc_A):
            has_rem = arm in tmp.A
            has_loc = arm in loc_A
            if has_rem:
                n_r = tmp.n[arm]
                if max_prior_samples > 0 and n_r > max_prior_samples:
                    scale = max_prior_samples / float(n_r)
                    A_r = (1.0 - scale) * ridge_I + scale * tmp.A[arm]
                    b_r = tmp.b[arm] * scale
                    eff_n_r = max_prior_samples
                else:
                    A_r = tmp.A[arm]
                    b_r = tmp.b[arm]
                    eff_n_r = n_r
            else:
                A_r = ridge_I
                b_r = np.zeros(self.bandit.dim)
                eff_n_r = 0
            if has_loc:
                dA_l = loc_A[arm] - ridge_I
                db_l = loc_b[arm]
                n_l = loc_n[arm]
            else:
                dA_l = np.zeros_like(ridge_I)
                db_l = np.zeros(self.bandit.dim)
                n_l = 0
            M = 0.5 * ((A_r + dA_l) + (A_r + dA_l).T)
            if np.linalg.eigvalsh(M).min() <= 1e-6:
                M = loc_A[arm] if has_loc else A_r
                b_vec = loc_b[arm] if has_loc else b_r
                n_tot = n_l if has_loc else eff_n_r
            else:
                b_vec = b_r + db_l
                n_tot = min(self.bandit.MAX_SAMPLES_PER_ARM, eff_n_r + n_l)
            combined[arm] = {"A": M.tolist(), "b": b_vec.tolist(), "n": n_tot}
        if combined:
            self.bandit.import_state(combined)
        if isinstance(remote_medians, dict):
            with self._io_lock:
                for pair, ms in itertools.islice(remote_medians.items(), 2048):
                    try:
                        ms_val = float(ms)
                    except (TypeError, ValueError):
                        continue
                    if not (0.0 < ms_val <= 86_400_000.0):
                        continue
                    cur = self.median_ms.get(pair)
                    if cur is None or ms_val < cur:
                        self.median_ms[pair] = ms_val
        self._save_state()
        return len(tmp.A)

    # -- API used by engine -------------------------------------------------- #
    def extract_features(self, src: Path, dst_ext: str, src_size: Optional[int] = None) -> Any:
        return np.array(_features(src, dst_ext, src_size=src_size), dtype=float)

    def rank(self, src: Path, dst_ext: str, routes: list,
             features: Optional[Any] = None, src_size: Optional[int] = None) -> list:
        ranked, _ = self.rank_debug(src, dst_ext, routes, features=features, src_size=src_size)
        return ranked

    def rank_debug(self, src: Path, dst_ext: str, routes: list,
                   features: Optional[Any] = None, src_size: Optional[int] = None) -> tuple[list, dict]:
        """Return (routes sorted best-first, debug info). Stable, never drops fallbacks."""
        if len(routes) <= 1:
            return routes, {"enabled": True, "scores": {},
                            "order_after": [r.name for r in routes],
                            "history_samples": self.bandit.total_samples()}
        x = features if features is not None else self.extract_features(src, dst_ext, src_size=src_size)
        pair_key = f"{src.suffix.lstrip('.').lower()}>{dst_ext.lstrip('.').lower()}"
        ranked, scores, total_samples = self.bandit.rank_routes(
            pair_key, x, routes, self.MIN_SAMPLES_TO_TRUST
        )
        info = {"enabled": True, "scores": scores,
                "order_after": [r.name for r in ranked],
                "history_samples": total_samples}
        return ranked, info

    def _compact_history_unlocked(self) -> None:
        tmp = self.history_file.with_suffix(f".hist.{threading.get_ident()}.tmp")
        try:
            with open(self.history_file, "rb") as fh:
                fh.seek(0, os.SEEK_END)
                end_pos = fh.tell()
                tail_bytes = min(end_pos, 600_000)
                fh.seek(end_pos - tail_bytes)
                raw = fh.read(tail_bytes)
            if end_pos > tail_bytes:
                nl = raw.find(b"\n")
                if nl != -1:
                    raw = raw[nl + 1:]
            lines = raw.decode("utf-8", errors="replace").splitlines()[-3000:]
            tmp.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")
            tmp.replace(self.history_file)
        except Exception:
            try:
                tmp.unlink(missing_ok=True)
            except Exception:
                pass

    def record(self, src: Path, dst_ext: str, backend: str, elapsed_ms: float, ok: bool,
               features: Optional[Any] = None, src_size: Optional[int] = None) -> None:
        pair_key = f"{src.suffix.lstrip('.').lower()}>{dst_ext.lstrip('.').lower()}"
        arm = f"{pair_key}|{backend}"
        x = features if features is not None else self.extract_features(src, dst_ext, src_size=src_size)

        key = pair_key
        med = self.median_ms.get(key)
        ref = med if med and med > 0 else max(elapsed_ms, 1.0)
        speed_reward = 1.0 - 0.5 * abs(math.log(max(elapsed_ms, 1.0) / ref))
        reward = (1.0 if ok else -1.0) + max(min(speed_reward, 1.0), -1.0) * 0.5

        self.bandit.update(arm, x, float(reward))
        self.local_bandit.update(arm, x, float(reward))
        # exponential moving average of successful latency per pair
        if ok:
            with self._io_lock:
                prev = self.median_ms.get(key)
                self.median_ms[key] = elapsed_ms if prev is None else prev * 0.7 + elapsed_ms * 0.3
                prev_loc = self.local_median_ms.get(key)
                self.local_median_ms[key] = (
                    elapsed_ms if prev_loc is None else prev_loc * 0.7 + elapsed_ms * 0.3
                )

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
        defer = False
        with self._io_lock:
            self._dirty = True
            defer = self._defer_save_depth > 0
            try:
                with open(self.history_file, "a", encoding="utf-8") as fh:
                    fh.write(json.dumps(entry) + "\n")
                if self.history_file.stat().st_size > 1_000_000:
                    self._compact_history_unlocked()
            except Exception:
                pass
        if not defer:
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
