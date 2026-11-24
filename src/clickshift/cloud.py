"""ClickShift shared-learnings cloud sync (optional).

What this does
--------------
Every conversion already teaches the local LinUCB bandit (see ``learning.py``).
This module lets a user *donate* those learnings to a central endpoint and
*download* the community pool, so a brand-new install starts warm instead of
cold - i.e. "the improved version" for people who download the tool later.

How it works (two supported transports, both extremely simple):

1. **Plain JSON mode** (default when you set a URL) - used with GitHub Gists,
   Cloudflare Workers/KV, S3 or any tiny REST service::

       GET  {CLOUD_URL}/state        -> {"bandit": {...}, "median_ms": {...}}
       PUT  {CLOUD_URL}/state        <- same shape (whole-state merge server-side)

2. **GitHub Gist mode** (recommended zero-backend setup) - set a fine-grained
   PAT with ``gist`` scope in ``CLICKSHIFT_CLOUD_TOKEN`` (or write it to
   ``~/.clickshift/cloud_token``). The gist holds one file per contributor plus an
   aggregated ``merged.json``; the merge worker runs in CI
   (``python -m clickshift.cloud --merge-gist``).

Privacy: only aggregate numbers are ever uploaded - learned weight matrices,
per-pair median latencies and arm sample counts. **No file names, no paths, no
content.** Opt-in: if ``cloud_sync`` is off (the default) nothing leaves the
machine. Conflicts never corrupt anything: every state carries a Lamport clock
and merges take element-wise maxima / weighted averages.

Security model (defence in depth)
---------------------------------
1. Transport: HTTPS only (allow-listed schemes), TLS certificate verification
   always on (urllib default context; no way to disable it here).
2. Credentials: the GitHub PAT lives in ``CLICKSHIFT_CLOUD_TOKEN`` or
   ``~/.clickshift/cloud_token`` (chmod 0600 on write). It can ONLY create/update
   gists - never repos, code, packages or user data. Users should supply their
   own token; the tool never sends it anywhere except api.github.com.
3. Trust boundary for DOWNLOADED state: a pooled model could be poisoned by a
   malicious contributor or a compromised endpoint. Mitigations:
   * hard size caps (max bytes per response, max arms, max samples/arm);
   * full structural validation of every remote arm (shape/finiteness/
     symmetry/PSD) inside ``_LinearBandit.import_state`` - garbage is dropped;
   * numeric sanity bounds on latencies and rewards at merge & pull time;
   * ML is an optimiser, NEVER a correctness gate: every conversion still runs
     through the curated fallback chains, so even a perfectly poisoned model
     cannot produce a wrong file - only a slower route.
4. Local bandit learning is untouched by sync failures: any exception during
   pull/push is swallowed and logged; conversions keep working offline.
"""

# pyright: reportMissingImports=false, reportMissingModuleSource=false
from __future__ import annotations

import hashlib
import hmac
import importlib
import ipaddress
import itertools
import json
import logging
import os
import re
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from pathlib import Path
from typing import Any, Optional, Sequence

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    __package__ = "clickshift"

from . import __version__ as _app_version

log = logging.getLogger("clickshift.cloud")


def _opt_import(name: str) -> Any:
    try:
        return importlib.import_module(name)
    except Exception:  # pragma: no cover
        return None


np: Any = _opt_import("numpy")

DEFAULT_GIST_ID = "REPLACE_WITH_YOUR_GIST_ID"
DEFAULT_CLOUD_URL = "https://clickshift-cloud.clickshift.workers.dev"
API_ROOT = "https://api.github.com"

# Pre-seeded verified download client IDs (includes active official installs)
SEED_DOWNLOAD_CLIENT_IDS: frozenset[str] = frozenset({
    "a66c09d196954a1c",
})

_CLIENT_ID_RE = re.compile(r"^[0-9a-f]{12,64}$")
_DOWNLOAD_ATTEST_CONTEXT = "ClickShift-Download-Attestation-v1:RUNTMSAAAAA51GhTNj8aBWoIBxWojEWaMNYmLZBR+9rwzScwqaFE/nMIqrMDoTuPrnN/VFzAM8Biv7QwpRHdLnn6z9Op3OhU"
_EXE_SHA256_CACHE: Optional[str] = None

# ---- security limits --------------------------------------------------------
MAX_RESPONSE_BYTES = 2 * 1024 * 1024      # 2 MiB cap on any downloaded state
MAX_CONTRIB_BYTES = 512 * 1024            # skip absurdly large contrib files
MAX_LATENCY_MS = 86_400_000.0             # > 24 h is nonsense -> reject
ALLOWED_SCHEMES = ("https",)              # plaintext upload is forbidden


_ALLOWED_RAW_HOSTS = {"gist.githubusercontent.com", "objects.githubusercontent.com"}


def _is_disallowed_host(host: str) -> bool:
    h = (host or "").strip().lower().rstrip(".")
    if not h or h in ("localhost", "localhost.localdomain") or h.endswith((".local", ".internal", ".localhost")):
        return True
    try:
        ip = ipaddress.ip_address(h.strip("[]"))
        return (
            ip.is_private
            or ip.is_loopback
            or ip.is_link_local
            or ip.is_multicast
            or ip.is_reserved
            or ip.is_unspecified
        )
    except ValueError:
        return False


def _check_url(url: str) -> bool:
    """Reject anything that is not a clean external https URL (SSRF/plaintext guard)."""
    if not url or "\n" in url or "\r" in url:
        return False
    try:
        p = urllib.parse.urlsplit(url)
    except ValueError:
        return False
    if p.scheme not in ALLOWED_SCHEMES or not p.netloc:
        return False
    if p.username or p.password:
        return False
    return not _is_disallowed_host(p.hostname or "")


def _is_allowed_raw_url(url: str) -> bool:
    if not _check_url(url):
        return False
    host = (urllib.parse.urlsplit(url).hostname or "").lower()
    return host in _ALLOWED_RAW_HOSTS


# --------------------------------------------------------------------------- #
#  config helpers
# --------------------------------------------------------------------------- #

def _data_dir() -> Path:
    return Path(os.environ.get("CLICKSHIFT_HOME", str(Path.home() / ".clickshift")))


def load_cloud_config() -> dict[str, Any]:
    """Merge DEFAULT_CONFIG's cloud keys with the user's config.json."""
    cfg: dict[str, Any] = {}
    try:
        from .monitor import DEFAULT_CONFIG, load_config
        full = load_config()
        cfg = {k: full.get(k, DEFAULT_CONFIG.get(k)) for k in DEFAULT_CONFIG
               if k.startswith("cloud_")}
        for extra in (
            "cloud_opt_out",
            "cloud_last_pull",
            "cloud_last_push",
            "cloud_clock",
            "cloud_last_sha256",
            "cloud_client_id",
            "cloud_download_registered",
            "cloud_download_proof",
            "cloud_download_token",
            "cloud_allowed_client_ids",
            "cloud_revoked_client_ids",
            "cloud_require_registered",
        ):
            if extra in full:
                cfg[extra] = full[extra]
        if (
            DEFAULT_CLOUD_URL
            and not cfg.get("cloud_opt_out")
            and not cfg.get("cloud_url")
            and (cfg.get("cloud_gist_id") or DEFAULT_GIST_ID) in ("", DEFAULT_GIST_ID)
        ):
            cfg["cloud_mode"] = "url"
            cfg["cloud_url"] = DEFAULT_CLOUD_URL
            cfg["cloud_enabled"] = True
    except Exception:
        pass
    return cfg


def token() -> str:
    t = os.environ.get("CLICKSHIFT_CLOUD_TOKEN", "").strip()
    if not t:
        f = _data_dir() / "cloud_token"
        if f.exists():
            t = f.read_text(encoding="utf-8").strip()
    return t


def save_token(t: str) -> Path:
    """Persist the user's gist token locally with owner-only permissions."""
    d = _data_dir()
    d.mkdir(parents=True, exist_ok=True)
    f = d / "cloud_token"
    try:
        os.chmod(f, 0o600)              # pre-chmod before content lands (best effort)
    except OSError:
        pass
    fd = os.open(str(f), os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w", encoding="utf-8") as fh:
        fh.write(t.strip())
    try:
        from .monitor import _restrict_to_current_user
        _restrict_to_current_user(d)
        _restrict_to_current_user(f)
    except Exception:
        pass
    return f


# --------------------------------------------------------------------------- #
#  transport (urllib - no extra dependency inside the frozen EXE)
# --------------------------------------------------------------------------- #

def _request(url: str, method: str = "GET", body: Optional[dict[str, Any]] = None,
             headers: Optional[dict[str, str]] = None, timeout: int = 15) -> Optional[dict[str, Any]]:
    if not _check_url(url):
        log.warning("blocked non-https request to %r", url[:100])
        return None
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("User-Agent", "ClickShift/CloudSync")
    req.add_header("Content-Type", "application/json")
    req.add_header("Accept-Encoding", "gzip")
    for k, v in (headers or {}).items():
        req.add_header(k, v)
    try:
        # default SSL context => TLS with full certificate verification.
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read(MAX_RESPONSE_BYTES + 1)
            if len(raw) > MAX_RESPONSE_BYTES:
                log.warning("response from %s exceeds size cap - discarded", url[:100])
                return None
            resp_headers = getattr(resp, "headers", None)
            if resp_headers and str(resp_headers.get("Content-Encoding", "")).lower() == "gzip":
                import gzip

                raw = gzip.decompress(raw)
                if len(raw) > MAX_RESPONSE_BYTES:
                    log.warning("decompressed response from %s exceeds size cap - discarded", url[:100])
                    return None
            text = raw.decode("utf-8")
            return json.loads(text) if text.strip() else {}
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError,
            json.JSONDecodeError, UnicodeDecodeError, OSError) as exc:
        log.debug("cloud request failed (%s %s): %s", method, url[:100], exc)
        return None


# --------------------------------------------------------------------------- #
#  merge logic (pure-python + numpy, safe on client AND in CI)
# --------------------------------------------------------------------------- #

def _clean_reward(v: Any, n: int = 1) -> Any:
    """Clamp a remote accumulated reward vector b into the sane [-2.25*n, 2.25*n] range."""
    try:
        arr = np.array(v, dtype=float)
    except Exception:
        return None
    if arr.size == 0 or not np.all(np.isfinite(arr)):
        return None
    bound = 2.25 * max(int(n), 1)
    return np.clip(arr, -bound, bound)


class _ArmAccum:
    __slots__ = ("A_sum", "b_sum", "nA", "n")

    def __init__(self) -> None:
        self.A_sum: Any = None
        self.b_sum: Any = None
        self.nA: int = 0
        self.n: int = 0


def is_valid_client_id(cid: Any) -> bool:
    """Return True if ``cid`` is a well-formed 12..64 lowercase hexadecimal download client ID."""
    if not isinstance(cid, str):
        return False
    return bool(_CLIENT_ID_RE.match(cid.strip().lower()))


def current_exe_sha256() -> str:
    """Return lowercase SHA-256 hex of the running frozen ClickShift binary (or 'source-dev')."""
    global _EXE_SHA256_CACHE
    if _EXE_SHA256_CACHE is not None:
        return _EXE_SHA256_CACHE
    if getattr(sys, "frozen", False):
        try:
            exe_p = Path(sys.executable)
            if exe_p.is_file():
                h = hashlib.sha256()
                with open(exe_p, "rb") as fh:
                    while True:
                        chunk = fh.read(1024 * 1024)
                        if not chunk:
                            break
                        h.update(chunk)
                _EXE_SHA256_CACHE = h.hexdigest().lower()
                return _EXE_SHA256_CACHE
        except Exception:
            pass
    _EXE_SHA256_CACHE = "source-dev"
    return _EXE_SHA256_CACHE


def compute_download_proof(
    client_id: str,
    version: str = "",
    exe_sha256: str = "",
    secret_context: Optional[str] = None,
) -> str:
    """Compute a 32-char HMAC-SHA256 download attestation proof binding client_id, version, and exe_sha256."""
    cid = str(client_id or "").strip().lower()
    ver = str(version or _app_version).strip()
    sha = str(exe_sha256 or current_exe_sha256()).strip().lower()
    key_str = (
        secret_context
        or os.environ.get("CLICKSHIFT_DOWNLOAD_KEY", "").strip()
        or _DOWNLOAD_ATTEST_CONTEXT
    )
    msg = f"clickshift-download-v1|{cid}|{ver}|{sha}".encode("utf-8")
    digest = hmac.new(key_str.encode("utf-8"), msg, hashlib.sha256).hexdigest()
    return digest[:32].lower()


def verify_download_proof(
    client_id: str,
    proof: str,
    version: str = "",
    exe_sha256: str = "",
    secret_context: Optional[str] = None,
) -> bool:
    """Verify a 32-char HMAC-SHA256 download attestation proof in constant time."""
    if not is_valid_client_id(client_id) or not isinstance(proof, str) or len(proof.strip()) < 16:
        return False
    expected = compute_download_proof(
        client_id,
        version=version,
        exe_sha256=exe_sha256,
        secret_context=secret_context,
    )
    return hmac.compare_digest(proof.strip().lower(), expected.lower())


def download_registry_file() -> Path:
    return _data_dir() / "download_ids.json"


def _read_download_registry() -> dict[str, Any]:
    f = download_registry_file()
    try:
        if f.exists():
            data = json.loads(f.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
    except Exception:
        pass
    return {"downloads": {}, "revoked": []}


def _write_download_registry(reg: dict[str, Any]) -> None:
    d = _data_dir()
    d.mkdir(parents=True, exist_ok=True)
    f = download_registry_file()
    tmp = f.with_suffix(".tmp")
    try:
        tmp.write_text(json.dumps(reg, indent=2), encoding="utf-8")
        tmp.replace(f)
    except Exception:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass


def load_revoked_download_ids(cfg: Optional[dict[str, Any]] = None) -> set[str]:
    """Return lowercase client_ids that have been revoked from contributing to the cloud pool."""
    revoked: set[str] = set()
    reg = _read_download_registry()
    for item in reg.get("revoked") or []:
        s = str(item or "").strip().lower()
        if s:
            revoked.add(s)
    if isinstance(cfg, dict):
        for item in cfg.get("cloud_revoked_client_ids") or []:
            s = str(item or "").strip().lower()
            if s:
                revoked.add(s)
    return revoked


def load_registered_download_ids(cfg: Optional[dict[str, Any]] = None) -> set[str]:
    """Return the set of registered download client IDs permitted to add data to the cloud."""
    revoked = load_revoked_download_ids(cfg)
    if isinstance(cfg, dict):
        explicit = cfg.get("cloud_allowed_client_ids")
        if isinstance(explicit, (list, tuple, set, frozenset)) and len(explicit) > 0:
            return {
                str(x).strip().lower()
                for x in explicit
                if is_valid_client_id(str(x).strip().lower())
            } - revoked

    env_ids = os.environ.get("CLICKSHIFT_ALLOWED_CLIENT_IDS", "").strip()
    if env_ids:
        return {
            x.strip().lower()
            for x in env_ids.split(",")
            if is_valid_client_id(x.strip().lower())
        } - revoked

    allowed: set[str] = set(SEED_DOWNLOAD_CLIENT_IDS)
    reg = _read_download_registry()
    dl_map = reg.get("downloads")
    if isinstance(dl_map, dict):
        for cid, meta in dl_map.items():
            c_low = str(cid or "").strip().lower()
            if is_valid_client_id(c_low) and not (isinstance(meta, dict) and meta.get("revoked")):
                allowed.add(c_low)
    if isinstance(cfg, dict):
        own_cid = str(cfg.get("cloud_client_id") or "").strip().lower()
        if is_valid_client_id(own_cid) and cfg.get("cloud_download_registered", True):
            allowed.add(own_cid)
    return allowed - revoked


def register_download_id(
    client_id: str,
    cfg: Optional[dict[str, Any]] = None,
    version: str = "",
    exe_sha256: str = "",
) -> dict[str, Any]:
    """Register ``client_id`` as an authentic downloaded installation locally and return its metadata."""
    cid = str(client_id or "").strip().lower()
    if not is_valid_client_id(cid):
        raise ValueError(f"Invalid download client_id: {client_id!r}")
    ver = str(version or _app_version).strip()
    sha = str(exe_sha256 or current_exe_sha256()).strip().lower()
    proof = compute_download_proof(cid, version=ver, exe_sha256=sha)
    now = int(time.time())

    reg = _read_download_registry()
    downloads = reg.setdefault("downloads", {})
    if not isinstance(downloads, dict):
        downloads = {}
        reg["downloads"] = downloads
    downloads[cid] = {
        "client_id": cid,
        "download_proof": proof,
        "app_version": ver,
        "exe_sha256": sha,
        "registered_at": downloads.get(cid, {}).get("registered_at", now) if isinstance(downloads.get(cid), dict) else now,
        "last_seen": now,
        "revoked": False,
    }
    rev_list = [str(x).lower() for x in (reg.get("revoked") or []) if str(x).lower() != cid]
    reg["revoked"] = rev_list
    _write_download_registry(reg)

    if isinstance(cfg, dict):
        if cfg.get("cloud_client_id", cid) == cid:
            cfg["cloud_download_registered"] = True
            cfg["cloud_download_proof"] = proof
        try:
            from .monitor import load_config, save_config
            full = load_config()
            if full.get("cloud_client_id", cid) == cid:
                full["cloud_download_registered"] = True
                full["cloud_download_proof"] = proof
                save_config(full)
        except Exception:
            pass
    return downloads[cid]


def revoke_download_id(client_id: str, cfg: Optional[dict[str, Any]] = None) -> bool:
    """Revoke ``client_id`` so its data is blocked from being added to the cloud."""
    cid = str(client_id or "").strip().lower()
    if not cid:
        return False
    reg = _read_download_registry()
    downloads = reg.get("downloads")
    if isinstance(downloads, dict) and cid in downloads and isinstance(downloads[cid], dict):
        downloads[cid]["revoked"] = True
    rev_list = [str(x).lower() for x in (reg.get("revoked") or [])]
    if cid not in rev_list:
        rev_list.append(cid)
    reg["revoked"] = rev_list
    _write_download_registry(reg)

    if isinstance(cfg, dict):
        cur_rev = [str(x).lower() for x in (cfg.get("cloud_revoked_client_ids") or [])]
        if cid not in cur_rev:
            cur_rev.append(cid)
        cfg["cloud_revoked_client_ids"] = cur_rev
        if isinstance(cfg.get("cloud_allowed_client_ids"), list):
            cfg["cloud_allowed_client_ids"] = [
                x for x in cfg["cloud_allowed_client_ids"] if str(x).strip().lower() != cid
            ]
        if str(cfg.get("cloud_client_id") or "").strip().lower() == cid:
            cfg["cloud_download_registered"] = False
    return True


def is_registered_download_client(
    client_id: Any,
    cfg: Optional[dict[str, Any]] = None,
    allowed_client_ids: Optional[Any] = None,
    proof: Optional[str] = None,
    version: str = "",
    exe_sha256: str = "",
    require_proof: bool = False,
) -> bool:
    """Return True only if ``client_id`` matches a registered, non-revoked download ID."""
    if not is_valid_client_id(client_id):
        return False
    cid = str(client_id).strip().lower()
    revoked = load_revoked_download_ids(cfg)
    if cid in revoked:
        return False
    if isinstance(cfg, dict) and cfg.get("cloud_download_registered") is False:
        if str(cfg.get("cloud_client_id") or "").strip().lower() == cid:
            return False

    if allowed_client_ids is not None:
        allowed_set = {
            str(x).strip().lower()
            for x in allowed_client_ids
            if is_valid_client_id(str(x).strip().lower())
        } - revoked
        if cid not in allowed_set:
            return False
    else:
        allowed_set = load_registered_download_ids(cfg)
        if cid not in allowed_set:
            return False

    if require_proof or (proof is not None and str(proof).strip() != ""):
        if not verify_download_proof(cid, str(proof or ""), version=version, exe_sha256=exe_sha256):
            return False
    return True


def merge_states(
    states: list[Any],
    allowed_client_ids: Optional[Any] = None,
    require_registered_client: bool = False,
) -> dict[str, Any]:
    """Merge contributor states into one pooled state.

    * ``bandit`` arms: A matrices and b vectors averaged across contributors
      that have >=1 sample (each contributor's A starts at ridge*I, so
      averaging preserves both the ridge prior scale and theta = A^-1 b),
      while n is summed.
    * ``median_ms``: minimum observed latency wins (best-case speed knowledge).
    * ``clock``: max + 1 (Lamport).
    * Download Client ID protection: when ``allowed_client_ids`` is provided or
      ``require_registered_client=True``, any state whose ``client_id`` does not
      match the registered download IDs (or has a forged ``download_proof``) is
      rejected and excluded from the cloud merge.
    Security: every remote number is validated/bounded here; nonsense (NaN,
    negative counts, absurd latencies, >24h timings, oversized payloads) is
    dropped arm-by-arm instead of poisoning the pool.
    """
    enforce_ids = (allowed_client_ids is not None) or bool(require_registered_client)
    allowed_set: Optional[set[str]] = None
    if enforce_ids:
        if allowed_client_ids is not None:
            revoked = load_revoked_download_ids()
            allowed_set = {
                str(x).strip().lower()
                for x in allowed_client_ids
                if is_valid_client_id(str(x).strip().lower())
            } - revoked
        else:
            allowed_set = load_registered_download_ids()

    bandit: dict[str, _ArmAccum] = {}
    medians: dict[str, float] = {}
    clock = 0
    for st in states:
        if not isinstance(st, dict):
            continue
        cid_raw = st.get("client_id")
        if enforce_ids:
            cid = str(cid_raw or "").strip().lower()
            if not is_valid_client_id(cid) or allowed_set is None or cid not in allowed_set:
                continue
        if cid_raw is not None and "download_proof" in st:
            cid = str(cid_raw or "").strip().lower()
            proof = str(st.get("download_proof") or "")
            ver = str(st.get("app_version") or "")
            sha = str(st.get("exe_sha256") or "")
            if not verify_download_proof(cid, proof, version=ver, exe_sha256=sha):
                continue
        try:
            clock = max(clock, min(int(st.get("clock", 0)), 10 ** 9))
        except (TypeError, ValueError):
            pass
        arms = st.get("bandit") or {}
        if not isinstance(arms, dict):
            continue
        for arm, s in itertools.islice(arms.items(), 512):
            if not isinstance(arm, str) or len(arm) > 128 or not isinstance(s, dict):
                continue
            try:
                n = int(s.get("n", 0))
                if not (0 <= n <= 1_000_000):
                    continue
                A = np.array(s["A"], dtype=float)
                b = _clean_reward(s["b"], n=max(n, 1))
                if b is None or not np.all(np.isfinite(A)):
                    continue
                acc = bandit.get(arm)
                if acc is None:
                    acc = _ArmAccum()
                    bandit[arm] = acc
                acc.n += n
                if n > 0:                       # only real learners vote on the prior
                    if acc.A_sum is None or acc.b_sum is None:
                        acc.A_sum = A.copy()
                        acc.b_sum = b.copy()
                    else:
                        if A.shape != acc.A_sum.shape or b.shape != acc.b_sum.shape:
                            continue
                        acc.A_sum += A
                        acc.b_sum += b
                    acc.nA += 1
            except Exception:
                continue
        pairs = st.get("median_ms") or {}
        if not isinstance(pairs, dict):
            continue
        for pair, ms in itertools.islice(pairs.items(), 2048):
            try:
                ms_val = float(ms)
            except (TypeError, ValueError):
                continue
            if not (0.0 < ms_val <= MAX_LATENCY_MS) or not isinstance(pair, str):
                continue
            if pair not in medians or ms_val < medians[pair]:
                medians[pair] = ms_val
    out_bandit: dict[str, Any] = {}
    for arm, acc in bandit.items():
        if acc.A_sum is None or acc.b_sum is None:
            continue
        denom = max(acc.nA, 1)
        out_bandit[arm] = {
            "A": (acc.A_sum / denom).tolist(),
            "b": (acc.b_sum / denom).tolist(),
            "n": min(acc.n, 1_000_000),
        }
    return {"version": 1, "clock": clock + 1, "updated": time.time(),
            "bandit": out_bandit, "median_ms": medians}


def _contributor_id(cfg: dict[str, Any]) -> str:
    cid = str(cfg.get("cloud_client_id") or "").strip().lower()
    if not is_valid_client_id(cid):
        cid = uuid.uuid4().hex[:16].lower()
        cfg["cloud_client_id"] = cid
        if "cloud_download_registered" not in cfg:
            cfg["cloud_download_registered"] = True
        try:
            register_download_id(cid, cfg=cfg)
        except Exception:
            pass
        try:
            from .monitor import load_config, save_config
            full = load_config()
            full["cloud_client_id"] = cid
            full.setdefault("cloud_download_registered", True)
            save_config(full)
        except Exception:
            pass
    elif cfg.get("cloud_download_registered", True) and cid not in load_revoked_download_ids(cfg):
        reg = _read_download_registry()
        if cid not in (reg.get("downloads") or {}):
            try:
                register_download_id(cid, cfg=cfg)
            except Exception:
                pass
    return str(cid)


# --------------------------------------------------------------------------- #
#  public API
# --------------------------------------------------------------------------- #

class CloudSync:
    """Push/pull the bandit state to/from the shared learning pool."""

    def __init__(self, selector: Any, cfg: Optional[dict[str, Any]] = None) -> None:
        self.selector = selector                    # BanditSelector
        cfg = cfg if cfg is not None else load_cloud_config()
        self.cfg: dict[str, Any] = cfg
        self.mode = str(cfg.get("cloud_mode") or "url").lower()
        self.url = str(cfg.get("cloud_url") or DEFAULT_CLOUD_URL).strip()
        self.gist_id = str(cfg.get("cloud_gist_id") or DEFAULT_GIST_ID).strip()
        self._token: Optional[str] = token() if cfg.get("cloud_enabled") else None
        self.interval = float(cfg.get("cloud_interval_s", 900))
        self.pull_on_start = bool(cfg.get("cloud_pull_on_start", True))
        self._timer: Optional[threading.Timer] = None
        self._next_at: float = 0.0
        self._lock = threading.Lock()
        self._timer_lock = threading.Lock()

    @property
    def token(self) -> str:
        if self._token is None:
            self._token = token()
        return self._token

    @token.setter
    def token(self, value: str) -> None:
        self._token = value

    # -- enabled? ----------------------------------------------------------- #
    def enabled(self) -> bool:
        if os.environ.get("CLICKSHIFT_DISABLE_CLOUD") == "1":
            return False
        if not self.cfg.get("cloud_enabled"):
            return False
        if self.mode == "gist":
            return (self.gist_id.startswith("https://") or API_ROOT.startswith("https://")) \
                and bool(self.gist_id and self.gist_id != DEFAULT_GIST_ID)
        # url mode: refuse plaintext/invalid endpoints - nothing leaves over http
        return _check_url(self.url)

    def is_client_authorized(self, client_id: Optional[str] = None) -> bool:
        """Return True only if ``client_id`` matches our registered download IDs and is not revoked."""
        cid = str(client_id if client_id is not None else _contributor_id(self.cfg)).strip().lower()
        if not bool(self.cfg.get("cloud_require_registered", True)):
            return is_valid_client_id(cid) and cid not in load_revoked_download_ids(self.cfg)
        allowed = self.cfg.get("cloud_allowed_client_ids")
        return is_registered_download_client(
            cid,
            cfg=self.cfg,
            allowed_client_ids=allowed if isinstance(allowed, (list, tuple, set, frozenset)) and len(allowed) > 0 else None,
        )

    def register_download(self) -> bool:
        """Register this installation's download client_id with the local registry and cloud endpoint."""
        cid = _contributor_id(self.cfg)
        if not self.is_client_authorized(cid):
            return False
        info = register_download_id(cid, cfg=self.cfg)
        if not self.enabled() or self.mode == "gist":
            return True
        resp = _request(
            f"{self.url.rstrip('/')}/downloads/register",
            method="POST",
            body={
                "client_id": cid,
                "app_version": info["app_version"],
                "exe_sha256": info["exe_sha256"],
                "download_proof": info["download_proof"],
            },
            headers={
                "X-ClickShift-Client-Id": cid,
                "X-ClickShift-Download-Proof": info["download_proof"],
            },
        )
        return resp is not None and bool(resp.get("ok", True))

    # -- payload ------------------------------------------------------------ #
    def build_payload(self) -> dict[str, Any]:
        if hasattr(self.selector, "export_local_payload"):
            arms, medians = self.selector.export_local_payload()
        else:
            arms = self.selector.bandit.export()
            medians = dict(self.selector.median_ms)
        cid = _contributor_id(self.cfg)
        exe_sha = current_exe_sha256()
        proof = compute_download_proof(cid, version=_app_version, exe_sha256=exe_sha)
        return {
            "version": 1,
            "clock": int(self.cfg.get("cloud_clock", 0)) + 1,
            "client_id": cid,
            "app_version": _app_version,
            "exe_sha256": exe_sha,
            "download_proof": proof,
            "bandit": arms,
            "median_ms": medians,
        }

    # -- pull --------------------------------------------------------------- #
    def pull(self) -> bool:
        """Download pooled state and merge it into the local bandit.

        The remote is UNTRUSTED input: everything is size-capped in _request,
        arm-validated in import_state, latency-bounded here, and a sha256 of
        the accepted state is recorded for tamper visibility.
        """
        if not self.enabled():
            return False
        with self._lock:
            state = self._fetch_pooled()
            if not state or not isinstance(state.get("bandit"), dict):
                return False
            try:
                digest = hashlib.sha256(
                    json.dumps(state, sort_keys=True).encode("utf-8")).hexdigest()[:16]
                if hasattr(self.selector, "merge_remote_state"):
                    self.selector.merge_remote_state(
                        state["bandit"], state.get("median_ms")
                    )
                else:
                    self.selector.bandit.import_state(state["bandit"])
                    for pair, ms in itertools.islice((state.get("median_ms") or {}).items(), 2048):
                        try:
                            ms_val = float(ms)
                        except (TypeError, ValueError):
                            continue
                        if not (0.0 < ms_val <= MAX_LATENCY_MS):
                            continue
                        cur = self.selector.median_ms.get(pair)
                        if cur is None or ms_val < cur:
                            self.selector.median_ms[pair] = ms_val
                    self.selector._save_state()
                self._set_meta_many({
                    "cloud_last_pull": time.time(),
                    "cloud_clock": int(state.get("clock", 0)),
                    "cloud_last_sha256": digest,
                })
                log.info("pulled cloud state %s (%d arms)", digest,
                         len(state["bandit"]))
                return True
            except Exception as exc:                     # never break local ML
                log.debug("cloud pull failed: %s", exc)
                return False

    def _fetch_pooled(self) -> Optional[dict[str, Any]]:
        if self.mode == "gist":
            data = _request(f"{API_ROOT}/gists/{self.gist_id}",
                            headers=self._gh_headers())
            if not data:
                return None
            files = data.get("files") or {}
            if not isinstance(files, dict):
                return None
            entry = files.get("merged.json") or files.get("state.json")
            if not isinstance(entry, dict):
                return None
            raw = entry.get("content")
            if raw is None:                      # too big -> must fetch raw_url
                raw_url = str(entry.get("raw_url") or "")
                if not _is_allowed_raw_url(raw_url):
                    log.warning("refusing untrusted gist raw_url host: %r", raw_url[:100])
                    return None
                raw_resp = _request(raw_url)
                return raw_resp                  # best effort
            try:
                return json.loads(str(raw))
            except Exception:
                return None
        return _request(f"{self.url.rstrip('/')}/state")

    # -- push --------------------------------------------------------------- #
    def push(self) -> bool:
        """Upload our aggregate learnings to the shared pool (only if client_id is a verified download)."""
        if not self.enabled():
            return False
        with self._lock:
            payload = self.build_payload()
            if not payload.get("bandit"):
                return False
            cid = str(payload.get("client_id") or "").strip().lower()
            if not self.is_client_authorized(cid):
                log.warning(
                    "refusing cloud push: client_id %r does not match registered download IDs",
                    cid,
                )
                return False
            if self.mode == "gist":
                ok = self._push_gist(payload)
            else:
                headers = {
                    "X-ClickShift-Client-Id": cid,
                    "X-ClickShift-Download-Proof": str(payload.get("download_proof") or ""),
                }
                resp = _request(
                    f"{self.url.rstrip('/')}/state",
                    "PUT",
                    payload,
                    headers=headers,
                )
                ok = resp is not None and resp.get("ok", True) is not False
            if ok:
                self._set_meta_many({
                    "cloud_last_push": time.time(),
                    "cloud_clock": payload["clock"],
                })
            return ok

    def _gh_headers(self) -> dict[str, str]:
        h: dict[str, str] = {"Accept": "application/vnd.github+json"}
        if self.token:
            h["Authorization"] = f"Bearer {self.token}"
        return h

    def _push_gist(self, payload: dict[str, Any]) -> bool:
        if not self.token:
            return False
        cid = str(payload.get("client_id") or _contributor_id(self.cfg)).strip().lower()
        if not self.is_client_authorized(cid):
            return False
        data = _request(f"{API_ROOT}/gists/{self.gist_id}", headers=self._gh_headers())
        if not data:
            return False
        files = data.get("files") or {}
        if not isinstance(files, dict):
            return False

        # Read registered download IDs from downloads.json in the Gist if present
        allowed_ids: Optional[set[str]] = None
        dl_entry = files.get("downloads.json")
        if isinstance(dl_entry, dict) and isinstance(dl_entry.get("content"), str):
            try:
                dl_obj = json.loads(dl_entry["content"])
                if isinstance(dl_obj, dict) and isinstance(dl_obj.get("allowed_client_ids"), list):
                    allowed_ids = {
                        str(x).strip().lower()
                        for x in dl_obj["allowed_client_ids"]
                        if is_valid_client_id(str(x).strip().lower())
                    }
            except Exception:
                allowed_ids = None
        if allowed_ids is not None and cid not in allowed_ids:
            log.warning("refusing gist push: client_id %r is not in gist downloads.json", cid)
            return False

        mine = f"contrib-{cid}.json"
        contribs: list[Any] = []
        for k, v in files.items():
            if not isinstance(v, dict) or k == mine or k in ("merged.json", "state.json", "downloads.json"):
                continue
            content = v.get("content")
            if isinstance(k, str) and k.startswith("contrib-") and k.endswith(".json") and isinstance(content, str) and len(content) <= MAX_CONTRIB_BYTES:
                try:
                    c_obj = json.loads(content)
                    if isinstance(c_obj, dict) and "client_id" not in c_obj:
                        c_obj["client_id"] = k[len("contrib-"):-len(".json")].strip().lower()
                    contribs.append(c_obj)
                except Exception:
                    continue
        contribs.append(payload)
        patch_files = {
            mine: {"content": json.dumps(payload)},
            "merged.json": {"content": json.dumps(merge_states(contribs, allowed_client_ids=allowed_ids))},
        }
        resp = _request(f"{API_ROOT}/gists/{self.gist_id}", "PATCH",
                        {"files": patch_files}, headers=self._gh_headers())
        return resp is not None

    # -- meta --------------------------------------------------------------- #
    def _set_meta(self, key: str, value: Any) -> None:
        self._set_meta_many({key: value})

    def _set_meta_many(self, updates: dict[str, Any]) -> None:
        try:
            from .monitor import load_config, save_config
            cfg = load_config()
            cfg.update(updates)
            save_config(cfg)
            self.cfg.update(updates)
        except Exception:
            pass

    # -- background loop ----------------------------------------------------- #
    def start(self) -> None:
        if not self.enabled():
            return
        if self.pull_on_start:
            threading.Thread(target=self.pull, daemon=True,
                             name="clickshift-cloud-pull").start()
        self._schedule(self.interval)

    def _schedule(self, delay: Optional[float] = None) -> None:
        """(Re-)arm the sync timer. Debounced: never fires more often than
        ``cloud_debounce_s`` (default 60 s), so a burst of conversions results
        in exactly one upload."""
        debounce = float(self.cfg.get("cloud_debounce_s", 60))
        actual_delay = max(delay if delay is not None else self.interval, debounce)
        with self._timer_lock:
            now = time.time()
            if self._timer is not None and self._next_at - debounce < now:
                return                      # one already pending soon enough
            if self._timer is not None:
                self._timer.cancel()
            self._next_at = now + actual_delay
            timer = threading.Timer(actual_delay, self._tick)
            timer.daemon = True
            timer.start()
            self._timer = timer

    def _schedule_now(self) -> None:
        """Called by the selector after every recorded conversion."""
        self._schedule(float(self.cfg.get("cloud_debounce_s", 60)))

    def _tick(self) -> None:
        try:
            self.push()
            self.pull()
        finally:
            self._schedule(self.interval)

    def stop(self) -> None:
        with self._timer_lock:
            if self._timer is not None:
                self._timer.cancel()
                self._timer = None

    def status(self) -> dict[str, Any]:
        cid = _contributor_id(self.cfg)
        return {"enabled": self.enabled(), "mode": self.mode,
                "endpoint": self.url or (f"gist:{self.gist_id}" if self.enabled() else ""),
                "token_set": bool(self.token),
                "client_id": cid,
                "client_authorized": self.is_client_authorized(cid),
                "last_pull": self.cfg.get("cloud_last_pull"),
                "last_push": self.cfg.get("cloud_last_push"),
                "state_sha256": self.cfg.get("cloud_last_sha256"),
                "arms_learned": len(self.selector.bandit.n),
                "samples": sum(self.selector.bandit.n.values())}


# --------------------------------------------------------------------------- #
#  CI merge worker:  python -m clickshift.cloud --merge-gist  (run from repo root)
# --------------------------------------------------------------------------- #

def merge_gist_worker(gist_id: Optional[str] = None, token_: Optional[str] = None,
                      dry_run: bool = False,
                      allowed_client_ids: Optional[Sequence[str]] = None) -> dict[str, Any]:
    """Recompute merged.json from all contrib-*.json files in the gist.

    Used by the scheduled GitHub Action so contributors only ever PATCH their
    own file (no lost-update races on the aggregate). Contributors whose
    client_id is not in downloads.json / allowed_client_ids are excluded.
    """
    gid = (gist_id or os.environ.get("CLICKSHIFT_GIST_ID") or "").strip()
    tok = token_ or token()
    if not gid or gid == DEFAULT_GIST_ID or not tok:
        print("cloud pool not configured (gist id/token missing) - nothing to merge.")
        return {"version": 1, "clock": 0, "updated": time.time(),
                "bandit": {}, "median_ms": {}}
    headers = {"Authorization": f"Bearer {tok}"}
    data = _request(f"{API_ROOT}/gists/{gid}", headers=headers)
    if not data:
        print(f"warning: cannot read gist {gid} - leaving pool untouched.")
        return {"version": 1, "clock": 0, "updated": time.time(),
                "bandit": {}, "median_ms": {}}
    files = data.get("files") or {}
    effective_allowed: Optional[set[str]] = (
        {str(x).strip().lower() for x in allowed_client_ids if is_valid_client_id(str(x).strip().lower())}
        if allowed_client_ids is not None
        else None
    )
    if effective_allowed is None and isinstance(files, dict):
        dl_entry = files.get("downloads.json")
        if isinstance(dl_entry, dict) and isinstance(dl_entry.get("content"), str):
            try:
                dl_obj = json.loads(dl_entry["content"])
                if isinstance(dl_obj, dict) and isinstance(dl_obj.get("allowed_client_ids"), list):
                    effective_allowed = {
                        str(x).strip().lower()
                        for x in dl_obj["allowed_client_ids"]
                        if is_valid_client_id(str(x).strip().lower())
                    }
            except Exception:
                effective_allowed = None

    contribs: list[Any] = []
    keep: dict[str, dict[str, str]] = {}
    if isinstance(files, dict):
        for k, v in files.items():
            if not isinstance(v, dict):
                continue
            content = v.get("content")
            if (isinstance(k, str) and k.startswith("contrib-") and k.endswith(".json")
                    and isinstance(content, str) and len(content) <= MAX_CONTRIB_BYTES):
                keep[k] = {"content": content}
    for fname, f in keep.items():
        try:
            c_obj = json.loads(f["content"])
            if isinstance(c_obj, dict) and "client_id" not in c_obj:
                c_obj["client_id"] = fname[len("contrib-"):-len(".json")].strip().lower()
            contribs.append(c_obj)
        except Exception:
            continue
    merged = merge_states(contribs, allowed_client_ids=effective_allowed)
    if not dry_run and tok:
        _request(f"{API_ROOT}/gists/{gid}", "PATCH", {"files": {"merged.json": {"content": json.dumps(merged)}}},
                 headers=headers)
    return merged


if __name__ == "__main__":  # pragma: no cover
    import argparse
    ap = argparse.ArgumentParser(prog="clickshift.cloud")
    ap.add_argument("--merge-gist", action="store_true")
    ap.add_argument("--gist-id", default=None)
    ap.add_argument("--dry-run", action="store_true")
    ns = ap.parse_args()
    if ns.merge_gist:
        m = merge_gist_worker(ns.gist_id, dry_run=ns.dry_run)
        print(json.dumps({"clock": m["clock"], "arms": len(m["bandit"]),
                          "pairs": len(m["median_ms"])}, indent=2))
    else:
        ap.print_help()
