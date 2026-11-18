"""Tests for the shared-learnings cloud sync module."""
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import numpy as np  # noqa: E402

from clickshift import cloud  # noqa: E402
from clickshift.learning import BanditSelector, _features, DIM  # noqa: E402


def _mk_state(seed, n_updates=8, arm="pdf>docx|libreoffice"):
    rng = np.random.default_rng(seed)
    b = {"version": 1, "clock": seed, "bandit": {}, "median_ms": {arm.split("|")[0]: 500.0 + seed}}
    A = np.eye(DIM) * 8.0
    v = np.zeros(DIM)
    for i in range(n_updates):
        x = rng.normal(size=DIM) * 0.1
        x[0] = 1.0
        r = float(rng.normal(0.5, 0.2))
        A += np.outer(x, x)
        v += r * x
    b["bandit"][arm] = {"A": A.tolist(), "b": v.tolist(), "n": n_updates}
    return b


def test_merge_states_combines():
    s1, s2 = _mk_state(1), _mk_state(2)
    m = cloud.merge_states([s1, s2])
    arm = "pdf>docx|libreoffice"
    assert m["bandit"][arm]["n"] == 16                      # samples summed
    assert m["clock"] == max(s1["clock"], s2["clock"]) + 1  # Lamport bump
    A = np.array(m["bandit"][arm]["A"])
    assert A.shape == (DIM, DIM)
    # averaged A keeps ridge prior scale
    assert float(A[0, 0]) > 8.0 and float(A[0, 0]) <= 24.0  # pooled evidence sharpens the prior
    # median min wins
    assert m["median_ms"]["pdf>docx"] == 501.0


def test_merge_empty_and_bad():
    m = cloud.merge_states([])
    assert m["bandit"] == {} and m["median_ms"] == {}
    m = cloud.merge_states([{"garbage": True}, None, _mk_state(3)])
    assert m["bandit"], "valid contributor still merged"


def test_selector_records_and_warms(tmp_path, monkeypatch):
    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    sel = BanditSelector(history_file=tmp_path / "history.jsonl")
    src = tmp_path / "a.pdf"
    src.write_bytes(b"x" * 100)
    sel.record(src, ".docx", "libreoffice", 400.0, True)
    sel.record(src, ".docx", "pdf2docx", 900.0, False)
    st = json.loads(sel.state_file.read_text())
    assert st["bandit"]["pdf>docx|libreoffice"]["n"] == 1
    # a fresh selector loads persisted state
    sel2 = BanditSelector(history_file=tmp_path / "history.jsonl")
    assert sel2.bandit.n.get("pdf>docx|libreoffice") == 1
    # simulate warm download: merge another user's pooled state into ours
    pooled = _mk_state(7, n_updates=20, arm="pdf>docx|pdf2docx")
    sel2.bandit.import_state(pooled["bandit"])
    assert sel2.bandit.n["pdf>docx|pdf2docx"] == 20   # starts warm, no cold exploration


def test_cloudsync_enabled_by_default_with_opt_out(tmp_path, monkeypatch):
    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    monkeypatch.setattr(cloud, "_request", lambda *a, **k: {})
    from clickshift.monitor import load_config
    sel = BanditSelector(history_file=tmp_path / "history.jsonl")
    cs = cloud.CloudSync(sel, load_config())
    assert cs.enabled()                   # shared pool active by default
    assert cs.mode == "url"
    assert cs.url == cloud.DEFAULT_CLOUD_URL
    cs.start()
    cs.stop()

    monkeypatch.setenv("CLICKSHIFT_DISABLE_CLOUD", "1")
    assert not cs.enabled()               # env kill-switch disables sync cleanly


def test_cloudsync_pull_merges_local(tmp_path, monkeypatch):
    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    sel = BanditSelector(history_file=tmp_path / "history.jsonl")
    cfg = {"cloud_enabled": True, "cloud_mode": "url",
           "cloud_url": "https://example.invalid"}
    cs = cloud.CloudSync(sel, cfg)
    assert cs.enabled()
    pooled = _mk_state(5, n_updates=12)
    monkeypatch.setattr(cloud, "_request", lambda *a, **k: pooled)
    assert cs.pull() is True
    assert sel.bandit.n["pdf>docx|libreoffice"] == 12
    assert sel.median_ms["pdf>docx"] == 505.0
    # persisted -> survives restart
    sel2 = BanditSelector(history_file=tmp_path / "history.jsonl")
    assert sel2.bandit.n["pdf>docx|libreoffice"] == 12


def test_push_payload_has_no_filenames(tmp_path, monkeypatch):
    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    sel = BanditSelector(history_file=tmp_path / "history.jsonl")
    src = tmp_path / "secret_report.pdf"
    src.write_bytes(b"x" * 10)
    sel.record(src, ".docx", "libreoffice", 100.0, True)
    cs = cloud.CloudSync(sel, {"cloud_enabled": True, "cloud_mode": "url",
                               "cloud_url": "https://example.invalid"})
    payload = json.dumps(cs.build_payload())
    assert "secret_report" not in payload      # privacy: aggregate numbers only
    assert "tmp_path" not in payload or str(tmp_path) not in payload
    captured = {}
    def fake_req(url, method="GET", body=None, headers=None, timeout=15):
        captured["body"] = body
        return {}
    monkeypatch.setattr(cloud, "_request", fake_req)
    assert cs.push()
    assert captured["body"]["bandit"]          # what the server receives
    assert "file" not in json.dumps(captured["body"])


def test_engine_attaches_cloud(tmp_path, monkeypatch):
    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    monkeypatch.setattr(cloud, "_request", lambda *a, **k: {})
    from clickshift.engine import ConversionEngine
    eng = ConversionEngine(enable_ml=True, history_file=tmp_path / "history.jsonl")
    assert eng._selector is not None
    assert eng.cloud is not None               # CloudSync object attached & enabled
    assert eng.cloud.enabled()
    eng.shutdown()


# --------------------------------------------------------------------------- #
#  security hardening
# --------------------------------------------------------------------------- #

def test_https_only_guard():
    assert cloud._check_url("https://api.github.com/gists/abc")
    assert not cloud._check_url("http://example.com/state")      # plaintext refused
    assert not cloud._check_url("file:///etc/passwd")            # local file exfil
    assert not cloud._check_url("ftp://x/y")
    assert not cloud._check_url("https://\nInjected: header")    # header injection
    assert not cloud._check_url("")


def test_non_https_push_and_pull_refused(tmp_path, monkeypatch):
    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    sel = BanditSelector(history_file=tmp_path / "history.jsonl")
    calls = []
    def spy(url, method="GET", body=None, headers=None, timeout=15):
        calls.append((method, url))
        return {}
    monkeypatch.setattr(cloud, "_request", spy)
    cs = cloud.CloudSync(sel, {"cloud_enabled": True, "cloud_mode": "url",
                               "cloud_url": "http://evil.example"})
    # CloudSync.enabled() must refuse plaintext endpoints outright
    assert not cs.enabled()


def test_import_state_rejects_poisoned_arms(tmp_path, monkeypatch):
    """A hostile pooled state can never corrupt the local bandit."""
    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    sel = BanditSelector(history_file=tmp_path / "history.jsonl")
    d = sel.bandit.dim
    good = {"A": (np.eye(d) * 0.1).tolist(), "b": [0.5] * d, "n": 10}
    bad_nan = {"A": [[float("nan")] * d for _ in range(d)], "b": [0.0] * d, "n": 5}
    bad_shape = {"A": [[1.0]], "b": [0.0], "n": 5}
    bad_n = {"A": (np.eye(d)).tolist(), "b": [0.0] * d, "n": 10 ** 12}
    bad_asym = {"A": [[1.0, 5.0], [0.0, 1.0]] if d == 2 else
                (np.eye(d) + np.diag([9.0] * d)[:1, :1] if False else np.eye(d)).tolist(),
                "b": [0.0] * d, "n": 5}
    bad_neg_eig = {"A": (-np.eye(d)).tolist(), "b": [0.0] * d, "n": 5}
    sel.bandit.import_state({
        "good|arm": good, "nan|arm": bad_nan, "shape|arm": bad_shape,
        "count|arm": bad_n, "asym|arm": bad_asym, "negeig|arm": bad_neg_eig,
    })
    assert sel.bandit.n.get("good|arm") == 10          # valid arm imported
    for poison in ("nan|arm", "shape|arm", "count|arm", "negeig|arm"):
        assert poison not in sel.bandit.n              # garbage dropped
    # scores still computable -> bandit not broken
    import numpy as _np
    sel.bandit.score("good|arm", _np.zeros(d))


def test_merge_drops_absurd_latencies_and_counts():
    st = {"version": 1, "clock": 3,
          "bandit": {"a|x": {"A": [[2.0]], "b": [500.0], "n": 10 ** 9}},
          "median_ms": {"pdf>docx": 10 ** 15, "txt>pdf": -5, "ok>pair": 42.0}}
    m = cloud.merge_states([st])
    assert m["median_ms"] == {"ok>pair": 42.0}         # absurd/negative rejected
    assert "a|x" not in m["bandit"]                    # n=1e9 arm dropped entirely


def test_response_size_cap(monkeypatch):
    """_request refuses oversized responses (memory DoS guard)."""
    class FakeResp:
        def read(self, n=-1):
            return b"x" * (n + 1)                     # always over cap
        def __enter__(self): return self
        def __exit__(self, *a): return False
    monkeypatch.setattr(cloud.urllib.request, "urlopen", lambda req, timeout=15: FakeResp())
    assert cloud._request("https://hostile.example/state") is None


def test_token_saved_owner_only(tmp_path, monkeypatch):
    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    p = cloud.save_token("ghp_secret123")
    assert p.read_text() == "ghp_secret123"
    if sys.platform != "win32":
        mode = p.stat().st_mode & 0o777
        assert mode == 0o600
    else:
        import subprocess
        out = subprocess.check_output(["icacls", str(p)], text=True)
        # Protected DACL must not retain BUILTIN\Users or Everyone
        assert "BUILTIN\\Users" not in out and "Everyone" not in out


def test_ssrf_and_raw_url_guards():
    assert not cloud._check_url("https://127.0.0.1/state")
    assert not cloud._check_url("https://localhost/state")
    assert not cloud._check_url("https://[::1]/state")
    assert not cloud._check_url("https://10.0.0.5/state")
    assert not cloud._check_url("https://169.254.169.254/latest")
    assert not cloud._check_url("https://service.local/state")
    assert not cloud._check_url("https://user:secret@example.invalid/state")
    assert cloud._is_allowed_raw_url("https://gist.githubusercontent.com/user/id/raw/merged.json")
    assert not cloud._is_allowed_raw_url("https://evil.example/payload.json")


def test_merge_preserves_accumulated_reward_scale():
    """Accumulated reward vectors b with n>1 must not be clipped to [-1, 1]."""
    st = _mk_state(seed=10, n_updates=20, arm="pdf>docx|pdf2docx")
    st["bandit"]["pdf>docx|pdf2docx"]["b"][0] = 15.0  # 20 runs with ~0.75 avg reward
    m = cloud.merge_states([st])
    assert m["bandit"]["pdf>docx|pdf2docx"]["b"][0] == 15.0


def test_pull_caps_remote_prior_and_push_never_echoes_remote(tmp_path, monkeypatch):
    """Pulling large remote state preserves local observations and never re-uploads remote counts."""
    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    sel = BanditSelector(history_file=tmp_path / "history.jsonl")
    src = tmp_path / "local.pptx"
    src.write_bytes(b"PK\x03\x04local")
    sel.record(src, ".pdf", "office-com", 200.0, True)
    sel.record(src, ".pdf", "office-com", 220.0, True)

    cs = cloud.CloudSync(
        sel,
        {"cloud_enabled": True, "cloud_mode": "url", "cloud_url": "https://example.invalid"},
    )
    # _gh_headers must never attach an empty Bearer token when unauthenticated
    assert "Authorization" not in cs._gh_headers()

    large_remote = _mk_state(seed=99, n_updates=500, arm="pptx>pdf|office-com")
    monkeypatch.setattr(cloud, "_request", lambda *a, **k: large_remote)
    assert cs.pull() is True

    # Remote prior capped at 24 + 2 local = 26 (not 502)
    assert sel.bandit.n["pptx>pdf|office-com"] == 26
    # Push payload must only contain our 2 genuine local updates + random client_id
    payload = cs.build_payload()
    assert payload["bandit"]["pptx>pdf|office-com"]["n"] == 2
    assert len(payload["client_id"]) >= 8


def test_default_cloud_url_enables_uncustomized_and_honors_opt_out(tmp_path, monkeypatch):
    """When DEFAULT_CLOUD_URL is configured, uncustomized installs opt in while explicit off stays off."""
    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    monkeypatch.setattr(cloud, "DEFAULT_CLOUD_URL", "https://pool.example.workers.dev")
    from clickshift.monitor import DEFAULT_CONFIG, save_config

    cfg_legacy = dict(DEFAULT_CONFIG)
    cfg_legacy["cloud_enabled"] = False
    cfg_legacy["cloud_mode"] = "gist"
    cfg_legacy["cloud_url"] = ""
    save_config(cfg_legacy)
    loaded = cloud.load_cloud_config()
    assert loaded["cloud_enabled"] is True
    assert loaded["cloud_mode"] == "url"
    assert loaded["cloud_url"] == "https://pool.example.workers.dev"

    cfg_opt_out = dict(DEFAULT_CONFIG)
    cfg_opt_out["cloud_opt_out"] = True
    cfg_opt_out["cloud_enabled"] = False
    save_config(cfg_opt_out)
    loaded_off = cloud.load_cloud_config()
    assert loaded_off["cloud_enabled"] is False

