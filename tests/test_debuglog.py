"""Tests for the full debugging layer (debuglog)."""
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pytest  # noqa: E402


@pytest.fixture()
def home(tmp_path, monkeypatch):
    h = tmp_path / ".clickshift"
    monkeypatch.setenv("CLICKSHIFT_HOME", str(h))
    # reload monitor so DATA_DIR picks up env var
    for m in ("clickshift.monitor", "clickshift.engine", "clickshift.debuglog"):
        sys.modules.pop(m, None)
    from clickshift import monitor, debuglog
    return h, monitor, debuglog


def test_default_produces_no_debug_txt(home):
    h, monitor, debuglog = home
    src = Path(monitor.__file__).parent.parent.parent / "README.md"
    assert src.exists()
    from clickshift.engine import ConversionEngine
    eng = ConversionEngine(enable_ml=False)
    out_dir = h.parent / "out_default"
    out_dir.mkdir(parents=True, exist_ok=True)
    res = eng.convert(src, "txt", out_dir=out_dir)
    assert res.ok, res.error
    assert not (Path(res.target + ".clickshift-debug.txt")).exists()
    assert not (h / "debug").exists() or not list((h / "debug").glob("*.txt"))


def test_success_writes_central_report_only_when_enabled(home, monkeypatch):
    h, monitor, debuglog = home
    monkeypatch.setenv("CLICKSHIFT_DEBUG_REPORTS", "always")
    src = Path(monitor.__file__).parent.parent.parent / "README.md"
    assert src.exists()
    from clickshift.engine import ConversionEngine
    eng = ConversionEngine(enable_ml=False)
    out_dir = h.parent / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    res = eng.convert(src, "txt", out_dir=out_dir)
    assert res.ok, res.error
    rep = Path(res.target + ".clickshift-debug.txt")
    assert not rep.exists(), "no debug txt should be dropped next to converted user file"
    central = list((h / "debug").glob("*-OK.clickshift-debug.txt"))
    assert central, "central copy stored under ~/.clickshift/debug/ when explicitly enabled"
    text = central[-1].read_text(encoding="utf-8")
    assert "ClickShift CONVERSION DEBUG REPORT" in text
    assert "SUCCESS" in text
    assert "What went RIGHT" in text
    assert "[CANDIDATE CHAIN & ATTEMPTS]" in text
    assert "+ SUCCESS" in text


def test_failure_report_shows_traceback(home, monkeypatch):
    h, monitor, debuglog = home
    monkeypatch.setenv("CLICKSHIFT_DEBUG_REPORTS", "onerror")
    # corrupt pdf -> all backends should fail (no libreoffice/office here)
    bad = h.parent / "broken.pdf"
    bad.write_bytes(b"%PDF-1.4 not a real pdf at all \x00\x01garbage")
    from clickshift.engine import ConversionEngine
    eng = ConversionEngine(enable_ml=False)
    res = eng.convert(bad, "docx")
    if res.ok:  # environment with real backends may succeed; skip then
        pytest.skip("a backend succeeded on this platform")
    reps = list((h / "debug").glob("*-FAIL.clickshift-debug.txt"))
    assert reps, "failure must still produce a report"
    text = reps[-1].read_text(encoding="utf-8")
    assert "FAILURE" in text
    assert "What went WRONG" in text
    assert "error:" in text or "FAILED after" in text


def test_mode_off_suppresses(home, monkeypatch):
    h, monitor, debuglog = home
    monkeypatch.setenv("CLICKSHIFT_DEBUG_REPORTS", "off")
    cfgdir = h
    cfgdir.mkdir(parents=True, exist_ok=True)
    from clickshift.engine import ConversionEngine
    eng = ConversionEngine(enable_ml=False)
    src = Path(monitor.__file__).parent.parent.parent / "README.md"
    res = eng.convert(src, "txt", out_dir=h.parent / "o2")
    assert res.ok
    assert not (Path(res.target).with_name(Path(res.target).name + ".clickshift-debug.txt")).exists()
    assert not (h / "debug").exists() or not list((h / "debug").glob("*.txt"))


def test_mode_onerror_only_failures(home, monkeypatch):
    h, monitor, debuglog = home
    monkeypatch.setenv("CLICKSHIFT_DEBUG_REPORTS", "onerror")
    from clickshift.engine import ConversionEngine
    eng = ConversionEngine(enable_ml=False)
    src = Path(monitor.__file__).parent.parent.parent / "README.md"
    res = eng.convert(src, "txt", out_dir=h.parent / "o3")
    assert res.ok
    assert not list((h / "debug").glob("*.txt")) if (h / "debug").exists() else True


def test_report_never_breaks_conversion(home, monkeypatch):
    """Even if rendering/writing explodes, conversion still succeeds."""
    h, monitor, debuglog = home
    monkeypatch.setattr(debuglog, "render_report",
                        lambda s: (_ for _ in ()).throw(RuntimeError("boom")))
    from clickshift.engine import ConversionEngine
    eng = ConversionEngine(enable_ml=False)
    src = Path(monitor.__file__).parent.parent.parent / "README.md"
    res = eng.convert(src, "txt", out_dir=h.parent / "o4")
    assert res.ok, "conversion must be unaffected by broken debug writer"


def test_ml_scores_recorded(home, monkeypatch):
    h, monitor, debuglog = home
    monkeypatch.setenv("CLICKSHIFT_DEBUG_REPORTS", "always")
    monkeypatch.setenv("CLICKSHIFT_DISABLE_CLOUD", "1")
    from clickshift.engine import ConversionEngine
    eng = ConversionEngine(enable_ml=True, history_file=h / "history.jsonl")
    src = Path(monitor.__file__).parent.parent.parent / "README.md"
    res = eng.convert(src, "txt", out_dir=h.parent / "o5")
    assert res.ok
    central = list((h / "debug").glob("*-OK.clickshift-debug.txt"))
    assert central
    text = central[-1].read_text(encoding="utf-8")
    assert "[ML BACKEND SELECTION]" in text
    assert "enabled        : True" in text


def test_retention_pruning(home, monkeypatch):
    h, monitor, debuglog = home
    monkeypatch.setenv("CLICKSHIFT_DEBUG_REPORTS", "always")
    dd = h / "debug"
    dd.mkdir(parents=True, exist_ok=True)
    for i in range(25):
        f = dd / f"old{i:02d}.clickshift-debug.txt"
        f.write_text("x")
        os.utime(f, (i, i))  # distinct old mtimes
    monkeypatch.setattr(debuglog, "keep_count", lambda: 10)
    sess = {
        "run_id": "t", "t_start": __import__("time").time(), "src": "a.txt",
        "dst_ext": ".docx", "planned_out": "", "out_path": "",
        "backends": {}, "ml": {"enabled": False}, "routes": [], "attempts": [],
        "ok": True, "winning_backend": "x", "total_ms": 1.0,
        "data_dir": str(h),
    }
    debuglog.write_report(sess)
    remaining = list(dd.glob("*.clickshift-debug.txt"))
    assert len(remaining) <= 11  # 10 kept + fresh one


def test_lazy_backend_snapshot_and_exc_traceback(tmp_path, monkeypatch):
    from clickshift import debuglog
    import clickshift.engine as eng_mod

    monkeypatch.setenv("CLICKSHIFT_HOME", str(tmp_path))
    monkeypatch.setenv("CLICKSHIFT_DEBUG_REPORTS", "off")

    called = False

    def fail_summary():
        nonlocal called
        called = True
        return {}

    monkeypatch.setattr(eng_mod.BACKENDS, "summary", fail_summary)
    sess = debuglog.new_session("foo.txt", ".html")
    assert sess["backends"] is None
    assert called is False

    try:
        raise ValueError("boom trace")
    except ValueError as err:
        rep = debuglog.render_report({
            "run_id": "r1", "src": "foo.txt", "dst_ext": ".html",
            "planned_out": "foo.html", "out_path": "foo.html",
            "backends": None, "ml": {"enabled": False},
            "routes": [{"name": "x", "weight": 50}],
            "attempts": [{"backend": "x", "ok": False, "ms": 2.0, "error": "boom", "exc": err}],
            "ok": False, "winning_backend": "", "total_ms": 2.0,
            "data_dir": str(tmp_path),
        })
    assert called is True
    assert "ValueError: boom trace" in rep

