import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
os.environ.setdefault("CONFAS_HOME", str(Path(__file__).parent / "_tmp_home"))

import numpy as np  # noqa: E402
from confas.learning import BanditSelector, _features, DIM  # noqa: E402
from confas.engine import Route  # noqa: E402


def make_routes():
    return [Route("fast", lambda s, d: None, 90), Route("slow", lambda s, d: None, 60)]


def test_features_dim():
    p = Path(__file__)
    f = _features(p, ".docx")
    assert len(f) == DIM


def test_bandit_learns_faster_backend(tmp_path):
    sel = BanditSelector(history_file=tmp_path / "h.jsonl")
    src = tmp_path / "doc.pdf"
    src.write_bytes(b"%PDF-1.7 fake")
    routes = make_routes()

    # simulate: 'fast' always ~100ms, 'slow' always ~900ms
    for i in range(24):
        ranked = sel.rank(src, ".docx", routes)
        chosen = ranked[0]
        ms = 100 if chosen.name == "fast" else 900
        sel.record(src, ".docx", chosen.name, ms, True)

    final = sel.rank(src, ".docx", routes)
    assert final[0].name == "fast", "bandit should converge to the faster backend"


def test_state_persists(tmp_path):
    sel = BanditSelector(history_file=tmp_path / "h.jsonl")
    src = tmp_path / "d.pdf"
    src.write_bytes(b"x")
    sel.record(src, ".docx", "fast", 120, True)
    sel2 = BanditSelector(history_file=tmp_path / "h.jsonl")
    assert sum(sel2.bandit.n.values()) >= 1


def test_content_aware_features_distinguish_scanned_vs_text_pdf(tmp_path):
    text_pdf = tmp_path / "text.pdf"
    scan_pdf = tmp_path / "scan.pdf"
    text_pdf.write_bytes(b"%PDF-1.7\n/Type /Page\n/Font /F1\n/Font /F2\n" + b"a" * 500)
    scan_pdf.write_bytes(b"%PDF-1.7\n/Type /Page\n/Subtype /Image\n/Subtype /Image\n" + b"a" * 500)
    f_text = _features(text_pdf, ".docx")
    f_scan = _features(scan_pdf, ".docx")
    assert len(f_text) == DIM and len(f_scan) == DIM
    # media_ratio (index 7) is higher for the image-heavy PDF than the font-only PDF
    assert f_scan[7] > f_text[7]


def test_sherman_morrison_matches_exact_inverse(tmp_path):
    sel = BanditSelector(history_file=tmp_path / "h.jsonl")
    rng = np.random.default_rng(42)
    arm = "pdf>docx|pdf2docx"
    for _ in range(15):
        x = rng.normal(size=DIM)
        x[0] = 1.0
        sel.bandit.update(arm, x, float(rng.uniform(-1.0, 1.0)))
    exact_inv = np.linalg.inv(sel.bandit.A[arm])
    assert np.allclose(sel.bandit.A_inv[arm], exact_inv, atol=1e-6)
