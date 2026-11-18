"""Release integrity helper: generate ECDSA P-256 keys, SHA256SUMS.txt, and .sig files.

Usage:
  python scripts/sign_release.py keygen
  python scripts/sign_release.py sums dist/ClickShift-v*.exe --out dist/SHA256SUMS.txt
  python scripts/sign_release.py sign dist/ClickShift-v1.0.19.exe
  python scripts/sign_release.py verify dist/ClickShift-v1.0.19.exe dist/ClickShift-v1.0.19.exe.sig --pubkey <b64>
"""
from __future__ import annotations

import argparse
import base64
import glob
import hashlib
import os
import struct
import sys
from pathlib import Path
from typing import Sequence

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))


def _file_sha256(path: Path) -> bytes:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        while True:
            chunk = fh.read(65536)
            if not chunk:
                break
            h.update(chunk)
    return h.digest()


def _load_ec_module():
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import ec
    from cryptography.hazmat.primitives.asymmetric.utils import Prehashed, decode_dss_signature
    return hashes, serialization, ec, Prehashed, decode_dss_signature


def cmd_keygen() -> int:
    _hashes, serialization, ec, _pre, _dec = _load_ec_module()
    priv = ec.generate_private_key(ec.SECP256R1())
    priv_pem = priv.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption(),
    )
    pub_nums = priv.public_key().public_numbers()
    cng_blob = (
        struct.pack("<II", 0x31534345, 32)
        + pub_nums.x.to_bytes(32, "big")
        + pub_nums.y.to_bytes(32, "big")
    )
    pub_b64 = base64.b64encode(cng_blob).decode("ascii")
    priv_b64 = base64.b64encode(priv_pem).decode("ascii")
    print("=== Embed this PUBLIC key in src/clickshift/updater.py (TRUSTED_ECDSA_PUBKEYS_B64) ===")
    print(pub_b64)
    print("\n=== Store this PRIVATE key ONLY in your password manager & MFA-gated GitHub Secret (CLICKSHIFT_SIGNING_KEY_B64) ===")
    print(priv_b64)
    return 0


def _resolve_input_files(patterns: Sequence[Path | str]) -> list[Path]:
    """Resolve explicit paths and wildcard globs (needed on Windows shells like pwsh)."""
    seen: set[Path] = set()
    resolved: list[Path] = []
    for item in patterns:
        raw = str(item)
        direct = Path(raw)
        if direct.is_file():
            candidates = [direct]
        elif glob.has_magic(raw):
            candidates = [Path(m) for m in sorted(glob.glob(raw, recursive=True))]
        else:
            candidates = []
        for cand in candidates:
            if cand.is_file():
                key = cand.resolve()
                if key not in seen:
                    seen.add(key)
                    resolved.append(cand)
    return resolved


def cmd_sums(files: Sequence[Path | str], out: Path) -> int:
    resolved = _resolve_input_files(files)
    lines: list[str] = []
    for f in sorted(resolved, key=lambda p: p.name.lower()):
        digest = _file_sha256(f).hex()
        lines.append(f"{digest}  {f.name}")
    if not lines:
        print(
            f"Error: no matching files found for SHA256SUMS.txt ({[str(x) for x in files]})",
            file=sys.stderr,
        )
        return 1
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {len(lines)} SHA-256 entry(ies) to {out}")
    return 0


def _load_private_key(key_env_or_file: str | None):
    _hashes, serialization, _ec, _pre, _dec = _load_ec_module()
    raw = (key_env_or_file or os.environ.get("CLICKSHIFT_SIGNING_KEY_B64") or "").strip()
    if not raw:
        raise RuntimeError("Missing signing key (pass --key or set CLICKSHIFT_SIGNING_KEY_B64)")
    p = Path(raw)
    pem_bytes = p.read_bytes() if p.exists() and p.is_file() else base64.b64decode(raw)
    return serialization.load_pem_private_key(pem_bytes, password=None)


def sign_bytes_with_pem(digest32: bytes, priv) -> bytes:
    """Return raw 64-byte IEEE P1363 (r||s) ECDSA P-256 signature over ``digest32``."""
    hashes, _ser, ec, Prehashed, decode_dss_signature = _load_ec_module()
    der_sig = priv.sign(digest32, ec.ECDSA(Prehashed(hashes.SHA256())))
    r, s = decode_dss_signature(der_sig)
    return r.to_bytes(32, "big") + s.to_bytes(32, "big")


def cmd_sign(target: Path, out: Path | None, key_arg: str | None) -> int:
    priv = _load_private_key(key_arg)
    digest = _file_sha256(target)
    sig64 = sign_bytes_with_pem(digest, priv)
    sig_path = out or target.with_name(target.name + ".sig")
    sig_path.write_bytes(base64.b64encode(sig64))
    print(f"Signed {target.name} -> {sig_path}")
    return 0


def cmd_verify(target: Path, sig: Path, pubkey_b64: str) -> int:
    from clickshift import updater
    digest = _file_sha256(target)
    ok = updater.verify_ecdsa_p256_sha256(digest, sig.read_bytes(), [pubkey_b64])
    print("Signature valid:" if ok else "Signature INVALID!", target.name)
    return 0 if ok else 1


def cmd_blockmap(files: Sequence[Path | str]) -> int:
    from clickshift import updater

    resolved = _resolve_input_files(files)
    if not resolved:
        print(f"Error: no matching files found for blockmap ({[str(x) for x in files]})", file=sys.stderr)
        return 1
    for f in resolved:
        bm_bytes = updater.build_blockmap(f)
        out = f.with_name(f.name + ".blockmap")
        out.write_bytes(bm_bytes)
        print(f"Wrote blockmap ({len(bm_bytes) / 1024:.1f} KiB) -> {out}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="ClickShift release checksum & ECDSA signer")
    sub = ap.add_subparsers(dest="cmd", required=True)

    sub.add_parser("keygen", help="Generate a new ECDSA P-256 keypair")

    sp_sums = sub.add_parser("sums", help="Write SHA256SUMS.txt")
    sp_sums.add_argument("files", nargs="+", type=str)
    sp_sums.add_argument("--out", type=Path, required=True)

    sp_bm = sub.add_parser("blockmap", help="Write gzipped SHA-256 chunk .blockmap files for delta updates")
    sp_bm.add_argument("files", nargs="+", type=str)

    sp_sign = sub.add_parser("sign", help="Create <file>.sig using CLICKSHIFT_SIGNING_KEY_B64")
    sp_sign.add_argument("target", type=Path)
    sp_sign.add_argument("--out", type=Path, default=None)
    sp_sign.add_argument("--key", default=None)

    sp_ver = sub.add_parser("verify", help="Verify <file> + <file>.sig against public key")
    sp_ver.add_argument("target", type=Path)
    sp_ver.add_argument("sig", type=Path)
    sp_ver.add_argument("--pubkey", required=True)

    ns = ap.parse_args()
    if ns.cmd == "keygen":
        return cmd_keygen()
    if ns.cmd == "sums":
        return cmd_sums(ns.files, ns.out)
    if ns.cmd == "blockmap":
        return cmd_blockmap(ns.files)
    if ns.cmd == "sign":
        return cmd_sign(ns.target, ns.out, ns.key)
    return cmd_verify(ns.target, ns.sig, ns.pubkey)


if __name__ == "__main__":
    raise SystemExit(main())
