"""Generate assets/confas.ico and assets/confas.png only when missing (or --force).

When custom assets/confas.ico already exists, missing confas.png is derived
directly from the highest-resolution layer of assets/confas.ico instead of
overwriting the designed icon.
"""
import sys
from pathlib import Path

try:
    from PIL import Image, ImageDraw
except ImportError:
    raise SystemExit("pip install Pillow first")


def make(size: int) -> Image.Image:
    """Fallback geometric icon matching the navy / split-doc / gold-bolt design."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    m = max(1, int(size * 0.04))
    r = max(2, int(size * 0.22))
    # Dark navy rounded square
    d.rounded_rectangle([m, m, size - m, size - m], radius=r, fill=(13, 30, 62, 255))
    # Left white document half
    x0, x_mid, x1 = int(size * 0.24), int(size * 0.50), int(size * 0.76)
    y0, y1 = int(size * 0.17), int(size * 0.83)
    d.rounded_rectangle([x0, y0, x1, y1], radius=max(2, size // 18), fill=(56, 189, 248, 255))
    d.rectangle([x0, y0, x_mid, y1], fill=(248, 250, 252, 255))
    # Gold lightning bolt
    d.polygon([
        (int(size * 0.54), int(size * 0.26)),
        (int(size * 0.38), int(size * 0.54)),
        (int(size * 0.50), int(size * 0.54)),
        (int(size * 0.45), int(size * 0.76)),
        (int(size * 0.63), int(size * 0.46)),
        (int(size * 0.51), int(size * 0.46)),
    ], fill=(250, 204, 21, 255))
    return img


if __name__ == "__main__":
    force = "--force" in sys.argv[1:]
    out = Path(__file__).resolve().parent.parent / "assets"
    out.mkdir(exist_ok=True)
    ico_path = out / "confas.ico"
    png_path = out / "confas.png"

    if ico_path.exists() and not force:
        if not png_path.exists():
            src = Image.open(ico_path)
            src.load()
            src.resize((256, 256), Image.LANCZOS).convert("RGBA").save(png_path, format="PNG")
            print("derived", png_path, "from existing", ico_path)
        else:
            print("keeping existing", ico_path, "and", png_path)
        raise SystemExit(0)

    base = make(256)
    base.save(ico_path, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    base.save(png_path, format="PNG")
    print("wrote", ico_path, "and", png_path)
