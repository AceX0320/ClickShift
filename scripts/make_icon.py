"""Generate assets/clickshift.ico and assets/clickshift.png only when missing (or --force).

When custom assets/clickshift.ico already exists, missing clickshift.png is derived
directly from the highest-resolution layer of assets/clickshift.ico instead of
overwriting the designed icon.
"""
import sys
from pathlib import Path

try:
    from PIL import Image, ImageDraw
except ImportError:
    raise SystemExit("pip install Pillow first")


def make(size: int) -> Image.Image:
    """Fallback geometric icon matching the obsidian / crimson-white pages / cursor design."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    m = max(1, int(size * 0.04))
    r = max(2, int(size * 0.22))
    # Obsidian rounded square
    d.rounded_rectangle([m, m, size - m, size - m], radius=r, fill=(20, 10, 11, 255))
    # Shifted crimson back page
    d.rounded_rectangle(
        [int(size * 0.40), int(size * 0.11), int(size * 0.80), int(size * 0.65)],
        radius=max(2, size // 18),
        fill=(225, 29, 46, 255),
    )
    # Front white page
    d.rounded_rectangle(
        [int(size * 0.20), int(size * 0.24), int(size * 0.61), int(size * 0.80)],
        radius=max(2, size // 18),
        fill=(255, 255, 255, 255),
        outline=(11, 11, 11, 255),
        width=max(1, size // 48),
    )
    # Pointer cursor
    d.polygon([
        (int(size * 0.44), int(size * 0.48)),
        (int(size * 0.44), int(size * 0.83)),
        (int(size * 0.53), int(size * 0.74)),
        (int(size * 0.59), int(size * 0.89)),
        (int(size * 0.65), int(size * 0.86)),
        (int(size * 0.59), int(size * 0.72)),
        (int(size * 0.71), int(size * 0.72)),
    ], fill=(11, 11, 11, 255))
    return img


if __name__ == "__main__":
    force = "--force" in sys.argv[1:]
    out = Path(__file__).resolve().parent.parent / "assets"
    out.mkdir(exist_ok=True)
    ico_path = out / "clickshift.ico"
    png_path = out / "clickshift.png"

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
