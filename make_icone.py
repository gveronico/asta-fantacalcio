#!/usr/bin/env python3
"""Genera le icone per la schermata Home (iOS/Android).

Uso: python make_icone.py
Crea icon-180.png (apple-touch-icon), icon-192.png e icon-512.png.
"""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
BG = (15, 20, 25)
RING = (91, 159, 212)
TEXT = (242, 245, 248)
SIZES = {180: "icon-180.png", 192: "icon-192.png", 512: "icon-512.png"}


def load_font(px: int):
    for name in ("segoeuib.ttf", "arialbd.ttf", "seguisb.ttf", "DejaVuSans-Bold.ttf"):
        try:
            return ImageFont.truetype(name, px)
        except OSError:
            continue
    return ImageFont.load_default()


def draw_icon(size: int) -> Image.Image:
    img = Image.new("RGB", (size, size), BG)
    d = ImageDraw.Draw(img)
    pad = size * 0.11
    d.ellipse([pad, pad, size - pad, size - pad], outline=RING, width=max(2, int(size * 0.045)))
    label = "A"
    font = load_font(int(size * 0.5))
    box = d.textbbox((0, 0), label, font=font)
    d.text(
        ((size - (box[2] - box[0])) / 2 - box[0], (size - (box[3] - box[1])) / 2 - box[1]),
        label,
        font=font,
        fill=TEXT,
    )
    return img


def main() -> int:
    for size, name in SIZES.items():
        draw_icon(size).save(ROOT / name, "PNG", optimize=True)
        print(f"creato {name} ({size}x{size})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
