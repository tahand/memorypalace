"""Generate the Memory Palace app icon (a simple columned temple) as a .ico file.

    python tools/make_icon.py build/icon.ico
"""
import sys
from pathlib import Path

from PIL import Image, ImageDraw


def draw(size: int = 256) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    s = size / 256
    d.rounded_rectangle((8 * s, 8 * s, 248 * s, 248 * s), radius=48 * s, fill=(31, 106, 165))
    white = (255, 255, 255)
    d.polygon([(128 * s, 40 * s), (220 * s, 96 * s), (36 * s, 96 * s)], fill=white)  # roof
    d.rectangle((44 * s, 104 * s, 212 * s, 116 * s), fill=white)                      # architrave
    for x in (56, 96, 140, 180):                                                        # columns
        d.rectangle((x * s, 124 * s, (x + 20) * s, 192 * s), fill=white)
    d.rectangle((36 * s, 198 * s, 220 * s, 212 * s), fill=white)                      # base
    return img


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "build/icon.ico")
    out.parent.mkdir(parents=True, exist_ok=True)
    draw().save(out, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print(f"wrote {out}")
