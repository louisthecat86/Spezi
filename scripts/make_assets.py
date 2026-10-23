"""Erzeugt GUI- und Build-Assets: logo.png (Header), icon.png + SpeziSmith.ico.

Nur fuer Entwicklung/Build noetig - die fertige EXE nutzt die erzeugten Dateien.
Aufruf: python scripts/make_assets.py
"""

import os

from PIL import Image, ImageDraw

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(ROOT, "assets")

NAVY = (31, 41, 66)       # AULBACH-Blau
ORANGE = (240, 140, 33)   # AULBACH-Orange
WHITE = (255, 255, 255)


def make_logo_png():
    src = Image.open(os.path.join(ASSETS, "logo.jpeg")).convert("RGB")
    target_h = 96
    w = round(src.width * target_h / src.height)
    src.resize((w, target_h), Image.LANCZOS).save(os.path.join(ASSETS, "logo.png"))
    print("logo.png:", w, "x", target_h)


def rounded(draw, box, radius, **kw):
    draw.rounded_rectangle(box, radius=radius, **kw)


def make_icon():
    size = 512
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # Hintergrund: abgerundetes Quadrat in Navy
    rounded(d, (8, 8, size - 8, size - 8), 110, fill=NAVY)

    # Dokument
    doc = (118, 96, 330, 416)
    rounded(d, doc, 28, fill=WHITE)
    # Zeilen im Dokument
    for i, w in enumerate((150, 170, 120)):
        y = 160 + i * 56
        rounded(d, (150, y, 150 + w, y + 22), 11, fill=(208, 216, 228))
    # Dokument-Titelzeile in Orange
    rounded(d, (150, 120, 262, 138), 9, fill=ORANGE)

    # Stift (orange, diagonal)
    d.polygon([(300, 430), (430, 300), (472, 342), (342, 472)], fill=ORANGE)
    d.polygon([(300, 430), (318, 412), (336, 430), (342, 472)], fill=(198, 108, 20))
    d.polygon([(430, 300), (452, 278), (494, 320), (472, 342)], fill=WHITE)

    img.save(os.path.join(ASSETS, "icon.png"))

    ico_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    img.save(
        os.path.join(ASSETS, "SpeziSmith.ico"),
        sizes=ico_sizes,
        format="ICO",
    )
    print("icon.png + SpeziSmith.ico erzeugt")


if __name__ == "__main__":
    make_logo_png()
    make_icon()
