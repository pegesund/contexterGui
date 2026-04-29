#!/usr/bin/env python3
"""Generate a placeholder Spell app icon (1024x1024 PNG)."""
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("Spell-1024.png")

SIZE = 1024
RADIUS = 230
BG = (37, 99, 235)
ACCENT = (59, 130, 246)
TEXT = (255, 255, 255)
DOT = (251, 191, 36)

img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

d.rounded_rectangle((0, 0, SIZE, SIZE), RADIUS, fill=BG)
d.rounded_rectangle((40, 40, SIZE - 40, SIZE - 40), RADIUS - 40, outline=ACCENT, width=8)

font = None
for path in [
    "/System/Library/Fonts/Helvetica.ttc",
    "/System/Library/Fonts/SFCompact.ttf",
    "/Library/Fonts/Arial Bold.ttf",
]:
    try:
        font = ImageFont.truetype(path, 640)
        break
    except OSError:
        continue
if font is None:
    font = ImageFont.load_default()

text = "S"
bbox = d.textbbox((0, 0), text, font=font)
tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
tx = (SIZE - tw) / 2 - bbox[0]
ty = (SIZE - th) / 2 - bbox[1] - 20
d.text((tx, ty), text, fill=TEXT, font=font)

dot_r = 60
d.ellipse((SIZE - 240, SIZE - 240, SIZE - 240 + dot_r * 2, SIZE - 240 + dot_r * 2), fill=DOT)

img.save(OUT, "PNG")
print(f"wrote {OUT} ({SIZE}x{SIZE})")
