#!/usr/bin/env python3
"""Generate the DMG background for Spell.

Produces a 1200x800 PNG at 144 DPI. macOS Finder displays the DMG window at
600x400 logical pixels; the backing image is 2x for retina sharpness.

Layout:
    [Spell.app icon]         →         [Applications folder]
              with a soft arrow pointing right between them
                          "Dra Spell til Applications"
"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("assets/dmg_background.png")
W, H = 1200, 800

# Spell brand colors
BG_TOP = (240, 245, 255)     # very light blue-white
BG_BOTTOM = (255, 255, 255)  # white
ARROW = (37, 99, 235, 80)    # Spell blue, semi-transparent
TEXT = (60, 70, 90)
SUBTLE = (140, 150, 170)

img = Image.new("RGBA", (W, H), BG_BOTTOM)
d = ImageDraw.Draw(img)

# Vertical gradient
for y in range(H):
    t = y / H
    r = int(BG_TOP[0] * (1 - t) + BG_BOTTOM[0] * t)
    g = int(BG_TOP[1] * (1 - t) + BG_BOTTOM[1] * t)
    b = int(BG_TOP[2] * (1 - t) + BG_BOTTOM[2] * t)
    d.line([(0, y), (W, y)], fill=(r, g, b, 255))

# Soft arrow between the two icon spots
# Icons sit at logical (140, 180) and (400, 180); in 2x backing that's (280, 360) and (800, 360)
# Draw arrow from x=420 to x=720 at y=400
arrow_y = 400
arrow_overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
ad = ImageDraw.Draw(arrow_overlay)
# Shaft
ad.rectangle([(420, arrow_y - 8), (720, arrow_y + 8)], fill=ARROW)
# Head (triangle)
ad.polygon([(720, arrow_y - 30), (720, arrow_y + 30), (770, arrow_y)], fill=ARROW)
img = Image.alpha_composite(img, arrow_overlay)
d = ImageDraw.Draw(img)

# Tagline at bottom
font_main = None
font_small = None
for path in [
    "/System/Library/Fonts/Helvetica.ttc",
    "/System/Library/Fonts/SFCompact.ttf",
    "/Library/Fonts/Arial.ttf",
]:
    try:
        font_main = ImageFont.truetype(path, 44)
        font_small = ImageFont.truetype(path, 28)
        break
    except OSError:
        continue
if font_main is None:
    font_main = ImageFont.load_default()
    font_small = font_main

text = "Spell"
bbox = d.textbbox((0, 0), text, font=font_main)
tw = bbox[2] - bbox[0]
d.text(((W - tw) / 2 - bbox[0], 600), text, fill=TEXT, font=font_main)

sub = "Norsk staving og grammatikksjekk"
bbox = d.textbbox((0, 0), sub, font=font_small)
tw = bbox[2] - bbox[0]
d.text(((W - tw) / 2 - bbox[0], 670), sub, fill=SUBTLE, font=font_small)

img.save(OUT, "PNG", optimize=True)
print(f"wrote {OUT} ({W}x{H} @ 2x, {OUT.stat().st_size}B)")
