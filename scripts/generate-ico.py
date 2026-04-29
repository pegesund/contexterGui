#!/usr/bin/env python3
"""Generate Spell.ico from Spell-1024.png with Windows installer-friendly sizes."""
import sys
from pathlib import Path
from PIL import Image

SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("assets/Spell-1024.png")
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("assets/Spell.ico")

src = Image.open(SRC).convert("RGBA")
sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
src.save(OUT, format="ICO", sizes=sizes)
print(f"wrote {OUT} ({OUT.stat().st_size}B, sizes: {[s[0] for s in sizes]})")
