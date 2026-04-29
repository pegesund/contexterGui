#!/usr/bin/env python3
"""Resize Spell-1024.png into the icon sizes Chrome extensions need."""
import sys
from pathlib import Path
from PIL import Image

SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("assets/Spell-1024.png")
OUT_DIR = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("assets/extension-icons")

OUT_DIR.mkdir(parents=True, exist_ok=True)
src = Image.open(SRC).convert("RGBA")

for size in (16, 32, 48, 128):
    img = src.resize((size, size), Image.LANCZOS)
    out = OUT_DIR / f"icon-{size}.png"
    img.save(out, "PNG", optimize=True)
    print(f"  wrote {out} ({size}x{size}, {out.stat().st_size}B)")
