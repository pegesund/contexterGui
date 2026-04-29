#!/usr/bin/env bash
# Build Spell.icns from a 1024x1024 source PNG.
set -euo pipefail
SRC="${1:-assets/Spell-1024.png}"
OUT="${2:-assets/Spell.icns}"
ICONSET=$(mktemp -d)/Spell.iconset
mkdir -p "$ICONSET"

declare -a SIZES=(16 32 64 128 256 512 1024)
for s in "${SIZES[@]}"; do
  sips -z "$s" "$s" "$SRC" --out "$ICONSET/icon_${s}x${s}.png" >/dev/null
  if [ "$s" -lt 1024 ]; then
    s2=$((s * 2))
    sips -z "$s2" "$s2" "$SRC" --out "$ICONSET/icon_${s}x${s}@2x.png" >/dev/null
  fi
done

iconutil -c icns "$ICONSET" -o "$OUT"
echo "wrote $OUT ($(wc -c <"$OUT") bytes)"
