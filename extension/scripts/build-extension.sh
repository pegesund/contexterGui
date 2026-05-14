#!/usr/bin/env bash
# Package the native-messaging Spell browser extension as a Chrome Web Store .zip.
#
# Output: dist/Spell-Browser-Extension-<version>.zip
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
EXT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
DIST="$EXT_DIR/dist"
STAGE="$DIST/staging"

# Read version from manifest.json (avoids drift)
VERSION="$(python3 -c "import json; print(json.load(open('$EXT_DIR/manifest.json'))['version'])")"
ZIP="$DIST/Spell-Browser-Extension-${VERSION}.zip"

echo "=== Build Spell browser extension v$VERSION ==="
rm -rf "$STAGE" "$ZIP"
mkdir -p "$STAGE"

# Files that ship to Chrome Web Store
declare -a SHIP=(
    manifest.json
    background.js
    content.js
    gdocs-inject.js
    icons
)

# Files that do NOT ship (sanity-check that they exist locally so we don't
# silently miss a critical file by typo)
declare -a NO_SHIP=(
    com.cognio.spell.bridge.json   # native messaging host config (ships with desktop installer, not extension)
    install_native_host.bat     # Windows-only desktop-side install helper
    setup.html                  # not referenced by manifest, dev-only
)

for f in "${SHIP[@]}"; do
    if [ ! -e "$EXT_DIR/$f" ]; then
        echo "ERROR: required file missing: $f"; exit 1
    fi
    cp -R "$EXT_DIR/$f" "$STAGE/"
done

# Validate JSON
python3 -c "import json; json.load(open('$STAGE/manifest.json'))" \
    || { echo "ERROR: manifest.json is invalid JSON"; exit 1; }

# Verify manifest has required Chrome Web Store fields
python3 - <<PY
import json, sys
m = json.load(open('$STAGE/manifest.json'))
errors = []
for key in ('manifest_version', 'name', 'version', 'description', 'icons'):
    if key not in m: errors.append(f"missing top-level key: {key}")
if m.get('manifest_version') != 3: errors.append('manifest_version must be 3')
for size in ('16', '48', '128'):
    if size not in m.get('icons', {}): errors.append(f"missing icon size: {size}")
if errors:
    for e in errors: print('  '+e)
    sys.exit(1)
print('  manifest validation OK')
PY

cd "$STAGE" && zip -qr "$ZIP" .
echo "  wrote $ZIP ($(du -h "$ZIP" | cut -f1))"
echo "  files in zip:"
unzip -l "$ZIP" | awk 'NR>3 && $1 != "----" && $1 != "" && $4 != "" {print "    " $4}' | sed '$d'
echo
echo "Done. Upload via Chrome Web Store dev console (see SUBMISSION_GUIDE.md)."
