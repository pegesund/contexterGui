#!/usr/bin/env bash
# Bump manifest.json version and build the Chrome Web Store .zip in one go.
# Mirrors the flag syntax of contexterGui's scripts/release-mac.sh so the
# muscle memory carries over.
#
# Usage:
#   bash scripts/release.sh [VERSION] [options]
#
# Version selection (pick one — defaults to --patch if neither given):
#   <VERSION>           Explicit semver (e.g. 0.1.0)
#   --patch             Auto-bump patch (0.1.0 -> 0.1.1) from manifest.json
#   --minor             Auto-bump minor (0.1.0 -> 0.2.0)
#   --major             Auto-bump major (0.1.0 -> 1.0.0)
#
# Options:
#   --dry-run           Print what would happen, change nothing
#   --no-build          Only bump manifest.json, don't run build-extension.sh
#
# Examples:
#   bash scripts/release.sh                       # auto-patch bump + build
#   bash scripts/release.sh --minor               # auto-minor bump + build
#   bash scripts/release.sh 0.2.0                 # explicit version + build
#   bash scripts/release.sh --patch --dry-run     # preview only
#   bash scripts/release.sh --patch --no-build    # bump only, skip zip
#
# Output:
#   dist/Spell-Browser-Extension-<version>.zip
#
# Git: this script does NOT commit or tag. The companion extension lives
# inside the contexterGui repo whose tags drive the desktop app's release
# CI (v0.1.0 etc.), so auto-tagging here would collide with desktop tags
# and trigger CI accidentally. Commit by hand after reviewing the diff:
#   cd ../..                       # back to contexterGui root
#   git diff extension/manifest.json
#   git add extension/manifest.json
#   git commit -m "Browser ext: bump to vX.Y.Z"
#   git push origin main
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
EXT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$EXT_DIR"

# ── Defaults ─────────────────────────────────────────────────────────────────
EXPLICIT_VERSION=""
BUMP=""
DRY_RUN=false
DO_BUILD=true

# ── Parse args ───────────────────────────────────────────────────────────────
while [ $# -gt 0 ]; do
    case "$1" in
        --patch)     BUMP="patch" ;;
        --minor)     BUMP="minor" ;;
        --major)     BUMP="major" ;;
        --dry-run)   DRY_RUN=true ;;
        --no-build)  DO_BUILD=false ;;
        -h|--help)
            sed -n '2,30p' "$0" | sed 's/^# \{0,1\}//'
            exit 0
            ;;
        [0-9]*.[0-9]*.[0-9]*)
            EXPLICIT_VERSION="$1"
            ;;
        *) echo "ERROR: unknown arg: $1"; echo "Run: $0 --help"; exit 2 ;;
    esac
    shift
done

if [ -z "$EXPLICIT_VERSION" ] && [ -z "$BUMP" ]; then
    BUMP="patch"
fi
if [ -n "$EXPLICIT_VERSION" ] && [ -n "$BUMP" ]; then
    echo "ERROR: pass either an explicit version OR a bump flag, not both"; exit 2
fi

# ── Resolve version ──────────────────────────────────────────────────────────
CURRENT="$(python3 -c "import json; print(json.load(open('manifest.json'))['version'])")"

bump_version() {
    local current="$1" part="$2"
    IFS='.' read -r major minor patch <<<"$current"
    case "$part" in
        major) echo "$((major + 1)).0.0" ;;
        minor) echo "${major}.$((minor + 1)).0" ;;
        patch) echo "${major}.${minor}.$((patch + 1))" ;;
    esac
}

if [ -n "$EXPLICIT_VERSION" ]; then
    VERSION="$EXPLICIT_VERSION"
else
    VERSION="$(bump_version "$CURRENT" "$BUMP")"
fi

if ! [[ "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    echo "ERROR: bad version: $VERSION (Chrome Web Store requires X.Y.Z, no suffixes)"; exit 2
fi

ZIP="$EXT_DIR/dist/Spell-Browser-Extension-${VERSION}.zip"

# ── Show plan ────────────────────────────────────────────────────────────────
echo
echo "================================================"
echo "  Spell browser ext (companion) release"
echo "  Current:    $CURRENT"
echo "  New:        $VERSION"
echo "  Build:      $DO_BUILD"
if $DO_BUILD; then
    echo "  Output:     $ZIP"
fi
echo "================================================"
echo

if $DRY_RUN; then
    echo "[DRY RUN] Would update manifest.json $CURRENT -> $VERSION"
    $DO_BUILD && echo "[DRY RUN] Would run scripts/build-extension.sh to produce $ZIP"
    exit 0
fi

# ── Bump manifest.json in place ──────────────────────────────────────────────
python3 - <<PY
import json
with open('manifest.json', 'r') as f:
    m = json.load(f)
m['version'] = '$VERSION'
with open('manifest.json', 'w') as f:
    json.dump(m, f, indent=2)
    f.write('\n')
PY
echo "  manifest.json: $CURRENT -> $VERSION"

# ── Build the zip ────────────────────────────────────────────────────────────
if $DO_BUILD; then
    echo
    bash scripts/build-extension.sh
    echo
    echo "  Done. Next steps:"
    echo "    1. Review the manifest change:"
    echo "       (cd ../.. && git diff extension/manifest.json)"
    echo "    2. Commit (no tag — would collide with desktop release tags):"
    echo "       (cd ../.. && git add extension/manifest.json \\"
    echo "          && git commit -m 'Browser ext: bump to v$VERSION' \\"
    echo "          && git push origin main)"
    echo "    3. Upload $ZIP to:"
    echo "       https://chrome.google.com/webstore/devconsole/"
else
    echo
    echo "  Skipped build (--no-build). To build the zip:"
    echo "    bash scripts/build-extension.sh"
fi
