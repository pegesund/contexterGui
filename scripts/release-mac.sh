#!/usr/bin/env bash
# Build, sign, notarize the Mac DMG for Spell, then (optionally) upload to
# pegesund/spell_binaries as a release asset under tag v<VERSION>.
#
# Usage:
#   bash scripts/release-mac.sh [VERSION] [options]
#
# Version selection (pick one — defaults to --patch if neither given):
#   <VERSION>           Explicit semver (e.g. 0.1.0)
#   --patch             Auto-bump patch (0.1.0 -> 0.1.1) from latest tag
#   --minor             Auto-bump minor (0.1.0 -> 0.2.0)
#   --major             Auto-bump major (0.1.0 -> 1.0.0)
#
# Options:
#   --dry-run           Print what would happen, do nothing
#   --no-upload         Build + sign + notarize, but don't push tag or upload
#   --no-notarize       Skip notarization (signing still happens)
#   --no-tag-push       Skip pushing git tag (won't trigger Windows CI)
#   --draft             Create the GitHub release as a draft
#
# Examples:
#   bash scripts/release-mac.sh                         # auto-patch bump
#   bash scripts/release-mac.sh --minor                 # auto-minor bump
#   bash scripts/release-mac.sh 0.2.0                   # explicit
#   bash scripts/release-mac.sh --patch --dry-run       # preview only
#   bash scripts/release-mac.sh --patch --no-upload     # local-only release build
#
# Requires:
#   - gh CLI authenticated with write access to pegesund/spell_binaries
#   - Xcode + Cognio Developer ID cert + Cognio-Notary keychain profile
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_DIR"

RELEASES_REPO="pegesund/spell_binaries"

# ── Defaults ─────────────────────────────────────────────────────────────────
EXPLICIT_VERSION=""
BUMP=""               # patch | minor | major | "" (default patch if no explicit)
DRY_RUN=false
DO_UPLOAD=true
DO_TAG_PUSH=true
NOTARIZE=true
DRAFT=false

# ── Parse args (order-independent) ───────────────────────────────────────────
while [ $# -gt 0 ]; do
    case "$1" in
        --patch)        BUMP="patch" ;;
        --minor)        BUMP="minor" ;;
        --major)        BUMP="major" ;;
        --dry-run)      DRY_RUN=true ;;
        --no-upload)    DO_UPLOAD=false; DO_TAG_PUSH=false ;;
        --no-notarize)  NOTARIZE=false ;;
        --no-tag-push)  DO_TAG_PUSH=false ;;
        --draft)        DRAFT=true ;;
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

# Default to --patch if neither explicit version nor bump specified
if [ -z "$EXPLICIT_VERSION" ] && [ -z "$BUMP" ]; then
    BUMP="patch"
fi
if [ -n "$EXPLICIT_VERSION" ] && [ -n "$BUMP" ]; then
    echo "ERROR: pass either an explicit version OR a bump flag, not both"; exit 2
fi

# ── Resolve version ──────────────────────────────────────────────────────────
latest_tag() {
    # Strict semver vX.Y.Z only — ignore pre-Spell-1.0 dev tags like
    # "v1.1-bert-sentence-scoring" left over from earlier work.
    # Source of truth: releases on $RELEASES_REPO. Fall back to local git tags.
    local from_gh
    from_gh="$( (gh release list --repo "$RELEASES_REPO" --limit 100 --json tagName 2>/dev/null \
        | python3 -c "
import json, sys, re
try:
    tags = [r['tagName'] for r in json.load(sys.stdin)]
except Exception:
    tags = []
tags = [t for t in tags if re.fullmatch(r'v\d+\.\d+\.\d+', t)]
tags.sort(key=lambda t: tuple(int(x) for x in t[1:].split('.')), reverse=True)
print(tags[0] if tags else '')
" 2>/dev/null) || true )"
    if [ -n "$from_gh" ]; then
        printf '%s\n' "$from_gh"
        return 0
    fi
    git fetch --tags --quiet 2>/dev/null || true
    ( git tag -l 'v*' 2>/dev/null \
        | grep -E '^v[0-9]+\.[0-9]+\.[0-9]+$' 2>/dev/null \
        | sort -t. -k1.2,1n -k2,2n -k3,3n -r \
        | head -1 ) || true
}

bump_version() {
    local current="$1" part="$2"
    IFS='.' read -r major minor patch <<<"${current#v}"
    case "$part" in
        major) echo "$((major + 1)).0.0" ;;
        minor) echo "${major}.$((minor + 1)).0" ;;
        patch) echo "${major}.${minor}.$((patch + 1))" ;;
    esac
}

if [ -n "$EXPLICIT_VERSION" ]; then
    VERSION="$EXPLICIT_VERSION"
else
    LATEST="$(latest_tag)"
    if [ -z "$LATEST" ]; then
        echo "No existing tags found on $RELEASES_REPO. Please pass an explicit version:"
        echo "  bash scripts/release-mac.sh 0.1.0"
        exit 1
    fi
    VERSION="$(bump_version "$LATEST" "$BUMP")"
    echo "  Latest released tag: $LATEST"
    echo "  Bumping ($BUMP) → v$VERSION"
fi

if ! [[ "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    echo "ERROR: bad version: $VERSION (must be semver like 0.1.0)"; exit 2
fi

TAG="v$VERSION"
ARCH="$(uname -m)"
DMG="$PROJECT_DIR/dist/releases/Spell-osx-${ARCH}-${VERSION}.dmg"

# ── Show plan ────────────────────────────────────────────────────────────────
echo
echo "================================================"
echo "  Spell Mac release"
echo "  Version:    $VERSION (tag: $TAG)"
echo "  Notarize:   $NOTARIZE"
echo "  Upload:     $DO_UPLOAD ($RELEASES_REPO)"
echo "  Tag push:   $DO_TAG_PUSH"
echo "  Draft:      $DRAFT"
echo "  DMG path:   $DMG"
echo "================================================"
echo

if $DRY_RUN; then
    actions=("build DMG (signed)")
    $NOTARIZE     && actions+=("notarize via Cognio-Notary")
    $DO_TAG_PUSH  && actions+=("push git tag $TAG")
    $DO_UPLOAD    && actions+=("create release $TAG on $RELEASES_REPO and upload DMG")
    echo "[DRY RUN] Would: $(printf ', %s' "${actions[@]}" | sed 's/^, //'). Nothing else."
    exit 0
fi

# ── Preflight ────────────────────────────────────────────────────────────────
if $DO_UPLOAD; then
    gh auth status >/dev/null 2>&1 || { echo "ERROR: gh CLI not authenticated. Run: gh auth login"; exit 1; }
    gh repo view "$RELEASES_REPO" >/dev/null 2>&1 || { echo "ERROR: cannot access $RELEASES_REPO"; exit 1; }
fi

# ── Build ────────────────────────────────────────────────────────────────────
BUILD_FLAGS=("--version" "$VERSION")
$NOTARIZE || BUILD_FLAGS+=("--no-notarize")
SPELL_VERSION="$VERSION" bash "$SCRIPT_DIR/build-mac.sh" "${BUILD_FLAGS[@]}"

[ -f "$DMG" ] || { echo "ERROR: DMG not produced at $DMG"; exit 1; }
echo "  built: $DMG ($(du -h "$DMG" | cut -f1))"

# ── Push git tag (if requested) ──────────────────────────────────────────────
if $DO_TAG_PUSH; then
    echo "=== Push git tag $TAG ==="
    if git rev-parse --git-dir >/dev/null 2>&1; then
        if git tag -l "$TAG" | grep -q .; then
            echo "  Tag $TAG already exists locally — skipping create."
        else
            git tag "$TAG"
        fi
        if git remote get-url origin >/dev/null 2>&1; then
            git push origin "$TAG" 2>&1 || echo "  WARN: tag push failed (Windows CI won't trigger). Push manually: git push origin $TAG"
        else
            echo "  No git remote 'origin' — skipping tag push (Windows CI won't auto-trigger)."
        fi
    else
        echo "  Not a git repo — skipping tag push."
    fi
fi

# ── Upload to spell_binaries ─────────────────────────────────────────────────
if $DO_UPLOAD; then
    echo "=== Create or update release $TAG on $RELEASES_REPO ==="
    if gh release view "$TAG" --repo "$RELEASES_REPO" >/dev/null 2>&1; then
        echo "  Release $TAG already exists — uploading asset to it."
    else
        DRAFT_FLAG=""
        $DRAFT && DRAFT_FLAG="--draft"
        gh release create "$TAG" \
            --repo "$RELEASES_REPO" \
            --title "Spell $VERSION" \
            --notes "Automated release. Mac DMG uploaded by release-mac.sh; Windows zip uploaded by release-windows.yml CI." \
            $DRAFT_FLAG
    fi

    echo "=== Upload Mac DMG ==="
    gh release upload "$TAG" "$DMG" --repo "$RELEASES_REPO" --clobber
fi

# ── Summary ──────────────────────────────────────────────────────────────────
echo
echo "================================================"
if $DO_UPLOAD; then
    echo "  Release $TAG complete (Mac side)"
    echo "  URL: https://github.com/$RELEASES_REPO/releases/tag/$TAG"
    if $DO_TAG_PUSH; then
        echo "  Windows CI:  https://github.com/pegesund/contexterGui/actions"
    fi
else
    echo "  Local build complete (no upload)"
    echo "  DMG: $DMG"
fi
echo "================================================"
