#!/usr/bin/env bash
# Build, sign, notarize the Mac DMG, then upload to pegesund/spell_binaries
# as a release asset under tag v<VERSION>.
#
# Usage:
#   bash scripts/release-mac.sh <version> [--draft] [--no-tag-push]
#
# Examples:
#   bash scripts/release-mac.sh 0.1.0
#   bash scripts/release-mac.sh 0.1.1 --draft
#
# Requires:
#   - gh CLI authenticated (via keyring) with write access to spell_binaries
#   - Xcode + Cognio Developer ID cert + Cognio-Notary keychain profile
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_DIR"

# ── Args ─────────────────────────────────────────────────────────────────────
[ $# -ge 1 ] || { echo "Usage: $0 <version> [--draft] [--no-tag-push]"; exit 2; }
VERSION="$1"; shift
DRAFT=false
TAG_PUSH=true
while [ $# -gt 0 ]; do
    case "$1" in
        --draft) DRAFT=true ;;
        --no-tag-push) TAG_PUSH=false ;;
        *) echo "Unknown arg: $1"; exit 2 ;;
    esac
    shift
done

if ! [[ "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    echo "ERROR: version must be semver like 0.1.0 (got: $VERSION)"; exit 2
fi

TAG="v$VERSION"
RELEASES_REPO="pegesund/spell_binaries"

# ── Preflight ────────────────────────────────────────────────────────────────
echo "=== Preflight ==="
gh auth status >/dev/null 2>&1 || { echo "ERROR: gh CLI not authenticated. Run: gh auth login"; exit 1; }
gh repo view "$RELEASES_REPO" >/dev/null 2>&1 || { echo "ERROR: cannot access $RELEASES_REPO"; exit 1; }

ARCH="$(uname -m)"  # arm64 on Apple Silicon
DMG="$PROJECT_DIR/dist/releases/Spell-osx-${ARCH}-${VERSION}.dmg"

# ── Build (signed + notarized + DMG) ─────────────────────────────────────────
echo "=== Build (signed + notarized + DMG) ==="
SPELL_VERSION="$VERSION" bash "$SCRIPT_DIR/build-mac.sh" --version "$VERSION"

[ -f "$DMG" ] || { echo "ERROR: DMG not produced at $DMG"; exit 1; }
echo "  built: $DMG ($(du -h "$DMG" | cut -f1))"

# ── Push git tag (if requested + we're in a git repo) ────────────────────────
if $TAG_PUSH; then
    echo "=== Push git tag $TAG ==="
    if git -C "$PROJECT_DIR" rev-parse --git-dir >/dev/null 2>&1; then
        if git -C "$PROJECT_DIR" tag -l "$TAG" | grep -q .; then
            echo "  Tag $TAG already exists locally — skipping create."
        else
            git -C "$PROJECT_DIR" tag "$TAG"
        fi
        # Push tag — best effort. If origin doesn't exist or push fails we keep going.
        if git -C "$PROJECT_DIR" remote get-url origin >/dev/null 2>&1; then
            git -C "$PROJECT_DIR" push origin "$TAG" 2>&1 || echo "  WARN: tag push failed (Windows CI won't trigger). Push manually: git push origin $TAG"
        else
            echo "  No git remote 'origin' — skipping tag push (Windows CI won't auto-trigger)."
        fi
    else
        echo "  Not a git repo — skipping tag push."
    fi
fi

# ── Create / find release in spell_binaries ──────────────────────────────────
echo "=== Create or update release $TAG on $RELEASES_REPO ==="
if gh release view "$TAG" --repo "$RELEASES_REPO" >/dev/null 2>&1; then
    echo "  Release $TAG already exists — uploading asset to it."
else
    DRAFT_FLAG=""
    $DRAFT && DRAFT_FLAG="--draft"
    gh release create "$TAG" \
        --repo "$RELEASES_REPO" \
        --title "Spell $VERSION" \
        --notes "Automated release from release-mac.sh. Mac and Windows binaries are uploaded by their respective release pipelines." \
        $DRAFT_FLAG
fi

# ── Upload Mac DMG ───────────────────────────────────────────────────────────
echo "=== Upload Mac DMG ==="
gh release upload "$TAG" "$DMG" \
    --repo "$RELEASES_REPO" \
    --clobber

URL="https://github.com/$RELEASES_REPO/releases/tag/$TAG"
echo
echo "================================================"
echo "  Mac upload complete!"
echo "  Tag:     $TAG"
echo "  DMG:     $(basename "$DMG")"
echo "  URL:     $URL"
echo
if $TAG_PUSH && git -C "$PROJECT_DIR" rev-parse --git-dir >/dev/null 2>&1; then
    echo "  Windows CI: should be running on push of $TAG. Check:"
    echo "    https://github.com/pegesund/contexterGui/actions"
fi
echo "================================================"
