#!/usr/bin/env bash
# Build, package, and publish a Spell browser companion ext release.
#
# End-to-end pipeline mirroring contexterGui/scripts/release-mac.sh — the
# flag syntax matches so the muscle memory carries over from the desktop
# release flow.
#
# Usage:
#   bash scripts/release.sh [VERSION] [options]
#
# Version selection (pick one — defaults to --patch if neither given):
#   <VERSION>           Explicit semver (e.g. 0.1.0)
#   --patch             Auto-bump patch (0.1.0 -> 0.1.1) from latest tag
#   --minor             Auto-bump minor (0.1.0 -> 0.2.0)
#   --major             Auto-bump major (0.1.0 -> 1.0.0)
#
# Options:
#   --dry-run           Print what would happen, change nothing
#   --no-build          Bump manifest only, don't run build-extension.sh
#   --no-commit         Bump + build but don't git commit/push or upload
#   --no-tag-push       Build + commit but don't push, don't upload
#   --no-upload         Tag and push but don't create the GitHub release
#   --draft             Create the GitHub release as a draft (link-only access)
#
# Examples:
#   bash scripts/release.sh                       # auto-patch + full pipeline
#   bash scripts/release.sh --minor               # auto-minor + full pipeline
#   bash scripts/release.sh 0.2.0                 # explicit + full pipeline
#   bash scripts/release.sh --patch --dry-run     # preview only
#   bash scripts/release.sh --patch --no-upload   # commit + push, no GH release
#   bash scripts/release.sh --patch --draft       # GH release as draft (Petter-only link)
#
# Output:
#   dist/Spell-Browser-Extension-<version>.zip
#   GitHub release at https://github.com/pegesund/spell_binaries/releases/tag/browser-ext-v<version>
#
# Tag namespace:
#   The contexterGui repo's `vX.Y.Z` tags are used by the desktop release
#   pipeline (release-mac.sh) and trigger Windows CI. To avoid collisions,
#   browser ext releases use the `browser-ext-vX.Y.Z` prefix AND publish to
#   the spell_binaries repo (where desktop DMGs already live), not
#   contexterGui's own tags. The source-code bump to extension/manifest.json
#   still gets committed + pushed to contexterGui main for traceability.
#
# Requires:
#   - gh CLI authenticated with write access to pegesund/spell_binaries
#     (only needed when uploading; --no-upload skips this)
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
EXT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
REPO_DIR="$(cd "$EXT_DIR/.." && pwd)"
cd "$EXT_DIR"

RELEASES_REPO="pegesund/spell_binaries"
TAG_PREFIX="browser-ext-v"

# ── Defaults ─────────────────────────────────────────────────────────────────
EXPLICIT_VERSION=""
BUMP=""
DRY_RUN=false
DO_BUILD=true
DO_COMMIT=true
DO_TAG_PUSH=true
DO_UPLOAD=true
DRAFT=false

# ── Parse args ───────────────────────────────────────────────────────────────
while [ $# -gt 0 ]; do
    case "$1" in
        --patch)        BUMP="patch" ;;
        --minor)        BUMP="minor" ;;
        --major)        BUMP="major" ;;
        --dry-run)      DRY_RUN=true ;;
        --no-build)     DO_BUILD=false; DO_COMMIT=false; DO_TAG_PUSH=false; DO_UPLOAD=false ;;
        --no-commit)    DO_COMMIT=false; DO_TAG_PUSH=false; DO_UPLOAD=false ;;
        --no-tag-push)  DO_TAG_PUSH=false; DO_UPLOAD=false ;;
        --no-upload)    DO_UPLOAD=false ;;
        --draft)        DRAFT=true ;;
        -h|--help)
            sed -n '2,42p' "$0" | sed 's/^# \{0,1\}//'
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
# Source of truth priority: latest browser-ext-v*.*.* tag on the releases
# repo (most authoritative — that's the user-visible version), then fall
# back to manifest.json's current version.
current_manifest_version() {
    python3 -c "import json; print(json.load(open('manifest.json'))['version'])"
}

latest_tag() {
    local from_gh
    from_gh="$( (gh release list --repo "$RELEASES_REPO" --limit 100 --json tagName 2>/dev/null \
        | python3 -c "
import json, sys, re
try:
    tags = [r['tagName'] for r in json.load(sys.stdin)]
except Exception:
    tags = []
# Only match our prefix
pat = re.compile(r'^browser-ext-v(\d+)\.(\d+)\.(\d+)$')
matches = [(t, tuple(int(x) for x in pat.match(t).groups())) for t in tags if pat.match(t)]
matches.sort(key=lambda x: x[1], reverse=True)
print(matches[0][0] if matches else '')
" 2>/dev/null) || true )"
    printf '%s\n' "$from_gh"
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

CURRENT_MANIFEST="$(current_manifest_version)"

if [ -n "$EXPLICIT_VERSION" ]; then
    VERSION="$EXPLICIT_VERSION"
    BASE_FOR_BUMP=""
else
    LATEST_TAG="$(latest_tag)"
    if [ -z "$LATEST_TAG" ]; then
        # First-ever release — bump from manifest's current version
        BASE_FOR_BUMP="$CURRENT_MANIFEST"
    else
        BASE_FOR_BUMP="${LATEST_TAG#$TAG_PREFIX}"
    fi
    VERSION="$(bump_version "$BASE_FOR_BUMP" "$BUMP")"
fi

if ! [[ "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    echo "ERROR: bad version: $VERSION (Chrome Web Store requires X.Y.Z, no suffixes)"; exit 2
fi

TAG="${TAG_PREFIX}${VERSION}"
ZIP="$EXT_DIR/dist/Spell-Browser-Extension-${VERSION}.zip"

# ── Show plan ───────────────────────────────────────────────────────────────
echo
echo "================================================"
echo "  Spell browser ext (companion) release"
echo "  Current manifest:  $CURRENT_MANIFEST"
if [ -n "${BASE_FOR_BUMP:-}" ] && [ "$BASE_FOR_BUMP" != "$CURRENT_MANIFEST" ]; then
    echo "  Latest tag:        ${TAG_PREFIX}${BASE_FOR_BUMP}"
fi
echo "  New version:       $VERSION"
echo "  Tag:               $TAG (on $RELEASES_REPO)"
echo "  Build:             $DO_BUILD"
echo "  Commit + push:     $DO_COMMIT (to contexterGui main)"
echo "  Push tag:          $DO_TAG_PUSH"
echo "  GitHub release:    $DO_UPLOAD ($RELEASES_REPO)"
$DRAFT && echo "  Release type:      DRAFT"
if $DO_BUILD; then
    echo "  Zip:               $ZIP"
fi
echo "================================================"
echo

if $DRY_RUN; then
    actions=()
    actions+=("bump extension/manifest.json $CURRENT_MANIFEST -> $VERSION")
    $DO_BUILD     && actions+=("build $ZIP")
    $DO_COMMIT    && actions+=("git commit + push to contexterGui main")
    $DO_TAG_PUSH  && actions+=("create + push tag $TAG to $RELEASES_REPO")
    $DO_UPLOAD    && actions+=("create GH release $TAG on $RELEASES_REPO + upload zip")
    echo "[DRY RUN] Would: $(printf ', %s' "${actions[@]}" | sed 's/^, //'). Nothing else."
    exit 0
fi

# ── Pre-flight checks ───────────────────────────────────────────────────────
if $DO_COMMIT; then
    # Only check extension/ tree — contexterGui repo may have desktop WIP we don't touch.
    if [ -n "$(cd "$REPO_DIR" && git status --porcelain -- extension/)" ]; then
        echo "ERROR: extension/ has uncommitted changes. Stash or commit first." >&2
        echo "       (run with --no-commit to bump + build without git ops)" >&2
        (cd "$REPO_DIR" && git status --short -- extension/) >&2
        exit 1
    fi
fi

if $DO_UPLOAD; then
    if ! command -v gh >/dev/null 2>&1; then
        echo "ERROR: gh CLI not found but --no-upload not set. brew install gh" >&2
        exit 1
    fi
    if ! gh auth status >/dev/null 2>&1; then
        echo "ERROR: gh CLI not authenticated. Run: gh auth login" >&2
        exit 1
    fi
fi

# ── Bump manifest.json (surgical — only the version line changes) ───────────
perl -pi -e 's/("version"\s*:\s*)"[^"]*"/${1}"'"$VERSION"'"/' manifest.json
python3 -c "import json; json.load(open('manifest.json'))" \
    || { echo "ERROR: manifest.json broke after version bump"; exit 1; }
echo "  manifest.json: $CURRENT_MANIFEST -> $VERSION"

# ── Build the zip ───────────────────────────────────────────────────────────
if $DO_BUILD; then
    echo
    bash scripts/build-extension.sh
fi

if ! $DO_COMMIT; then
    echo
    echo "  Skipped commit (--no-commit). To commit + tag manually:"
    echo "    (cd $REPO_DIR && git add extension/manifest.json && \\"
    echo "     git commit -m 'Browser ext: bump to v$VERSION' && \\"
    echo "     git push origin main)"
    exit 0
fi

# ── Commit + push the manifest bump to contexterGui ─────────────────────────
echo
(
    cd "$REPO_DIR"
    git add extension/manifest.json
    git commit -m "Browser ext: bump to v$VERSION" >/dev/null
    echo "  Committed: $(git rev-parse --short HEAD)  Browser ext: bump to v$VERSION"
)

SOURCE_SHA="$(cd "$REPO_DIR" && git rev-parse --short HEAD)"

if ! $DO_TAG_PUSH; then
    echo
    echo "  Skipped push (--no-tag-push). To push manually:"
    echo "    (cd $REPO_DIR && git push origin main)"
    echo "    gh release create $TAG '$ZIP' --repo $RELEASES_REPO --title '...' --notes '...'"
    exit 0
fi

(cd "$REPO_DIR" && git push origin main) >/dev/null
echo "  Pushed contexterGui main"

if ! $DO_UPLOAD; then
    echo
    echo "  Skipped GitHub release (--no-upload). To create manually:"
    echo "    gh release create $TAG '$ZIP' --repo $RELEASES_REPO \\"
    echo "      --title 'Spell Browser Companion Ext v$VERSION' --notes '...'"
    exit 0
fi

# ── Create the GitHub release on spell_binaries with the zip attached ───────
echo
NOTES_FILE="$(mktemp)"
trap 'rm -f "$NOTES_FILE"' EXIT

cat > "$NOTES_FILE" <<EOF
Spell browser companion extension v$VERSION.

This is the **companion** extension — it requires the Spell desktop app
to be installed on Mac/Windows. It pipes browser typing to the desktop
via native messaging (com.cognio.spell.bridge). Spell-check and grammar
checks render in the desktop Spell window.

## Install (for testing — not from Chrome Web Store)

1. Download \`Spell-Browser-Extension-${VERSION}.zip\` below
2. Unzip somewhere (Desktop is fine)
3. Open Chrome / Edge / Brave → \`chrome://extensions\`
4. Toggle **Developer mode** (top right)
5. Click **Load unpacked** and select the unzipped folder
6. Copy the extension's ID (32-char string under the card)
7. Open the native messaging host config in your editor:
   - Chrome: \`~/Library/Application Support/Google/Chrome/NativeMessagingHosts/com.cognio.spell.bridge.json\`
   - Edge:   \`~/Library/Application Support/Microsoft Edge/NativeMessagingHosts/com.cognio.spell.bridge.json\`
   Replace the existing \`allowed_origins\` entry with
   \`["chrome-extension://<YOUR_EXT_ID>/"]\`.
8. Reload the extension. Make sure the Spell desktop app is running.
9. Type in Gmail / Reddit / any textarea — the Spell desktop window
   should show suggestions.

## What's in this build (since last tag)

$(git -C "$REPO_DIR" log --pretty=format:'- %s' "${BASE_FOR_BUMP:+browser-ext-v}${BASE_FOR_BUMP:-}..HEAD" -- extension/ 2>/dev/null \
  | head -30 \
  || echo "- (no prior tag — see git log of extension/)")

Source commit: \`${SOURCE_SHA}\` in contexterGui main.

Built locally on macOS arm64.
EOF

DRAFT_FLAG=""
$DRAFT && DRAFT_FLAG="--draft"

# spell_binaries doesn't have the contexterGui source — the tag is just
# a release-coordination marker. Use --target to bind the tag to the
# default branch HEAD of the releases repo.
RELEASE_URL="$(gh release create "$TAG" "$ZIP" \
    --repo "$RELEASES_REPO" \
    --title "Spell Browser Companion Ext v$VERSION" \
    --notes-file "$NOTES_FILE" \
    $DRAFT_FLAG 2>&1 | grep -E '^https://' | head -1)"

echo "  Release published: $RELEASE_URL"
echo
echo "  Zip:        $ZIP"
echo "  Share with: $RELEASE_URL"
