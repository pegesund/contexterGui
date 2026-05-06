#!/usr/bin/env bash
# Build Spell.app for macOS, optionally codesign + notarize + DMG.
#
# Usage:
#   scripts/build-mac.sh [--no-sign] [--no-notarize] [--no-dmg] [--arch arm64|x86_64]
#
# Output:
#   dist/Spell.app
#   dist/Spell-osx-arm64.dmg   (if not --no-dmg)
set -euo pipefail

# ── Resolve project root regardless of where this is invoked from ────────────
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_DIR"

# ── Args ─────────────────────────────────────────────────────────────────────
SIGN=true
NOTARIZE=true
MAKE_DMG=true
ARCH="$(uname -m)"   # arm64 on Apple Silicon, x86_64 on Intel
VERSION="${SPELL_VERSION:-0.1.0}"

while [ $# -gt 0 ]; do
    case "$1" in
        --no-sign)      SIGN=false; NOTARIZE=false ;;
        --no-notarize)  NOTARIZE=false ;;
        --no-dmg)       MAKE_DMG=false ;;
        --arch)         shift; ARCH="$1" ;;
        --version)      shift; VERSION="$1" ;;
        -h|--help)
            sed -n '2,8p' "$0" | sed 's/^# \{0,1\}//'
            exit 0
            ;;
        *) echo "Unknown arg: $1"; exit 2 ;;
    esac
    shift
done

# ── Constants ────────────────────────────────────────────────────────────────
APP_NAME="Spell"
BUNDLE_ID="com.cognio.Spell"
SIGNING_IDENTITY="Developer ID Application: Cognio AS (LB6MH29HTB)"
NOTARY_PROFILE="Cognio-Notary"
ENTITLEMENTS="$SCRIPT_DIR/Spell.entitlements"
INFO_PLIST_TPL="$SCRIPT_DIR/Info.plist.template"
ICON="$PROJECT_DIR/assets/Spell.icns"

DIST="$PROJECT_DIR/dist"
APP="$DIST/${APP_NAME}.app"
CONTENTS="$APP/Contents"
MACOS="$CONTENTS/MacOS"
FRAMEWORKS="$CONTENTS/Frameworks"
RESOURCES="$CONTENTS/Resources"

# Homebrew paths for SWI-Prolog (resolve via brew so version isn't hardcoded)
SWIPL_HOME="$(brew --prefix swi-prolog)/lib/swipl"
SWIPL_DYLIB="$SWIPL_HOME/lib/${ARCH}-darwin/libswipl.dylib"
GMP_DYLIB="$(brew --prefix gmp)/lib/libgmp.10.dylib"
ONNX_DYLIB="$(brew --prefix onnxruntime)/lib/libonnxruntime.dylib"

# ── Helpers ──────────────────────────────────────────────────────────────────
step() { echo; echo "=== $* ==="; }

# Copy a dylib into Frameworks/, rewrite its install name to @rpath, and
# rewrite any @rpath/abs-path deps that we know about.
bundle_dylib() {
    local src="$1"
    local name="$(basename "$src")"
    local dest="$FRAMEWORKS/$name"
    cp -L "$src" "$dest"
    chmod 755 "$dest"
    install_name_tool -id "@rpath/$name" "$dest"
    # Rewrite known absolute deps to @rpath so the bundle is self-contained
    while IFS= read -r dep; do
        case "$dep" in
            /opt/homebrew/*|/usr/local/*)
                local depname="$(basename "$dep")"
                install_name_tool -change "$dep" "@rpath/$depname" "$dest" 2>/dev/null || true
                ;;
        esac
    done < <(otool -L "$dest" | tail -n +2 | awk '{print $1}')
}

retry_codesign() {
    local max=3 attempt=1
    while [ $attempt -le $max ]; do
        if codesign "$@"; then return 0; fi
        attempt=$((attempt + 1))
        [ $attempt -le $max ] && sleep $((10 * attempt))
    done
    return 1
}

# ── Clean previous build artifacts ───────────────────────────────────────────
# Always start fresh. We don't keep multiple DMGs around — the most recent
# build is the only one that matters (released versions live on GitHub).
step "Clean dist/"
rm -rf "$DIST/releases" "$APP"
mkdir -p "$DIST/releases"

# ── Preflight ────────────────────────────────────────────────────────────────
step "Preflight"
[ -f "$ICON" ] || { echo "ERROR: $ICON missing — run scripts/make-icns.sh first"; exit 1; }
[ -f "$SWIPL_DYLIB" ] || { echo "ERROR: SWI-Prolog libswipl not found at $SWIPL_DYLIB"; exit 1; }
[ -f "$GMP_DYLIB" ] || { echo "ERROR: libgmp not found at $GMP_DYLIB"; exit 1; }
[ -f "$ONNX_DYLIB" ] || { echo "ERROR: ONNX runtime not found at $ONNX_DYLIB — try: brew install onnxruntime"; exit 1; }
if $SIGN && ! security find-identity -v -p codesigning 2>/dev/null | grep -q "$SIGNING_IDENTITY"; then
    echo "ERROR: signing identity not in keychain: $SIGNING_IDENTITY"
    exit 1
fi
echo "  arch:           $ARCH"
echo "  version:        $VERSION"
echo "  swipl home:     $SWIPL_HOME"
echo "  sign:           $SIGN"
echo "  notarize:       $NOTARIZE"
echo "  dmg:            $MAKE_DMG"

# ── 1. Build release binary ──────────────────────────────────────────────────
step "Cargo build (release, $ARCH)"
case "$ARCH" in
    arm64)   RUST_TARGET="aarch64-apple-darwin" ;;
    x86_64)  RUST_TARGET="x86_64-apple-darwin" ;;
    *) echo "Unknown arch: $ARCH"; exit 2 ;;
esac
rustup target add "$RUST_TARGET" >/dev/null 2>&1 || true
cargo build --release --bin acatts-rust --target "$RUST_TARGET"
BIN_SRC="$PROJECT_DIR/target/$RUST_TARGET/release/acatts-rust"
[ -f "$BIN_SRC" ] || { echo "ERROR: built binary not found at $BIN_SRC"; exit 1; }

# ── 2. Assemble bundle skeleton ──────────────────────────────────────────────
step "Assemble Spell.app skeleton"
rm -rf "$APP"
mkdir -p "$MACOS" "$FRAMEWORKS" "$RESOURCES"

# Binary: install as Spell (not acatts-rust) so it matches CFBundleExecutable
cp "$BIN_SRC" "$MACOS/$APP_NAME"
chmod 755 "$MACOS/$APP_NAME"

# Add @rpath so the binary finds bundled dylibs at runtime
install_name_tool -add_rpath "@executable_path/../Frameworks" "$MACOS/$APP_NAME" 2>/dev/null || true

# Icon
cp "$ICON" "$RESOURCES/Spell.icns"

# Info.plist
sed -e "s/__VERSION__/$VERSION/g" -e "s/__ARCH__/$ARCH/g" "$INFO_PLIST_TPL" > "$CONTENTS/Info.plist"

# ── 3. Bundle SWI-Prolog (libs + home) ───────────────────────────────────────
step "Bundle SWI-Prolog"
bundle_dylib "$SWIPL_DYLIB"
# libswipl.dylib has a self-redirect to libswipl.10.dylib (same file). Provide both names.
ln -sf "libswipl.dylib" "$FRAMEWORKS/libswipl.10.dylib"

bundle_dylib "$GMP_DYLIB"

# Copy SWI's home dir (ABI, boot, library/, etc.) into Resources/swipl/
# nostos-cognio's swipl_checker.rs looks for it at <Frameworks>/../Resources/swipl/
SWIPL_BUNDLED_HOME="$RESOURCES/swipl"
mkdir -p "$SWIPL_BUNDLED_HOME"
for item in ABI LICENSE README.md app boot boot.prc cmake customize library swipl.home; do
    if [ -e "$SWIPL_HOME/$item" ]; then
        cp -RL "$SWIPL_HOME/$item" "$SWIPL_BUNDLED_HOME/"
    fi
done
# Skip 'bin', 'doc', 'demo', 'include' — not needed at runtime, saves ~30MB.
# 'lib' is the dylib dir; we've already bundled libswipl into Frameworks/

echo "  swipl home: $(du -sh "$SWIPL_BUNDLED_HOME" | cut -f1)"

# ── 4. Bundle ONNX runtime + recursive @rpath deps ──────────────────────────
step "Bundle ONNX runtime + transitive deps"
# ONNX Runtime depends on ~80 dylibs (libonnx, libonnx_proto, libprotobuf,
# libre2, libabsl_*, etc.). Use dylibbundler to walk them recursively and copy
# each into Frameworks/, rewriting install names to @rpath.
ONNX_REAL="$(readlink -f "$ONNX_DYLIB" 2>/dev/null || readlink "$ONNX_DYLIB" || echo "$ONNX_DYLIB")"
case "$ONNX_REAL" in /*) ;; *) ONNX_REAL="$(dirname "$ONNX_DYLIB")/$ONNX_REAL" ;; esac
cp -L "$ONNX_REAL" "$FRAMEWORKS/$(basename "$ONNX_REAL")"
chmod 755 "$FRAMEWORKS/$(basename "$ONNX_REAL")"

dylibbundler \
    -of -b \
    -x "$FRAMEWORKS/$(basename "$ONNX_REAL")" \
    -d "$FRAMEWORKS" \
    -p "@rpath/" \
    -s /opt/homebrew/lib \
    -s /opt/homebrew/Cellar \
    2>&1 | tail -20

# Provide the unversioned name (libonnxruntime.dylib) the Rust code looks for.
if [ ! -e "$FRAMEWORKS/libonnxruntime.dylib" ]; then
    ln -sf "$(basename "$ONNX_REAL")" "$FRAMEWORKS/libonnxruntime.dylib"
fi
echo "  Frameworks/ now $(du -sh "$FRAMEWORKS" | cut -f1) ($(ls "$FRAMEWORKS" | wc -l | tr -d ' ') files)"

# ── 5. Bundle resources ──────────────────────────────────────────────────────
step "Bundle resources"
# Fonts
cp "$PROJECT_DIR/fonts/OpenSans-Regular.ttf" "$RESOURCES/"

# Word add-in static files (server reads them from Contents/Resources/word-addin/)
# Only PUBLIC files are bundled — manifest + UI assets that Word fetches over
# the local HTTPS connection. The TLS cert + key are NOT bundled; the
# first-launch wizard (src/setup/word_addin_setup.rs) generates a per-user
# CA + leaf cert at ~/Library/Application Support/Spell/word-addin-certs/
# and adds the CA to the system keychain via a graphical sudo prompt.
#
# CRITICAL: do NOT add fullchain.pem or key.pem to the list below. Bundling a
# private key in a public installer would let anyone who downloads the .dmg
# MITM every Spell user's localhost traffic.
mkdir -p "$RESOURCES/word-addin"
for f in manifest.xml taskpane.html taskpane.js commands.html commands.js \
         icon-32.png icon-64.png icon-80.png; do
    if [ -f "$PROJECT_DIR/word-addin/$f" ]; then
        cp "$PROJECT_DIR/word-addin/$f" "$RESOURCES/word-addin/"
    fi
done

# ── 6. Sign (inside-out) ─────────────────────────────────────────────────────
if $SIGN; then
    step "Code sign (inside-out)"
    # Sign every dylib first
    find "$FRAMEWORKS" -type f \( -name "*.dylib" -o -name "*.so" \) -print0 \
        | while IFS= read -r -d '' f; do
            retry_codesign --force --timestamp --options runtime \
                --sign "$SIGNING_IDENTITY" "$f"
        done
    # Sign the main bundle. --deep is intentionally omitted because we already
    # signed dylibs above; --deep would re-sign with the wrong entitlements.
    retry_codesign --force --timestamp --options runtime \
        --sign "$SIGNING_IDENTITY" \
        --entitlements "$ENTITLEMENTS" \
        "$APP"
    echo "  Verifying signature…"
    codesign --verify --strict --verbose=2 "$APP" 2>&1 | tail -3
else
    echo "  Skipping signing (--no-sign)"
fi

# ── 7. Notarize ──────────────────────────────────────────────────────────────
# Skipped when we're going to vpk-pack (MAKE_DMG=true): vpk replaces the .app
# entirely with a Velopack-managed bundle, so notarizing this version of the
# .app is wasted work — vpk's own --notaryProfile handles it for the
# replacement, and step 8 notarizes the final DMG. Total release-path
# notarizations: 2 (vpk + DMG), matching Concentrate. Keeping notarization
# only for the --no-dmg path where this .app IS the artifact users get.
if $NOTARIZE && ! $MAKE_DMG; then
    step "Notarize Spell.app"
    SUBMIT_ZIP="$(mktemp -d)/Spell.zip"
    ditto -c -k --keepParent "$APP" "$SUBMIT_ZIP"
    if xcrun notarytool submit "$SUBMIT_ZIP" \
        --keychain-profile "$NOTARY_PROFILE" --wait; then
        echo "  Stapling…"
        xcrun stapler staple "$APP"
        echo "  Verifying staple…"
        xcrun stapler validate "$APP" || true
    else
        echo "  WARNING: notarization failed — Gatekeeper may block this build"
    fi
    rm -f "$SUBMIT_ZIP"
elif $NOTARIZE && $MAKE_DMG; then
    echo "  Skipping standalone-app notarization (deferred to vpk + DMG steps)"
else
    echo "  Skipping notarization (--no-notarize)"
fi

# ── 7.5 Velopack pack (auto-update artifacts) ────────────────────────────────
# Runs after the .app is fully signed + notarized. vpk pack:
#   - takes our existing .app as input (preserves Frameworks/, Resources/, etc)
#   - injects the `UpdateMac` bootstrapper into Contents/MacOS/
#   - emits a .nupkg + RELEASES manifest to dist/releases/velopack/
#   - emits a Portable.zip whose .app IS the Velopack-managed bundle we
#     then use as the source for create-dmg below — so the DMG users
#     install from website/drive ALSO gets the Velopack hooks. Without
#     this, first-install users would forever stay outside auto-update.
#
# Mirrors ConcentrateDotNet/build.sh's vpk-pack-then-DMG-from-output flow.
# Only runs when we're building a DMG (the release path); --no-dmg dev
# builds skip this step and ship the unsigned-by-vpk .app directly.
if $MAKE_DMG; then
VELO_OUT="$DIST/releases/velopack"
mkdir -p "$VELO_OUT"
step "Velopack pack ($ARCH)"

VPK_ARGS=(
    --packId "$APP_NAME"
    --packVersion "$VERSION"
    --packDir "$APP"
    --mainExe "$APP_NAME"
    --channel "osx-${ARCH}"
    --outputDir "$VELO_OUT"
)
if $SIGN; then
    VPK_ARGS+=(--signAppIdentity "$SIGNING_IDENTITY" --signEntitlements "$ENTITLEMENTS")
fi
if $NOTARIZE; then
    VPK_ARGS+=(--notaryProfile "$NOTARY_PROFILE")
fi
vpk pack "${VPK_ARGS[@]}"

# Replace our hand-built .app with the Velopack-managed one extracted
# from the Portable.zip. The Velopack version has UpdateMac + sq.version
# inside Contents/MacOS/ — required for in-place auto-updates to work.
PORTABLE_ZIP="$VELO_OUT/${APP_NAME}-osx-${ARCH}-Portable.zip"
[ -f "$PORTABLE_ZIP" ] || { echo "ERROR: vpk pack didn't produce $PORTABLE_ZIP"; exit 1; }
rm -rf "$APP"
unzip -q "$PORTABLE_ZIP" -d "$DIST"
rm -rf "$DIST/__MACOSX"
[ -d "$APP" ] || { echo "ERROR: Velopack .app not at expected path $APP"; exit 1; }
echo "  Substituted Velopack-managed .app into $APP"

# Drop artifacts we don't ship: the Setup.pkg (we distribute via DMG)
# and the Portable.zip (auto-update reads .nupkg, not the .zip).
#
# DO NOT delete releases.osx-${ARCH}.json — that's exactly the file the
# Rust velopack GithubSource downloads from GitHub Releases when
# checking for updates (`releases.{channel}.json` per
# velopack/sources/mod.rs::get_git_release_feed). Without it the check
# returns RemoteIsEmpty even though the .nupkg + RELEASES-osx-arm64 are
# present. RELEASES-osx-arm64 alone is for legacy compatibility; the
# JSON is the authoritative manifest.
#
# assets.osx-${ARCH}.json stays — vpk's `upload` command reads it (and
# we don't actually use that command, but keeping it here is harmless
# and matches Concentrate's layout). Patch it to drop the Installer
# entry so a future `vpk upload github` wouldn't try to re-upload the
# deleted .pkg. Mirrors ConcentrateDotNet/build.sh:973-980.
rm -f "$VELO_OUT/${APP_NAME}-osx-${ARCH}-Setup.pkg" \
      "$VELO_OUT/${APP_NAME}-osx-${ARCH}-Portable.zip"
python3 -c "
import json, glob, sys
for f in glob.glob(sys.argv[1] + '/assets.*.json'):
    with open(f) as fh: data = json.load(fh)
    data = [a for a in data if a.get('Type') != 'Installer']
    with open(f, 'w') as fh: json.dump(data, fh)
    print(f'  Patched {f} (dropped Installer entry)')
" "$VELO_OUT"
echo "  Velopack artifacts: $(ls "$VELO_OUT")"
fi  # end MAKE_DMG gate around vpk pack

# ── 8. Build DMG ─────────────────────────────────────────────────────────────
if $MAKE_DMG; then
    step "Build DMG"
    mkdir -p "$DIST/releases"
    DMG="$DIST/releases/Spell-osx-${ARCH}-${VERSION}.dmg"
    rm -f "$DMG"
    [ -d "/Volumes/$APP_NAME" ] && hdiutil detach "/Volumes/$APP_NAME" -force 2>/dev/null || true

    # The DMG window is 600x400 logical pixels. Background image is 1200x800
    # pre-set to 144 DPI so macOS treats it as 2x retina (sharp on HiDPI).
    BG_SRC="$PROJECT_DIR/assets/dmg_background.png"
    DMG_OPTS=(
        --volname "$APP_NAME"
        --window-pos 200 120
        --window-size 600 400
        --icon-size 128
        --icon "${APP_NAME}.app" 140 200
        --app-drop-link 460 200
        --hide-extension "${APP_NAME}.app"
        --no-internet-enable
    )
    if [ -f "$BG_SRC" ]; then
        BG_READY="$DIST/releases/dmg_bg_ready.png"
        cp "$BG_SRC" "$BG_READY"
        sips -s dpiWidth 144 -s dpiHeight 144 "$BG_READY" >/dev/null 2>&1
        DMG_OPTS=("${DMG_OPTS[@]}" --background "$BG_READY")
    else
        echo "  WARNING: $BG_SRC missing — DMG will have no background"
    fi
    create-dmg "${DMG_OPTS[@]}" "$DMG" "$APP" || true
    [ -f "$BG_READY" ] && rm -f "$BG_READY" || true
    [ -f "$DMG" ] || { echo "ERROR: create-dmg failed"; exit 1; }

    if $SIGN; then
        retry_codesign --force --sign "$SIGNING_IDENTITY" "$DMG"
        echo "  DMG signed"
    fi

    if $NOTARIZE; then
        echo "  Notarizing DMG…"
        if xcrun notarytool submit "$DMG" \
            --keychain-profile "$NOTARY_PROFILE" --wait; then
            xcrun stapler staple "$DMG"
            echo "  DMG notarized & stapled"
        else
            echo "  WARNING: DMG notarization failed"
        fi
    fi
    echo "  DMG: $DMG ($(du -h "$DMG" | cut -f1))"

    # The DMG is the only artifact we ship. Discard the intermediate .app
    # so dist/ stays clean — to launch Spell, install it from the DMG into
    # /Applications/. (Kept only when --no-dmg is passed: in that case the
    # .app IS the explicit goal, e.g. for direct exec testing.)
    rm -rf "$APP"
fi

step "Done"
if $MAKE_DMG; then
    echo "  DMG: $DMG ($(du -h "$DMG" | cut -f1))"
else
    echo "  App: $APP ($(du -sh "$APP" | cut -f1))"
fi
echo
