// Auto-register the browser-companion native messaging host with Chrome /
// Edge / Brave on startup.  Runs every launch; idempotent (just overwrites
// the manifest + registry entry).  Without this the user has to manually
// edit `com.cognio.spell.bridge.json` to point at the installed
// native_bridge binary AND register the host with their browser — which
// is exactly the friction Petter wanted us to remove on Windows
// (reported 2026-05-19: "I have installed the browser extension already
// but seems like the connection is not getting properly made with
// browser extension").
//
// The companion extension's stable ID (`dcnbinagicnahihcgjfnlhepckfcpgob`)
// is derived deterministically from the RSA public key pinned in
// extension/manifest.json's `key` field — same key everywhere → same ID
// everywhere → safe to hard-code here.

use std::path::{Path, PathBuf};

const HOST_NAME: &str = "com.cognio.spell.bridge";
const EXTENSION_ID: &str = "dcnbinagicnahihcgjfnlhepckfcpgob";

/// Locate the native_bridge executable. On a real install it sits next to
/// the running Spell binary; in dev (cargo run) it's in the same target/
/// dir as acatts-rust.
fn find_native_bridge_path() -> Option<PathBuf> {
    let cur = std::env::current_exe().ok()?;
    let dir = cur.parent()?;
    let bridge_name = if cfg!(windows) { "native_bridge.exe" } else { "native_bridge" };
    let candidate = dir.join(bridge_name);
    if candidate.exists() { Some(candidate) } else { None }
}

/// Build the JSON manifest body Chrome expects for a native messaging host.
fn build_manifest_json(bridge_path: &Path) -> String {
    // JSON-escape backslashes (Windows paths use them and JSON treats
    // them as escape introducers).  The raw path can also contain spaces,
    // accented characters, etc. — none of those need escaping in JSON
    // string values, only \ and ".
    let mut escaped = String::new();
    for ch in bridge_path.to_string_lossy().chars() {
        match ch {
            '\\' => escaped.push_str("\\\\"),
            '"' => escaped.push_str("\\\""),
            c => escaped.push(c),
        }
    }
    format!(
        r#"{{
  "name": "{name}",
  "description": "Spell browser-companion native messaging host (auto-registered by Spell on startup)",
  "path": "{path}",
  "type": "stdio",
  "allowed_origins": [
    "chrome-extension://{ext}/"
  ]
}}
"#,
        name = HOST_NAME,
        path = escaped,
        ext = EXTENSION_ID,
    )
}

/// Top-level entry point. Logs results, never panics. Returns Ok(()) on
/// best-effort success; callers don't need to handle failure since the
/// browser companion is an optional feature — if registration fails the
/// user still gets desktop + Word working.
pub fn register_native_messaging_host_best_effort() {
    match register_inner() {
        Ok(()) => {}
        Err(e) => {
            crate::log!(
                "Native host: registration failed: {} (browser-companion extension may not work until manually configured)",
                e
            );
        }
    }
}

#[cfg(target_os = "macos")]
fn register_inner() -> std::io::Result<()> {
    let bridge_path = find_native_bridge_path().ok_or_else(|| {
        std::io::Error::new(
            std::io::ErrorKind::NotFound,
            "native_bridge binary not found next to Spell.app",
        )
    })?;
    let manifest = build_manifest_json(&bridge_path);

    let home = dirs::home_dir().ok_or_else(|| {
        std::io::Error::new(std::io::ErrorKind::NotFound, "no home dir")
    })?;

    // Chrome, Edge, Brave all use the same NativeMessagingHosts directory
    // layout inside their per-app support folder.  We write to each of
    // them so the extension works regardless of which browser the user
    // installed it in.
    let targets = [
        home.join("Library/Application Support/Google/Chrome/NativeMessagingHosts"),
        home.join("Library/Application Support/Microsoft Edge/NativeMessagingHosts"),
        home.join("Library/Application Support/BraveSoftware/Brave-Browser/NativeMessagingHosts"),
        home.join("Library/Application Support/Chromium/NativeMessagingHosts"),
    ];

    let mut wrote_any = false;
    for dir in &targets {
        // Best-effort per browser: if Chrome is installed but Edge isn't,
        // creating Edge's dir is harmless (Edge would pick it up on
        // install) but a write failure on one browser shouldn't block the
        // others.
        if let Err(e) = std::fs::create_dir_all(dir) {
            crate::log!("Native host: skip {} (mkdir: {})", dir.display(), e);
            continue;
        }
        let file = dir.join(format!("{}.json", HOST_NAME));
        match std::fs::write(&file, manifest.as_bytes()) {
            Ok(()) => {
                crate::log!("Native host: wrote manifest to {}", file.display());
                wrote_any = true;
            }
            Err(e) => crate::log!("Native host: write {} failed: {}", file.display(), e),
        }
    }
    if !wrote_any {
        return Err(std::io::Error::new(
            std::io::ErrorKind::Other,
            "no browser native-messaging-host dir writable",
        ));
    }
    crate::log!(
        "Native host: registered (Mac) → bridge={} ext={}",
        bridge_path.display(),
        EXTENSION_ID
    );
    Ok(())
}

#[cfg(target_os = "windows")]
fn register_inner() -> std::io::Result<()> {
    let bridge_path = find_native_bridge_path().ok_or_else(|| {
        std::io::Error::new(
            std::io::ErrorKind::NotFound,
            "native_bridge.exe not found next to Spell.exe",
        )
    })?;
    let manifest = build_manifest_json(&bridge_path);

    // Put the manifest next to the binary it points at.  The HKCU
    // registry entry then references this file's full path.  Writing
    // here means the manifest naturally updates whenever Velopack ships
    // a new version (Velopack restages everything in a new versioned
    // folder; we re-register at next startup with the new path).
    let manifest_path = bridge_path
        .parent()
        .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::Other, "no parent dir for bridge"))?
        .join(format!("{}.json", HOST_NAME));
    std::fs::write(&manifest_path, manifest.as_bytes())?;
    crate::log!(
        "Native host: wrote manifest to {}",
        manifest_path.display()
    );

    let manifest_str = manifest_path.to_string_lossy().into_owned();
    let browsers = [
        ("Chrome", r"Software\Google\Chrome\NativeMessagingHosts"),
        ("Edge", r"Software\Microsoft\Edge\NativeMessagingHosts"),
        ("Brave", r"Software\BraveSoftware\Brave-Browser\NativeMessagingHosts"),
        ("Chromium", r"Software\Chromium\NativeMessagingHosts"),
    ];

    let mut registered_any = false;
    for (label, prefix) in &browsers {
        let key_path = format!(r"HKCU\{}\{}", prefix, HOST_NAME);
        // Shell out to `reg add` rather than adding a winreg crate
        // dependency just for one call site. /ve sets the default value,
        // /f overwrites without prompting, /d is the data (manifest
        // path), REG_SZ is the type.
        let out = std::process::Command::new("reg")
            .args([
                "add",
                &key_path,
                "/ve",
                "/t",
                "REG_SZ",
                "/d",
                &manifest_str,
                "/f",
            ])
            .output();
        match out {
            Ok(o) if o.status.success() => {
                crate::log!("Native host: registered {} → {}", key_path, manifest_str);
                registered_any = true;
            }
            Ok(o) => crate::log!(
                "Native host: reg add failed for {} ({}): {}",
                label,
                o.status,
                String::from_utf8_lossy(&o.stderr).trim()
            ),
            Err(e) => crate::log!("Native host: reg add for {} errored: {}", label, e),
        }
    }
    if !registered_any {
        return Err(std::io::Error::new(
            std::io::ErrorKind::Other,
            "no browser registry entry could be created",
        ));
    }
    crate::log!(
        "Native host: registered (Windows) → bridge={} ext={}",
        bridge_path.display(),
        EXTENSION_ID
    );
    Ok(())
}

#[cfg(not(any(target_os = "macos", target_os = "windows")))]
fn register_inner() -> std::io::Result<()> {
    // Other platforms (Linux etc.) — not in scope.
    Ok(())
}
