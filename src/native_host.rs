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

/// Stop Chrome/Edge-spawned native hosts that are running from this install.
///
/// Velopack updates and uninstalls need to replace/delete
/// AppData\Local\Spell\current. A live native_bridge.exe keeps its own file
/// locked, so the cleanup must happen before Velopack starts moving files.
pub fn stop_native_bridge_processes_best_effort() {
    #[cfg(target_os = "windows")]
    match stop_native_bridge_processes_inner() {
        Ok(killed) => {
            if killed > 0 {
                crate::log!("Native host: stopped {} native_bridge.exe process(es)", killed);
            }
        }
        Err(e) => crate::log!("Native host: stop native_bridge.exe failed: {}", e),
    }
}

/// Remove per-user browser native-host registration during uninstall.
pub fn unregister_native_messaging_host_best_effort() {
    #[cfg(target_os = "windows")]
    match unregister_inner() {
        Ok(removed) => {
            crate::log!("Native host: unregistered {} browser host key(s)", removed);
        }
        Err(e) => crate::log!("Native host: unregister failed: {}", e),
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
        let subkey = format!(r"{}\{}", prefix, HOST_NAME);
        // Use the Win32 registry API directly instead of spawning
        // `reg add` as a child process.  The previous shell-out path
        // produced an intermittent dialog on the user's Windows box:
        //   reg.exe - Application Error
        //   The application was unable to start correctly (0xc0000142).
        // 0xc0000142 is STATUS_DLL_INIT_FAILED — typically caused by
        // app-compat layers, security software, or environment-block
        // weirdness propagated from the parent process when launching
        // children. Reported 2026-05-19 ("I saw this error on windows
        // app 2 times when I opened it today and then it did started").
        //
        // RegCreateKeyExW + RegSetValueExW are in-process API calls.
        // No process spawn, no DLL-init dance for an external binary,
        // no possible UI dialog.
        match write_registry_default_string(&subkey, &manifest_str) {
            Ok(()) => {
                crate::log!("Native host: registered HKCU\\{} → {}", subkey, manifest_str);
                registered_any = true;
            }
            Err(e) => crate::log!(
                "Native host: registry write failed for {} (HKCU\\{}): {}",
                label,
                subkey,
                e
            ),
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

#[cfg(target_os = "windows")]
fn unregister_inner() -> std::io::Result<usize> {
    let browsers = [
        r"Software\Google\Chrome\NativeMessagingHosts",
        r"Software\Microsoft\Edge\NativeMessagingHosts",
        r"Software\BraveSoftware\Brave-Browser\NativeMessagingHosts",
        r"Software\Chromium\NativeMessagingHosts",
    ];

    let mut removed = 0usize;
    for prefix in &browsers {
        let subkey = format!(r"{}\{}", prefix, HOST_NAME);
        match delete_registry_tree(&subkey) {
            Ok(true) => {
                crate::log!("Native host: deleted HKCU\\{}", subkey);
                removed += 1;
            }
            Ok(false) => {}
            Err(e) => crate::log!("Native host: delete HKCU\\{} failed: {}", subkey, e),
        }
    }
    Ok(removed)
}

#[cfg(target_os = "windows")]
fn delete_registry_tree(subkey: &str) -> std::io::Result<bool> {
    use std::iter::once;
    use std::os::windows::ffi::OsStrExt;
    use windows::Win32::Foundation::{ERROR_FILE_NOT_FOUND, ERROR_PATH_NOT_FOUND};
    use windows::Win32::System::Registry::{RegDeleteTreeW, HKEY_CURRENT_USER};
    use windows::core::PCWSTR;

    let subkey_w: Vec<u16> = std::ffi::OsStr::new(subkey).encode_wide().chain(once(0)).collect();
    let status = unsafe {
        RegDeleteTreeW(HKEY_CURRENT_USER, PCWSTR::from_raw(subkey_w.as_ptr()))
    };
    if status == ERROR_FILE_NOT_FOUND || status == ERROR_PATH_NOT_FOUND {
        Ok(false)
    } else if status.is_err() {
        Err(std::io::Error::from_raw_os_error(status.0 as i32))
    } else {
        Ok(true)
    }
}

#[cfg(target_os = "windows")]
fn stop_native_bridge_processes_inner() -> std::io::Result<usize> {
    use windows::Win32::Foundation::CloseHandle;
    use windows::Win32::System::Diagnostics::ToolHelp::{
        CreateToolhelp32Snapshot, Process32FirstW, Process32NextW, PROCESSENTRY32W,
        TH32CS_SNAPPROCESS,
    };

    let target = find_native_bridge_path().ok_or_else(|| {
        std::io::Error::new(
            std::io::ErrorKind::NotFound,
            "native_bridge.exe not found next to current executable",
        )
    })?;

    let snapshot = unsafe { CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) }
        .map_err(|e| std::io::Error::from_raw_os_error(e.code().0))?;
    let mut entry = PROCESSENTRY32W::default();
    entry.dwSize = std::mem::size_of::<PROCESSENTRY32W>() as u32;

    let mut killed = 0usize;
    let mut has_entry = unsafe { Process32FirstW(snapshot, &mut entry).is_ok() };
    while has_entry {
        if process_entry_name(&entry).eq_ignore_ascii_case("native_bridge.exe")
            && entry.th32ProcessID != std::process::id()
        {
            if stop_process_if_image_matches(entry.th32ProcessID, &target)? {
                killed += 1;
            }
        }
        has_entry = unsafe { Process32NextW(snapshot, &mut entry).is_ok() };
    }

    let _ = unsafe { CloseHandle(snapshot) };
    Ok(killed)
}

#[cfg(target_os = "windows")]
fn stop_process_if_image_matches(pid: u32, target: &Path) -> std::io::Result<bool> {
    use windows::Win32::Foundation::CloseHandle;
    use windows::Win32::System::Threading::{
        OpenProcess, QueryFullProcessImageNameW, TerminateProcess, PROCESS_NAME_FORMAT,
        PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_TERMINATE,
    };
    use windows::core::PWSTR;

    let access = windows::Win32::System::Threading::PROCESS_ACCESS_RIGHTS(
        PROCESS_QUERY_LIMITED_INFORMATION.0 | PROCESS_TERMINATE.0,
    );
    let handle = match unsafe { OpenProcess(access, false, pid) } {
        Ok(handle) => handle,
        Err(_) => return Ok(false),
    };

    let mut buf = [0u16; 32768];
    let mut len = buf.len() as u32;
    let image = if unsafe {
        QueryFullProcessImageNameW(
            handle,
            PROCESS_NAME_FORMAT(0),
            PWSTR(buf.as_mut_ptr()),
            &mut len,
        )
    }
    .is_ok()
    {
        Some(PathBuf::from(String::from_utf16_lossy(&buf[..len as usize])))
    } else {
        None
    };

    let should_stop = image
        .as_ref()
        .map(|path| path_eq_windows(path, target))
        .unwrap_or(false);
    if should_stop {
        let _ = unsafe { TerminateProcess(handle, 0) };
    }
    let _ = unsafe { CloseHandle(handle) };
    Ok(should_stop)
}

#[cfg(target_os = "windows")]
fn process_entry_name(entry: &windows::Win32::System::Diagnostics::ToolHelp::PROCESSENTRY32W) -> String {
    let len = entry
        .szExeFile
        .iter()
        .position(|&c| c == 0)
        .unwrap_or(entry.szExeFile.len());
    String::from_utf16_lossy(&entry.szExeFile[..len])
}

#[cfg(target_os = "windows")]
fn path_eq_windows(a: &Path, b: &Path) -> bool {
    fn clean(path: &Path) -> String {
        path.to_string_lossy()
            .trim_start_matches(r"\\?\")
            .replace('/', "\\")
            .to_ascii_lowercase()
    }
    clean(a) == clean(b)
}

/// Set the default (unnamed) value of HKCU\<subkey> to a UTF-16 string.
/// Creates the key if it doesn't exist. Idempotent.
#[cfg(target_os = "windows")]
fn write_registry_default_string(subkey: &str, value: &str) -> std::io::Result<()> {
    use std::iter::once;
    use std::os::windows::ffi::OsStrExt;
    use windows::Win32::Foundation::WIN32_ERROR;
    use windows::Win32::System::Registry::{
        RegCloseKey, RegCreateKeyExW, RegSetValueExW, HKEY, HKEY_CURRENT_USER,
        KEY_SET_VALUE, REG_OPTION_NON_VOLATILE, REG_SZ,
    };
    use windows::core::PCWSTR;

    let subkey_w: Vec<u16> = std::ffi::OsStr::new(subkey).encode_wide().chain(once(0)).collect();
    let mut value_w: Vec<u16> = std::ffi::OsStr::new(value).encode_wide().chain(once(0)).collect();

    unsafe {
        let mut hkey = HKEY::default();
        // windows 0.61 signature: `reserved` is `Option<u32>` (the Win32
        // docs say "must be zero", so we pass None — equivalent to a NULL
        // reserved slot). The earlier comment claiming reserved was a bare
        // `u32` was wrong and caused E0308 / E0432 on Windows CI; Mac dev
        // can't catch this because the whole fn is #[cfg(target_os = "windows")].
        let create_status: WIN32_ERROR = RegCreateKeyExW(
            HKEY_CURRENT_USER,
            PCWSTR::from_raw(subkey_w.as_ptr()),
            None,
            PCWSTR::null(),
            REG_OPTION_NON_VOLATILE,
            KEY_SET_VALUE,
            None,
            &mut hkey,
            None,
        );
        if create_status.is_err() {
            return Err(std::io::Error::from_raw_os_error(create_status.0 as i32));
        }

        // RegSetValueExW expects the byte length INCLUDING the trailing
        // NUL terminator for REG_SZ values, which is exactly what
        // value_w.len() * 2 gives us (encode_wide().chain(once(0))).
        let bytes_len = value_w.len() * std::mem::size_of::<u16>();
        let value_bytes = std::slice::from_raw_parts(
            value_w.as_mut_ptr() as *const u8,
            bytes_len,
        );
        // windows 0.61 signature: `reserved` is `Option<u32>` here too.
        let set_status: WIN32_ERROR =
            RegSetValueExW(hkey, PCWSTR::null(), None, REG_SZ, Some(value_bytes));
        let _ = RegCloseKey(hkey);
        if set_status.is_err() {
            return Err(std::io::Error::from_raw_os_error(set_status.0 as i32));
        }
    }
    Ok(())
}

#[cfg(not(any(target_os = "macos", target_os = "windows")))]
fn register_inner() -> std::io::Result<()> {
    // Other platforms (Linux etc.) — not in scope.
    Ok(())
}
