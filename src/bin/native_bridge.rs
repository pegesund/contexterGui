/// Spell Native Messaging Host
///
/// Chrome/Edge launches this binary and communicates via stdin/stdout
/// using length-prefixed JSON messages.
///
/// Two threads:
/// - Main thread: reads stdin messages, writes data file, sends ack
/// - Reply thread: polls for reply file every 50ms, sends to extension

use serde_json::Value;
use std::io::{self, Read, Write};
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};
use std::time::Duration;

#[cfg(windows)]
#[link(name = "kernel32")]
unsafe extern "system" {
    fn MoveFileExW(
        lpExistingFileName: *const u16,
        lpNewFileName: *const u16,
        dwFlags: u32,
    ) -> i32;
}

fn data_path() -> PathBuf {
    std::env::temp_dir().join("spell-browser.json")
}

/// Sibling tmp file used to make writes replace the final payload cleanly.
/// Include the process id so two browser profiles/native hosts do not fight
/// over the same .tmp name while they both target spell-browser.json.
fn tmp_path_for(final_path: &Path) -> PathBuf {
    let mut name = final_path
        .file_name()
        .map(|s| s.to_os_string())
        .unwrap_or_else(|| "spell-browser.json".into());
    name.push(format!(".{}.tmp", std::process::id()));
    final_path.with_file_name(name)
}

fn reply_path() -> PathBuf {
    std::env::temp_dir().join("spell-browser-reply.json")
}

fn log_path() -> PathBuf {
    std::env::temp_dir().join("spell-native-bridge.log")
}

/// Replace `final_path` with `data` without ever exposing partial JSON.
///
/// The previous implementation called `std::fs::write(data_path(), data)`
/// directly. Native messaging payloads can be up to ~1 MB, and a writer
/// can be interrupted mid-write; if the desktop side read concurrently
/// it would see a truncated file and fail JSON parsing. Writing to a
/// sibling `.tmp` and then replacing the final file means the desktop
/// either sees the previous complete payload or the new complete payload.
fn write_data_atomic_to(final_path: &Path, data: &[u8]) -> io::Result<()> {
    let tmp = tmp_path_for(final_path);
    // Write to tmp first
    if let Err(e) = std::fs::write(&tmp, data) {
        // Best-effort cleanup; ignore rm errors so we surface the write error
        let _ = std::fs::remove_file(&tmp);
        return Err(e);
    }
    // Replace the final path. On failure leave no half-state behind.
    if let Err(e) = replace_file(&tmp, final_path) {
        let _ = std::fs::remove_file(&tmp);
        return Err(e);
    }
    Ok(())
}

fn write_data_atomic(data: &[u8]) -> io::Result<()> {
    write_data_atomic_to(&data_path(), data)
}

#[cfg(not(windows))]
fn replace_file(tmp: &Path, final_path: &Path) -> io::Result<()> {
    std::fs::rename(tmp, final_path)
}

#[cfg(windows)]
fn replace_file(tmp: &Path, final_path: &Path) -> io::Result<()> {
    use std::os::windows::ffi::OsStrExt;

    const MOVEFILE_REPLACE_EXISTING: u32 = 0x00000001;
    const MOVEFILE_WRITE_THROUGH: u32 = 0x00000008;

    let tmp_w: Vec<u16> = tmp.as_os_str().encode_wide().chain(std::iter::once(0)).collect();
    let final_w: Vec<u16> = final_path
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();

    let ok = unsafe {
        MoveFileExW(
            tmp_w.as_ptr(),
            final_w.as_ptr(),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH,
        )
    };
    if ok == 0 {
        Err(io::Error::last_os_error())
    } else {
        Ok(())
    }
}

fn log(msg: &str) {
    if let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true).append(true).open(log_path())
    {
        let _ = writeln!(f, "{}", msg);
    }
}

/// Read a native messaging frame: 4-byte little-endian length + JSON bytes
fn read_message() -> io::Result<Vec<u8>> {
    let mut len_buf = [0u8; 4];
    io::stdin().read_exact(&mut len_buf)?;
    let len = u32::from_le_bytes(len_buf) as usize;
    if len > 10_000_000 {
        return Err(io::Error::new(io::ErrorKind::InvalidData, "message too large"));
    }
    let mut buf = vec![0u8; len];
    io::stdin().read_exact(&mut buf)?;
    Ok(buf)
}

/// Write a native messaging frame (caller must hold the lock)
fn write_message_locked(out: &mut io::Stdout, data: &[u8]) -> io::Result<()> {
    let len = data.len() as u32;
    out.write_all(&len.to_le_bytes())?;
    out.write_all(data)?;
    out.flush()
}

fn main() {
    log("Native bridge started (threaded)");

    let stdout = Arc::new(Mutex::new(io::stdout()));
    let alive = Arc::new(std::sync::atomic::AtomicBool::new(true));

    // Immediately send any pending reply from a previous session
    // (e.g., Rust wrote reply while bridge was dead, keepalive reconnected us)
    {
        let reply = reply_path();
        if reply.exists() {
            if let Ok(data) = std::fs::read(&reply) {
                let _ = std::fs::remove_file(&reply);
                if !data.is_empty() {
                    log(&format!("Startup: sending pending reply: {}", String::from_utf8_lossy(&data)));
                    if let Ok(mut out) = stdout.lock() {
                        let _ = write_message_locked(&mut *out, &data);
                    }
                }
            }
        }
    }

    // Reply checker thread — polls for reply file every 50ms
    let stdout2 = stdout.clone();
    let alive2 = alive.clone();
    std::thread::spawn(move || {
        while alive2.load(std::sync::atomic::Ordering::Relaxed) {
            std::thread::sleep(Duration::from_millis(50));
            let reply = reply_path();
            if reply.exists() {
                if let Ok(data) = std::fs::read(&reply) {
                    let _ = std::fs::remove_file(&reply);
                    if !data.is_empty() {
                        log(&format!("Reply thread sending: {}", String::from_utf8_lossy(&data)));
                        if let Ok(mut out) = stdout2.lock() {
                            if write_message_locked(&mut *out, &data).is_err() {
                                break;
                            }
                        }
                    }
                }
            }
        }
    });

    // Main thread: read stdin, write data file, send ack
    let ack: &[u8] = br#"{"status":"ok"}"#;
    loop {
        match read_message() {
            Ok(msg) => {
                // Parse the JSON properly instead of substring-matching the
                // raw bytes. Previously `msg_str.contains("\"type\":\"keepalive\"")`
                // could be fooled by a user's typed text that happened to
                // contain the literal string — same for `"type":"log"` and
                // the URL/text extraction. Parsing once and reading fields
                // through serde_json::Value is both correct and not much
                // slower (payloads are small JSON objects).
                let parsed: Option<Value> = serde_json::from_slice(&msg).ok();
                let msg_type = parsed.as_ref()
                    .and_then(|v| v.get("type"))
                    .and_then(|v| v.as_str())
                    .unwrap_or("");

                match msg_type {
                    "keepalive" => {
                        // Keepalive pings: just ack, don't write to file.
                        if let Ok(mut out) = stdout.lock() {
                            if write_message_locked(&mut *out, ack).is_err() { break; }
                        }
                        continue;
                    }
                    "log" => {
                        // Log messages from the content script. Pull the
                        // message field through serde so embedded escapes
                        // (`\"`, `\n`) survive.
                        if let Some(s) = parsed.as_ref()
                            .and_then(|v| v.get("message"))
                            .and_then(|v| v.as_str())
                        {
                            log(&format!("JS: {}", s));
                        }
                        if let Ok(mut out) = stdout.lock() {
                            if write_message_locked(&mut *out, ack).is_err() { break; }
                        }
                        continue;
                    }
                    _ => {}
                }

                // Regular payload (text update from a content script).
                // Log preview using parsed fields, then atomically replace
                // the data file so the desktop side never sees a partial
                // write.
                let url = parsed.as_ref()
                    .and_then(|v| v.get("url"))
                    .and_then(|v| v.as_str())
                    .unwrap_or("");
                let text_preview: String = parsed.as_ref()
                    .and_then(|v| v.get("text"))
                    .and_then(|v| v.as_str())
                    .map(|s| s.chars().take(80).collect())
                    .unwrap_or_default();
                log(&format!("RECV url={} text='{}'", url, text_preview));

                if let Err(e) = write_data_atomic(&msg) {
                    log(&format!("Failed to write data file: {}", e));
                }

                if let Ok(mut out) = stdout.lock() {
                    if write_message_locked(&mut *out, ack).is_err() {
                        break;
                    }
                }
            }
            Err(e) => {
                log(&format!("Read error (extension closed?): {}", e));
                break;
            }
        }
    }

    alive.store(false, std::sync::atomic::Ordering::Relaxed);
    log("Native bridge exiting");
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn write_data_atomic_to_overwrites_existing_file() {
        let final_path = std::env::temp_dir().join(format!(
            "spell-native-bridge-test-{}.json",
            std::process::id()
        ));
        let _ = std::fs::remove_file(&final_path);
        std::fs::write(&final_path, br#"{"old":true}"#).unwrap();

        write_data_atomic_to(&final_path, br#"{"new":true}"#).unwrap();

        let actual = std::fs::read_to_string(&final_path).unwrap();
        assert_eq!(actual, r#"{"new":true}"#);
        let _ = std::fs::remove_file(&final_path);
    }
}
