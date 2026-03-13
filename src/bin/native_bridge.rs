/// NorskTale Native Messaging Host
///
/// Chrome/Edge launches this binary and communicates via stdin/stdout
/// using length-prefixed JSON messages.
///
/// Two threads:
/// - Main thread: reads stdin messages, writes data file, sends ack
/// - Reply thread: polls for reply file every 50ms, sends to extension

use std::io::{self, Read, Write};
use std::path::PathBuf;
use std::sync::{Arc, Mutex};
use std::time::Duration;

fn data_path() -> PathBuf {
    std::env::temp_dir().join("norsktale-browser.json")
}

fn reply_path() -> PathBuf {
    std::env::temp_dir().join("norsktale-browser-reply.json")
}

fn log_path() -> PathBuf {
    std::env::temp_dir().join("norsktale-native-bridge.log")
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
    loop {
        match read_message() {
            Ok(msg) => {
                let msg_str = String::from_utf8_lossy(&msg);

                // Handle keepalive pings (just ack, don't write to file)
                if msg_str.contains("\"type\":\"keepalive\"") {
                    let ack = br#"{"status":"ok"}"#;
                    if let Ok(mut out) = stdout.lock() {
                        if write_message_locked(&mut *out, ack).is_err() { break; }
                    }
                    continue;
                }

                // Handle log messages from content script
                if msg_str.contains("\"type\":\"log\"") {
                    if let Some(i) = msg_str.find("\"message\":\"") {
                        let s = &msg_str[i+11..];
                        let end = s.find('"').unwrap_or(s.len());
                        log(&format!("JS: {}", &s[..end]));
                    }
                    // Don't write log messages to the data file
                    let ack = br#"{"status":"ok"}"#;
                    if let Ok(mut out) = stdout.lock() {
                        if write_message_locked(&mut *out, ack).is_err() { break; }
                    }
                    continue;
                }

                // Log what we receive: URL and first 80 chars of text
                let url = msg_str.find("\"url\":\"").map(|i| {
                    let s = &msg_str[i+7..];
                    s.find('"').map(|e| &s[..e]).unwrap_or("")
                }).unwrap_or("");
                let text_preview = msg_str.find("\"text\":\"").map(|i| {
                    let s = &msg_str[i+8..];
                    let end = s.find('"').unwrap_or(80).min(80);
                    &s[..end]
                }).unwrap_or("");
                log(&format!("RECV url={} text='{}'", url, text_preview));

                if let Err(e) = std::fs::write(data_path(), &msg) {
                    log(&format!("Failed to write data file: {}", e));
                }

                let ack = br#"{"status":"ok"}"#;
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
