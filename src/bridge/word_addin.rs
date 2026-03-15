//! Word Add-in bridge via localhost HTTP.
//!
//! A Word Add-in (JavaScript) runs inside Word's process and POSTs
//! text context to a tiny HTTP server on localhost. This bridge
//! implements TextBridge by reading cached context from those POSTs.
//!
//! Architecture:
//!   Word Add-in → POST /context → HTTP thread → Arc<Mutex<CursorContext>>
//!   Rust main   → read_context() → reads cache (instant, never blocks)
//!   Rust main   → replace_word() → pushes to reply queue
//!   Word Add-in → GET /reply → pops from reply queue → applies in Word

use super::{CursorContext, TextBridge};
use std::io::{BufRead, BufReader, Read, Write};
use std::net::TcpListener;
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

const PORT: u16 = 52525;

pub struct WordAddinBridge {
    cached_context: Arc<Mutex<Option<(CursorContext, Instant)>>>,
    reply_queue: Arc<Mutex<Vec<String>>>,
}

impl WordAddinBridge {
    /// Start the HTTP server and return the bridge.
    /// Always succeeds — the add-in connects later.
    pub fn new() -> Self {
        let cached_context: Arc<Mutex<Option<(CursorContext, Instant)>>> =
            Arc::new(Mutex::new(None));
        let reply_queue: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));

        let ctx_clone = Arc::clone(&cached_context);
        let reply_clone = Arc::clone(&reply_queue);

        std::thread::Builder::new()
            .name("word-addin-http".into())
            .spawn(move || {
                let listener = match TcpListener::bind(format!("127.0.0.1:{}", PORT)) {
                    Ok(l) => {
                        eprintln!("Word Add-in HTTP bridge listening on port {}", PORT);
                        l
                    }
                    Err(e) => {
                        eprintln!("Word Add-in: failed to bind port {}: {}", PORT, e);
                        return;
                    }
                };

                for stream in listener.incoming() {
                    if let Ok(stream) = stream {
                        let ctx = Arc::clone(&ctx_clone);
                        let reply = Arc::clone(&reply_clone);
                        // Handle each request on a short-lived thread
                        // (Word add-in sends ~2-10 requests/sec, this is fine)
                        std::thread::spawn(move || {
                            handle_request(stream, &ctx, &reply);
                        });
                    }
                }
            })
            .expect("Failed to spawn Word Add-in HTTP server");

        WordAddinBridge {
            cached_context,
            reply_queue,
        }
    }

    fn push_reply(&self, json: String) {
        if let Ok(mut q) = self.reply_queue.lock() {
            q.push(json);
        }
    }
}

impl TextBridge for WordAddinBridge {
    fn name(&self) -> &str {
        "Word Add-in"
    }

    fn is_available(&self) -> bool {
        if let Ok(lock) = self.cached_context.lock() {
            lock.as_ref()
                .map(|(_, ts)| ts.elapsed() < Duration::from_secs(5))
                .unwrap_or(false)
        } else {
            false
        }
    }

    fn read_context(&self) -> Option<CursorContext> {
        if let Ok(lock) = self.cached_context.lock() {
            if let Some((ctx, ts)) = lock.as_ref() {
                if ts.elapsed() < Duration::from_secs(5) {
                    return Some(ctx.clone());
                }
            }
        }
        None
    }

    fn replace_word(&self, new_text: &str) -> bool {
        let json = format!(
            r#"{{"action":"replaceWord","text":"{}"}}"#,
            escape_json(new_text)
        );
        self.push_reply(json);
        true
    }

    fn find_and_replace(&self, find: &str, replace: &str) -> bool {
        let json = format!(
            r#"{{"action":"replace","expected":"{}","text":"{}"}}"#,
            escape_json(find),
            escape_json(replace)
        );
        self.push_reply(json);
        true
    }

    fn find_and_replace_in_context(&self, find: &str, replace: &str, _context: &str) -> bool {
        self.find_and_replace(find, replace)
    }

    fn find_and_replace_in_context_at(
        &self,
        find: &str,
        replace: &str,
        _context: &str,
        char_offset: usize,
    ) -> bool {
        let json = format!(
            r#"{{"action":"replace","expected":"{}","text":"{}","offset":{}}}"#,
            escape_json(find),
            escape_json(replace),
            char_offset
        );
        self.push_reply(json);
        true
    }

    fn read_full_document(&self) -> Option<String> {
        // The add-in sends before+after text (up to 2000 chars each).
        // Combine them as approximate full doc.
        if let Ok(lock) = self.cached_context.lock() {
            if let Some((ctx, _)) = lock.as_ref() {
                if !ctx.sentence.is_empty() {
                    return Some(ctx.sentence.clone());
                }
            }
        }
        None
    }

    fn mark_error_underline(&self, char_start: usize, char_end: usize) -> bool {
        let json = format!(
            r#"{{"action":"underline","start":{},"end":{}}}"#,
            char_start, char_end
        );
        self.push_reply(json);
        true
    }

    fn clear_error_underline(&self, char_start: usize, char_end: usize) -> bool {
        let json = format!(
            r#"{{"action":"clearUnderline","start":{},"end":{}}}"#,
            char_start, char_end
        );
        self.push_reply(json);
        true
    }

    fn clear_all_error_underlines(&self) -> bool {
        self.push_reply(r#"{"action":"clearAllUnderlines"}"#.to_string());
        true
    }
}

// ── HTTP handling ──

fn handle_request(
    mut stream: std::net::TcpStream,
    cached_context: &Arc<Mutex<Option<(CursorContext, Instant)>>>,
    reply_queue: &Arc<Mutex<Vec<String>>>,
) {
    let _ = stream.set_read_timeout(Some(Duration::from_secs(5)));

    let mut reader = BufReader::new(&stream);

    // Read request line
    let mut request_line = String::new();
    if reader.read_line(&mut request_line).is_err() {
        return;
    }

    // Parse method and path
    let parts: Vec<&str> = request_line.split_whitespace().collect();
    if parts.len() < 2 {
        return;
    }
    let method = parts[0];
    let path = parts[1];

    // Read headers to get Content-Length
    let mut content_length: usize = 0;
    loop {
        let mut header = String::new();
        if reader.read_line(&mut header).is_err() {
            return;
        }
        let trimmed = header.trim();
        if trimmed.is_empty() {
            break; // End of headers
        }
        if let Some(val) = trimmed.strip_prefix("Content-Length:") {
            content_length = val.trim().parse().unwrap_or(0);
        }
        if let Some(val) = trimmed.strip_prefix("content-length:") {
            content_length = val.trim().parse().unwrap_or(0);
        }
    }

    // Read body
    let body = if content_length > 0 {
        let mut buf = vec![0u8; content_length];
        let _ = reader.read_exact(&mut buf);
        String::from_utf8_lossy(&buf).to_string()
    } else {
        String::new()
    };

    // CORS headers (Word Add-in may send from https://localhost:*)
    let cors = "Access-Control-Allow-Origin: *\r\nAccess-Control-Allow-Methods: GET, POST, OPTIONS\r\nAccess-Control-Allow-Headers: Content-Type\r\n";

    // Handle OPTIONS preflight
    if method == "OPTIONS" {
        let response = format!("HTTP/1.1 204 No Content\r\n{}\r\n", cors);
        let _ = stream.write_all(response.as_bytes());
        return;
    }

    match (method, path) {
        ("POST", "/context") => {
            // Parse JSON body → CursorContext
            if let Some(ctx) = parse_context_json(&body) {
                if let Ok(mut lock) = cached_context.lock() {
                    *lock = Some((ctx, Instant::now()));
                }
            }
            let response = format!(
                "HTTP/1.1 200 OK\r\n{}Content-Type: application/json\r\nContent-Length: 14\r\n\r\n{{\"status\":\"ok\"}}",
                cors
            );
            let _ = stream.write_all(response.as_bytes());
        }

        ("GET", "/reply") => {
            let json = if let Ok(mut q) = reply_queue.lock() {
                if q.is_empty() {
                    "{}".to_string()
                } else {
                    q.remove(0)
                }
            } else {
                "{}".to_string()
            };
            let response = format!(
                "HTTP/1.1 200 OK\r\n{}Content-Type: application/json\r\nContent-Length: {}\r\n\r\n{}",
                cors,
                json.len(),
                json
            );
            let _ = stream.write_all(response.as_bytes());
        }

        ("GET", "/ping") => {
            let response = format!(
                "HTTP/1.1 200 OK\r\n{}Content-Type: text/plain\r\nContent-Length: 2\r\n\r\nok",
                cors
            );
            let _ = stream.write_all(response.as_bytes());
        }

        _ => {
            let response = format!("HTTP/1.1 404 Not Found\r\n{}\r\n", cors);
            let _ = stream.write_all(response.as_bytes());
        }
    }
}

/// Parse the JSON context from the Word Add-in POST body.
/// The add-in sends: { type, sentence, word, cursorStart, sentenceStart }
fn parse_context_json(body: &str) -> Option<CursorContext> {
    let sentence = extract_json_string(body, "sentence").unwrap_or_default();
    let word = extract_json_string(body, "word").unwrap_or_default();
    let cursor_start = extract_json_number(body, "cursorStart").unwrap_or(0);

    // Build masked_sentence: replace the word with <mask> in the sentence.
    // The BERT completion pipeline requires this.
    let masked_sentence = if word.is_empty() {
        // Cursor at word boundary (after space) — place mask at cursor position in sentence
        if sentence.is_empty() {
            None
        } else {
            Some(format!("{} <mask>", sentence.trim_end()))
        }
    } else if let Some(pos) = sentence.find(&word) {
        // Replace word with <mask>
        let before = &sentence[..pos];
        let after = &sentence[pos + word.len()..];
        Some(format!("{}<mask>{}", before, after))
    } else {
        // Word not found in sentence — just append mask
        Some(format!("{} <mask>", sentence.trim_end()))
    };

    Some(CursorContext {
        word,
        sentence,
        masked_sentence,
        caret_pos: None,
        cursor_doc_offset: Some(cursor_start),
    })
}

/// Extract a string value from JSON by key. Simple parser, no serde.
fn extract_json_string(json: &str, key: &str) -> Option<String> {
    let pattern = format!("\"{}\"", key);
    let key_pos = json.find(&pattern)?;
    let after_key = &json[key_pos + pattern.len()..];
    // Skip whitespace and colon
    let after_colon = after_key.trim_start().strip_prefix(':')?;
    let after_colon = after_colon.trim_start();
    if !after_colon.starts_with('"') {
        return None;
    }
    // Find closing quote (handle escapes)
    let content = &after_colon[1..];
    let mut result = String::new();
    let mut chars = content.chars();
    loop {
        match chars.next() {
            None => return None,
            Some('"') => break,
            Some('\\') => match chars.next() {
                Some('n') => result.push('\n'),
                Some('t') => result.push('\t'),
                Some('r') => result.push('\r'),
                Some('"') => result.push('"'),
                Some('\\') => result.push('\\'),
                Some(c) => {
                    result.push('\\');
                    result.push(c);
                }
                None => return None,
            },
            Some(c) => result.push(c),
        }
    }
    Some(result)
}

/// Extract a numeric value from JSON by key.
fn extract_json_number(json: &str, key: &str) -> Option<usize> {
    let pattern = format!("\"{}\"", key);
    let key_pos = json.find(&pattern)?;
    let after_key = &json[key_pos + pattern.len()..];
    let after_colon = after_key.trim_start().strip_prefix(':')?;
    let after_colon = after_colon.trim_start();
    let num_str: String = after_colon
        .chars()
        .take_while(|c| c.is_ascii_digit())
        .collect();
    num_str.parse().ok()
}

fn escape_json(s: &str) -> String {
    s.replace('\\', "\\\\")
        .replace('"', "\\\"")
        .replace('\n', "\\n")
        .replace('\r', "\\r")
        .replace('\t', "\\t")
}
