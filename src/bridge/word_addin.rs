//! Word Add-in bridge via localhost HTTPS.
//!
//! Serves the Word Add-in static files (taskpane.html, taskpane.js) AND
//! handles API requests — all over HTTPS on a single port.
//! No Python proxy needed.
//!
//! Architecture:
//!   Word Add-in → POST /context → HTTPS thread → Arc<Mutex<CursorContext>>
//!   Rust main   → read_context() → reads cache (instant, never blocks)
//!   Rust main   → replace_word() → pushes to reply queue
//!   Word Add-in → GET /reply → pops from reply queue → applies in Word

use super::{CursorContext, TextBridge};
use std::io::{BufRead, BufReader, Read, Write};
use std::net::TcpListener;
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};
use std::path::PathBuf;

const PORT: u16 = 3000;

/// A changed paragraph from the Word Add-in.
/// Rust side splits into sentences and handles hashing.
#[derive(Debug, Clone)]
pub struct ChangedParagraph {
    pub paragraph_id: String,
    pub text: String,
    /// Absolute cursor position (from sel.start) — used to derive word/sentence for suggestions
    pub cursor_start: Option<usize>,
}

pub struct WordAddinBridge {
    cached_context: Arc<Mutex<Option<(CursorContext, Instant)>>>,
    reply_queue: Arc<Mutex<Vec<String>>>,
    /// Reset flag — set when add-in sends /reset (new document or reload)
    reset_requested: Arc<std::sync::atomic::AtomicBool>,
    /// Current document name — used to detect document switches
    current_doc_name: Arc<Mutex<String>>,
    /// Changed sentences received from add-in (paragraph events).
    /// Main thread picks these up for grammar checking.
    changed_paragraphs: Arc<Mutex<Vec<ChangedParagraph>>>,
    /// Deleted paragraph IDs received from add-in.
    /// Main thread drains these to remove errors for deleted paragraphs.
    deleted_paragraphs: Arc<Mutex<Vec<String>>>,
    /// JSON snapshot of current errors — updated by main thread, read by /errors endpoint
    errors_json: Arc<Mutex<String>>,
}

impl WordAddinBridge {
    /// Start the HTTP server and return the bridge.
    /// Always succeeds — the add-in connects later.
    pub fn new() -> Self {
        let cached_context: Arc<Mutex<Option<(CursorContext, Instant)>>> =
            Arc::new(Mutex::new(None));
        // Queue a rescan command so the add-in re-sends all paragraphs on first connect
        let reply_queue: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(
            vec![r#"{"action":"rescan"}"#.to_string()]
        ));
        let changed_paragraphs: Arc<Mutex<Vec<ChangedParagraph>>> = Arc::new(Mutex::new(Vec::new()));
        let deleted_paragraphs: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
        let reset_requested = Arc::new(std::sync::atomic::AtomicBool::new(false));
        let current_doc_name: Arc<Mutex<String>> = Arc::new(Mutex::new(String::new()));
        let errors_json: Arc<Mutex<String>> = Arc::new(Mutex::new("[]".to_string()));

        let ctx_clone = Arc::clone(&cached_context);
        let reply_clone = Arc::clone(&reply_queue);
        let changed_clone = Arc::clone(&changed_paragraphs);
        let deleted_clone = Arc::clone(&deleted_paragraphs);
        let reset_clone = Arc::clone(&reset_requested);
        let errors_clone = Arc::clone(&errors_json);
        let doc_name_clone = Arc::clone(&current_doc_name);

        std::thread::Builder::new()
            .name("word-addin-https".into())
            .spawn(move || {
                // Load TLS certs from word-addin/ directory
                let addin_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("word-addin");
                let cert_path = addin_dir.join("fullchain.pem");
                let key_path = addin_dir.join("key.pem");

                // Build native-tls acceptor (uses macOS SecureTransport)
                let tls_acceptor = if PORT == 3000 {
                    load_native_tls(&cert_path, &key_path)
                } else {
                    None
                };

                if tls_acceptor.is_some() {
                    eprintln!("Word Add-in HTTPS bridge (native-tls) listening on port {}", PORT);
                } else {
                    eprintln!("Word Add-in HTTP bridge (no TLS certs) on port {}", PORT);
                }

                let listener = match TcpListener::bind(format!("127.0.0.1:{}", PORT)) {
                    Ok(l) => l,
                    Err(e) => {
                        eprintln!("Word Add-in: failed to bind port {}: {}", PORT, e);
                        return;
                    }
                };

                // Cache static files
                let html = std::fs::read_to_string(addin_dir.join("taskpane.html")).unwrap_or_default();
                let js = std::fs::read_to_string(addin_dir.join("taskpane.js")).unwrap_or_default();

                let tls_acceptor = tls_acceptor.map(Arc::new);

                for stream in listener.incoming() {
                    if let Ok(tcp_stream) = stream {
                        let peer = tcp_stream.peer_addr().map(|a| a.to_string()).unwrap_or_default();
                        log_to_file(&format!("TCP connection from {}", peer));
                        let ctx = Arc::clone(&ctx_clone);
                        let reply = Arc::clone(&reply_clone);
                        let changed = Arc::clone(&changed_clone);
                        let deleted = Arc::clone(&deleted_clone);
                        let reset = Arc::clone(&reset_clone);
                        let doc_name = Arc::clone(&doc_name_clone);
                        let errors = Arc::clone(&errors_clone);
                        let tls = tls_acceptor.clone();
                        let html = html.clone();
                        let js = js.clone();
                        std::thread::spawn(move || {
                            if let Some(ref acceptor) = tls {
                                match acceptor.accept(tcp_stream) {
                                    Ok(mut tls_stream) => {
                                        log_to_file("TLS handshake OK");
                                        handle_request_rw(&mut tls_stream, &ctx, &reply, &changed, &deleted, &reset, &doc_name, &errors, &html, &js);
                                    }
                                    Err(e) => {
                                        log_to_file(&format!("TLS accept FAILED: {}", e));
                                    }
                                }
                            } else {
                                let mut stream = tcp_stream;
                                handle_request_rw(&mut stream, &ctx, &reply, &changed, &deleted, &reset, &doc_name, &errors, &html, &js);
                            }
                        });
                    }
                }
            })
            .expect("Failed to spawn Word Add-in HTTP server");

        WordAddinBridge {
            cached_context,
            reply_queue,
            current_doc_name,
            reset_requested,
            changed_paragraphs,
            deleted_paragraphs,
            errors_json,
        }
    }

    /// Update the errors JSON snapshot (called by main thread)
    pub fn update_errors_json(&self, json: &str) {
        if let Ok(mut lock) = self.errors_json.lock() {
            *lock = json.to_string();
        }
    }

    /// Drain changed sentences received from add-in paragraph events.
    /// Main thread calls this to get sentences for grammar checking.
    pub fn drain_changed_paragraphs(&self) -> Vec<ChangedParagraph> {
        if let Ok(mut lock) = self.changed_paragraphs.lock() {
            std::mem::take(&mut *lock)
        } else {
            Vec::new()
        }
    }

    pub fn take_reset(&self) -> bool {
        self.reset_requested.swap(false, std::sync::atomic::Ordering::Relaxed)
    }

    pub fn drain_deleted_paragraphs(&self) -> Vec<String> {
        if let Ok(mut lock) = self.deleted_paragraphs.lock() {
            std::mem::take(&mut *lock)
        } else {
            Vec::new()
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

    fn should_skip_word_spelling(&self, _cursor_off: usize, _word_start: usize, _word_end: usize, _doc_char_len: usize, word_at_cursor: &str) -> bool {
        // The word_at_cursor is checked against ALL words in the sentence.
        // We only want to skip the word the user is currently typing.
        // Since we can't reliably compare offsets (doc vs sentence-relative),
        // we skip when word_at_cursor is mid-word AND is a prefix of or equals the word being checked.
        // This is handled by the caller — we just report if cursor is mid-word.
        false // Let the caller handle skip logic via word comparison
    }

    fn should_skip_sentence_grammar(&self, _cursor_off: usize, _sent_start: usize, _sent_end: usize, ends_with_punct: bool, _doc_char_len: usize, word_at_cursor: &str) -> bool {
        // Skip grammar if user is mid-word and sentence doesn't end with punctuation
        let mid_word = !word_at_cursor.is_empty()
            && word_at_cursor.chars().last().map(|c| c.is_alphanumeric()).unwrap_or(false);
        mid_word && !ends_with_punct
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

    fn select_word_in_paragraph(&self, word: &str, paragraph_id: &str) -> bool {
        let json = format!(
            r#"{{"action":"selectWord","word":"{}","paragraphId":"{}"}}"#,
            escape_json(word), escape_json(paragraph_id)
        );
        self.push_reply(json);
        true
    }

    fn underline_word(&self, word: &str, paragraph_id: &str, color: &str) -> bool {
        let json = format!(
            r#"{{"action":"underline","word":"{}","paragraphId":"{}","color":"{}"}}"#,
            escape_json(word), escape_json(paragraph_id), escape_json(color)
        );
        self.push_reply(json);
        true
    }

    fn clear_underline_word(&self, word: &str, paragraph_id: &str) -> bool {
        let json = format!(
            r#"{{"action":"clearUnderline","word":"{}","paragraphId":"{}"}}"#,
            escape_json(word), escape_json(paragraph_id)
        );
        self.push_reply(json);
        true
    }

    fn clear_paragraph_underlines(&self, paragraph_id: &str) -> bool {
        let json = format!(
            r#"{{"action":"clearParagraphUnderlines","paragraphId":"{}"}}"#,
            escape_json(paragraph_id)
        );
        self.push_reply(json);
        true
    }

    fn drain_changed_paragraphs(&self) -> Vec<ChangedParagraph> {
        self.drain_changed_paragraphs()
    }

    fn drain_deleted_paragraphs(&self) -> Vec<String> {
        self.drain_deleted_paragraphs()
    }

    fn take_reset(&self) -> bool {
        self.take_reset()
    }

    fn update_errors_json(&self, json: &str) {
        if let Ok(mut lock) = self.errors_json.lock() {
            *lock = json.to_string();
        }
    }

    fn push_reply(&self, json: &str) {
        if let Ok(mut q) = self.reply_queue.lock() {
            q.push(json.to_string());
        }
    }

    fn push_reply_urgent(&self, json: &str) {
        if let Ok(mut q) = self.reply_queue.lock() {
            q.insert(0, json.to_string());
        }
    }
}

// ── HTTP handling ──

fn handle_request_rw<S: Read + Write>(
    stream: &mut S,
    cached_context: &Arc<Mutex<Option<(CursorContext, Instant)>>>,
    reply_queue: &Arc<Mutex<Vec<String>>>,
    changed_paragraphs: &Arc<Mutex<Vec<ChangedParagraph>>>,
    deleted_paragraphs: &Arc<Mutex<Vec<String>>>,
    reset_requested: &Arc<std::sync::atomic::AtomicBool>,
    current_doc_name: &Arc<Mutex<String>>,
    errors_json: &Arc<Mutex<String>>,
    static_html: &str,
    static_js: &str,
) {
    // Read headers byte by byte until we find \r\n\r\n
    let mut header_buf = Vec::with_capacity(4096);
    let mut single = [0u8; 1];
    loop {
        match stream.read(&mut single) {
            Ok(1) => {
                header_buf.push(single[0]);
                if header_buf.len() >= 4 && &header_buf[header_buf.len()-4..] == b"\r\n\r\n" {
                    break;
                }
                if header_buf.len() > 8192 { return; } // too long
            }
            _ => return,
        }
    }
    let header_str = String::from_utf8_lossy(&header_buf).to_string();

    // Parse request line
    let first_line = header_str.lines().next().unwrap_or("");
    log_to_file(&format!("Request: {}", first_line));
    let parts: Vec<&str> = first_line.split_whitespace().collect();
    if parts.len() < 2 {
        return;
    }
    let method = parts[0];
    let path = parts[1];

    // Parse Content-Length from headers
    let content_length: usize = header_str.lines()
        .find(|l| l.to_lowercase().starts_with("content-length:"))
        .and_then(|l| l.split(':').nth(1))
        .and_then(|v| v.trim().parse().ok())
        .unwrap_or(0);

    // Read body based on Content-Length
    let body = if content_length > 0 {
        let mut body_buf = vec![0u8; content_length];
        let mut read = 0;
        while read < content_length {
            match stream.read(&mut body_buf[read..]) {
                Ok(0) => break,
                Ok(n) => read += n,
                Err(_) => break,
            }
        }
        String::from_utf8_lossy(&body_buf[..read]).to_string()
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

    // Strip query string from path
    let path = path.split('?').next().unwrap_or(path);

    match (method, path) {
        ("POST", "/log") => {
            let msg = extract_json_string(&body, "msg").unwrap_or_default();
            log_to_file(&format!("JS: {}", msg));
            let response = format!("HTTP/1.1 200 OK\r\n{}Content-Length: 2\r\n\r\n{{}}", cors);
            let _ = stream.write_all(response.as_bytes());
            return;
        }
        ("POST", "/context") => {
            // Check if document changed — compare documentName
            let doc_name = extract_json_string(&body, "documentName").unwrap_or_default();
            if doc_name.is_empty() {
                crate::log!("CONTEXT: no documentName in body (len={})", body.len());
            }
            let needs_rescan = if !doc_name.is_empty() {
                if let Ok(mut current) = current_doc_name.lock() {
                    if !current.is_empty() && *current != doc_name {
                        // Different non-empty name — real document switch
                        crate::log!("DOC SWITCH: '{}' → '{}' — requesting rescan", *current, doc_name);
                        *current = doc_name;
                        true
                    } else {
                        if current.is_empty() {
                            crate::log!("DOC NAME SET: '{}'", doc_name);
                        }
                        *current = doc_name;
                        false
                    }
                } else {
                    false
                }
            } else {
                false
            };

            // Parse JSON body → CursorContext
            if let Some(ctx) = parse_context_json(&body) {
                if let Ok(mut lock) = cached_context.lock() {
                    *lock = Some((ctx, Instant::now()));
                }
            }

            let status = if needs_rescan { "rescan" } else { "ok" };
            let resp_body = format!("{{\"status\":\"{}\"}}", status);
            let response = format!(
                "HTTP/1.1 200 OK\r\n{}Content-Type: application/json\r\nContent-Length: {}\r\n\r\n{}",
                cors, resp_body.len(), resp_body
            );
            let _ = stream.write_all(response.as_bytes());
        }

        ("GET", p) if p.starts_with("/reply") => {
            // Only return replies to the active document
            let req_doc = p.strip_prefix("/reply?doc=")
                .map(|d| d.replace("%3A", ":").replace("%2F", "/").replace("%20", " ").replace("%25", "%"))
                .unwrap_or_default();
            let is_active = if req_doc.is_empty() {
                true
            } else if let Ok(current) = current_doc_name.lock() {
                current.is_empty() || *current == req_doc
            } else {
                true
            };
            let json = if is_active {
                if let Ok(mut q) = reply_queue.lock() {
                    if q.is_empty() {
                        "{}".to_string()
                    } else {
                        let j = q.remove(0);
                        log_to_file(&format!("REPLY sending: {}", j));
                        j
                    }
                } else {
                    "{}".to_string()
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

        ("POST", "/changed") => {
            // Only accept from the active document
            let doc_name = extract_json_string(&body, "documentName").unwrap_or_default();
            let is_active = if doc_name.is_empty() {
                true // no doc name = legacy, accept
            } else if let Ok(current) = current_doc_name.lock() {
                current.is_empty() || *current == doc_name
            } else {
                true
            };
            if is_active {
                if let Some(sentences) = parse_changed_json(&body) {
                    if let Ok(mut lock) = changed_paragraphs.lock() {
                        lock.extend(sentences);
                    }
                }
            }
            let response = format!(
                "HTTP/1.1 200 OK\r\n{}Content-Type: application/json\r\nContent-Length: 14\r\n\r\n{{\"status\":\"ok\"}}",
                cors
            );
            let _ = stream.write_all(response.as_bytes());
        }

        ("POST", "/deleted") => {
            // Parse deleted paragraph IDs: {"paragraphIds":["id1","id2"]}
            if let Some(ids) = parse_deleted_json(&body) {
                eprintln!("HTTP /deleted: {} paragraph IDs: {:?}", ids.len(), &ids[..ids.len().min(5)]);
                if let Ok(mut lock) = deleted_paragraphs.lock() {
                    lock.extend(ids);
                }
            }
            let response = format!(
                "HTTP/1.1 200 OK\r\n{}Content-Type: application/json\r\nContent-Length: 14\r\n\r\n{{\"status\":\"ok\"}}",
                cors
            );
            let _ = stream.write_all(response.as_bytes());
        }

        ("POST", "/reset") => {
            reset_requested.store(true, std::sync::atomic::Ordering::Relaxed);
            // Store document name from reset so we know which doc is active
            let doc_name = extract_json_string(&body, "documentName").unwrap_or_default();
            if !doc_name.is_empty() {
                if let Ok(mut current) = current_doc_name.lock() {
                    crate::log!("RESET from doc: '{}'", doc_name);
                    *current = doc_name;
                }
            }
            // Clear reply queue so stale underlines from old document don't leak
            if let Ok(mut q) = reply_queue.lock() {
                q.clear();
            }
            eprintln!("HTTP /reset: clearing all state");
            let response = format!(
                "HTTP/1.1 200 OK\r\n{}Content-Type: application/json\r\nContent-Length: 14\r\n\r\n{{\"status\":\"ok\"}}",
                cors
            );
            let _ = stream.write_all(response.as_bytes());
        }

        // Test endpoint: push a command to the add-in reply queue
        ("POST", "/push-reply") => {
            if let Ok(mut q) = reply_queue.lock() {
                q.push(body.clone());
            }
            let response = format!(
                "HTTP/1.1 200 OK\r\n{}Content-Type: application/json\r\nContent-Length: 14\r\n\r\n{{\"status\":\"ok\"}}",
                cors
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

        ("GET", "/errors") => {
            let json = errors_json.lock().map(|l| l.clone()).unwrap_or_else(|_| "[]".to_string());
            let response = format!(
                "HTTP/1.1 200 OK\r\n{}Content-Type: application/json; charset=utf-8\r\nContent-Length: {}\r\n\r\n{}",
                cors, json.len(), json
            );
            let _ = stream.write_all(response.as_bytes());
        }

        ("GET", "/taskpane.html") | ("GET", "/") => {
            let response = format!(
                "HTTP/1.1 200 OK\r\n{}Content-Type: text/html; charset=utf-8\r\nCache-Control: no-cache, no-store, must-revalidate\r\nContent-Length: {}\r\n\r\n{}",
                cors, static_html.len(), static_html
            );
            let _ = stream.write_all(response.as_bytes());
        }

        ("GET", path) if path.starts_with("/taskpane.js") => {
            let response = format!(
                "HTTP/1.1 200 OK\r\n{}Content-Type: application/javascript; charset=utf-8\r\nCache-Control: no-cache, no-store, must-revalidate\r\nContent-Length: {}\r\n\r\n{}",
                cors, static_js.len(), static_js
            );
            let _ = stream.write_all(response.as_bytes());
        }

        _ => {
            let response = format!("HTTP/1.1 404 Not Found\r\n{}\r\n", cors);
            let _ = stream.write_all(response.as_bytes());
        }
    }
}

/// Load TLS identity from PEM cert+key files, build a native-tls acceptor.
/// native-tls uses macOS SecureTransport — same TLS backend as Python's ssl module.
fn load_native_tls(cert_path: &std::path::Path, key_path: &std::path::Path) -> Option<native_tls::TlsAcceptor> {
    let cert_pem = std::fs::read(cert_path).ok()?;
    let key_pem = std::fs::read(key_path).ok()?;

    // native-tls on macOS needs PKCS#12 format — convert PEM cert+key to PKCS#12
    let identity = native_tls::Identity::from_pkcs8(&cert_pem, &key_pem)
        .map_err(|e| eprintln!("native-tls Identity error: {}", e))
        .ok()?;

    let acceptor = native_tls::TlsAcceptor::new(identity)
        .map_err(|e| eprintln!("native-tls Acceptor error: {}", e))
        .ok()?;

    Some(acceptor)
}

/// Clean Word special characters from text.
/// Word uses \u{000b} (vertical tab) for soft line breaks (Shift+Enter).
fn clean_word_text(s: &str) -> String {
    s.replace('\u{000b}', " ")
     .replace('\u{0007}', "")   // bell (table cells)
     .replace('\u{000c}', " ")  // form feed (page break)
     .replace('\u{000d}', " ")  // carriage return
}

/// Parse the JSON context from the Word Add-in POST body.
/// The add-in sends: { type, sentence, word, cursorStart, sentenceStart }
fn parse_context_json(body: &str) -> Option<CursorContext> {
    let sentence = clean_word_text(&extract_json_string(body, "sentence").unwrap_or_default());
    let word = clean_word_text(&extract_json_string(body, "word").unwrap_or_default());
    let cursor_start = extract_json_number(body, "cursorStart").unwrap_or(0);
    let sync_ms = extract_json_number(body, "syncMs").unwrap_or(0);
    let paragraph_id = extract_json_string(body, "paragraphId").unwrap_or_default();
    if sync_ms > 0 {
        eprintln!("JS sync: {}ms word='{}' pos={}", sync_ms, word, cursor_start);
    }

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
        paragraph_id,
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
                Some('/') => result.push('/'),
                Some('u') => {
                    // Decode \uXXXX Unicode escape
                    let mut hex = String::with_capacity(4);
                    for _ in 0..4 {
                        match chars.next() {
                            Some(h) => hex.push(h),
                            None => break,
                        }
                    }
                    if let Ok(code) = u32::from_str_radix(&hex, 16) {
                        if let Some(c) = char::from_u32(code) {
                            result.push(c);
                        }
                    }
                }
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

/// Parse changed paragraphs from the add-in POST.
/// JSON: { "type": "changed", "paragraphs": [{"paragraphId": "...", "text": "..."}] }
fn parse_changed_json(body: &str) -> Option<Vec<ChangedParagraph>> {
    let arr_start = body.find("\"paragraphs\"")?;
    let arr_body = &body[arr_start..];
    let bracket_start = arr_body.find('[')?;
    let bracket_end = arr_body.rfind(']')?;
    let arr_content = &arr_body[bracket_start + 1..bracket_end];

    let mut results = Vec::new();
    let mut pos = 0;
    while pos < arr_content.len() {
        let obj_start = match arr_content[pos..].find('{') {
            Some(s) => pos + s,
            None => break,
        };
        let obj_end = match arr_content[obj_start..].find('}') {
            Some(e) => obj_start + e + 1,
            None => break,
        };
        let obj = &arr_content[obj_start..obj_end];

        let paragraph_id = extract_json_string(obj, "paragraphId").unwrap_or_default();
        let text = clean_word_text(&extract_json_string(obj, "text").unwrap_or_default());
        let cursor_start = extract_json_number(obj, "cursorStart");

        if !text.is_empty() {
            results.push(ChangedParagraph { paragraph_id, text, cursor_start });
        }

        pos = obj_end;
    }

    if results.is_empty() { None } else { Some(results) }
}

/// Parse deleted paragraph IDs from the add-in's POST.
/// JSON: { "paragraphIds": ["id1", "id2"] }
fn parse_deleted_json(body: &str) -> Option<Vec<String>> {
    let arr_start = body.find("\"paragraphIds\"")?;
    let arr_body = &body[arr_start..];
    let bracket_start = arr_body.find('[')?;
    let bracket_end = arr_body.rfind(']')?;
    let arr_content = &arr_body[bracket_start + 1..bracket_end];

    let mut results = Vec::new();
    let mut pos = 0;
    while pos < arr_content.len() {
        let quote_start = match arr_content[pos..].find('"') {
            Some(s) => pos + s,
            None => break,
        };
        let quote_end = match arr_content[quote_start + 1..].find('"') {
            Some(e) => quote_start + 1 + e,
            None => break,
        };
        let id = &arr_content[quote_start + 1..quote_end];
        if !id.is_empty() {
            results.push(id.to_string());
        }
        pos = quote_end + 1;
    }

    if results.is_empty() { None } else { Some(results) }
}

fn log_to_file(msg: &str) {
    use std::io::Write;
    let path = "/tmp/word_addin_bridge.log";
    if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true).open(path) {
        let _ = writeln!(f, "{} {}", chrono_now(), msg);
    }
}

fn chrono_now() -> String {
    let d = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default();
    format!("{}.{:03}", d.as_secs(), d.subsec_millis())
}
