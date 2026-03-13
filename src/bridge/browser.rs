/// Browser bridge — reads textarea data from Chrome/Edge extension via native messaging.
///
/// The extension sends text + cursor position to a native messaging host,
/// which writes it to a temp JSON file. This bridge reads that file.

use super::{CursorContext, RawCursorText, TextBridge, build_context};
use std::path::PathBuf;
use std::time::Instant;

fn log_browser(msg: &str) {
    use std::io::Write;
    if let Ok(mut f) = crate::LOG_FILE.lock() {
        let _ = writeln!(f, "{}", msg);
        let _ = f.flush();
    }
}

fn data_path() -> PathBuf {
    std::env::temp_dir().join("norsktale-browser.json")
}

fn reply_path() -> PathBuf {
    std::env::temp_dir().join("norsktale-browser-reply.json")
}

pub struct BrowserBridge {
    last_modified: std::cell::Cell<u64>,
    last_text: std::cell::RefCell<String>,
    last_cursor: std::cell::Cell<usize>,
    last_caret: std::cell::Cell<Option<(i32, i32)>>,
    last_read: std::cell::Cell<Option<Instant>>,
}

impl BrowserBridge {
    pub fn new() -> Self {
        // Delete stale data file from previous session to avoid processing old text
        let _ = std::fs::remove_file(data_path());
        BrowserBridge {
            last_modified: std::cell::Cell::new(0),
            last_text: std::cell::RefCell::new(String::new()),
            last_cursor: std::cell::Cell::new(0),
            last_caret: std::cell::Cell::new(None),
            last_read: std::cell::Cell::new(None),
        }
    }

    fn read_data_file(&self) -> Option<(String, usize, usize, Option<(i32, i32)>)> {
        // Rate-limit file reads to every 100ms
        if let Some(last) = self.last_read.get() {
            if last.elapsed().as_millis() < 100 {
                let text = self.last_text.borrow().clone();
                if !text.is_empty() {
                    let cursor = self.last_cursor.get();
                    return Some((text, cursor, cursor, self.last_caret.get()));
                }
                return None;
            }
        }
        self.last_read.set(Some(Instant::now()));

        let path = data_path();
        let metadata = std::fs::metadata(&path).ok()?;
        let modified = metadata.modified().ok()?
            .duration_since(std::time::UNIX_EPOCH).ok()?
            .as_millis() as u64;

        // Only re-read if file changed
        if modified == self.last_modified.get() {
            let text = self.last_text.borrow().clone();
            if !text.is_empty() {
                let cursor = self.last_cursor.get();
                return Some((text, cursor, cursor, self.last_caret.get()));
            }
            return None;
        }
        self.last_modified.set(modified);

        let content = std::fs::read_to_string(&path).ok()?;

        // Parse JSON: { "text": "...", "cursorStart": N, "cursorEnd": N, ... }
        // Minimal JSON parsing to avoid adding serde dependency
        let text = extract_json_string(&content, "text")?;
        let cursor_start = extract_json_number(&content, "cursorStart").unwrap_or(text.len());
        let cursor_end = extract_json_number(&content, "cursorEnd").unwrap_or(cursor_start);
        let caret = match (extract_json_number(&content, "caretX"), extract_json_number(&content, "caretY")) {
            (Some(x), Some(y)) => Some((x as i32, y as i32)),
            _ => None,
        };

        *self.last_text.borrow_mut() = text.clone();
        self.last_cursor.set(cursor_start);
        self.last_caret.set(caret);

        Some((text, cursor_start, cursor_end, caret))
    }

    /// Update the cached text to reflect a replacement we just sent to the extension.
    /// This prevents re-detecting errors in stale text before the extension sends the update back.
    fn update_cached_text(&self, char_start: usize, char_end: usize, replacement: &str) {
        let mut text = self.last_text.borrow_mut();
        let byte_start = text.char_indices()
            .nth(char_start)
            .map(|(i, _)| i)
            .unwrap_or(text.len());
        let byte_end = text.char_indices()
            .nth(char_end)
            .map(|(i, _)| i)
            .unwrap_or(text.len());
        text.replace_range(byte_start..byte_end, replacement);
        // Update cursor to end of replacement
        self.last_cursor.set(char_start + replacement.chars().count());
    }
}

impl TextBridge for BrowserBridge {
    fn name(&self) -> &str {
        "Browser"
    }

    fn replace_word(&self, new_text: &str) -> bool {
        // Find the current word boundaries from last known cursor position
        let text = self.last_text.borrow().clone();
        let cursor = self.last_cursor.get();
        if text.is_empty() { return false; }

        let cursor_byte = char_to_byte_offset(&text, cursor);
        let before = &text[..cursor_byte];
        let after = &text[cursor_byte..];

        // Find word boundaries (same logic as extract_word_before/after_cursor)
        let word_before_len: usize = before.chars().rev()
            .take_while(|c| c.is_alphanumeric() || *c == '-' || *c == '\'')
            .count();
        let word_after_len: usize = after.chars()
            .take_while(|c| c.is_alphanumeric() || *c == '-' || *c == '\'')
            .count();

        let start = cursor - word_before_len;
        let end = cursor + word_after_len;

        // Escape the replacement text for JSON
        let escaped = new_text.replace('\\', "\\\\").replace('"', "\\\"");
        let json = format!(
            r#"{{"action":"replace","start":{},"end":{},"text":"{}"}}"#,
            start, end, escaped
        );
        if std::fs::write(reply_path(), json.as_bytes()).is_ok() {
            // Update cached text so next read_context() sees the post-replacement text
            self.update_cached_text(start, end, new_text);
            true
        } else {
            false
        }
    }

    fn find_and_replace(&self, find: &str, replace: &str) -> bool {
        let text = self.last_text.borrow().clone();
        if text.is_empty() { return false; }
        let text_lower = text.to_lowercase();
        let find_lower = find.to_lowercase();
        if let Some(byte_pos) = text_lower.find(&find_lower) {
            let start = text[..byte_pos].chars().count();
            let end = start + find.chars().count();
            let escaped_text = replace.replace('\\', "\\\\").replace('"', "\\\"");
            let escaped_find = find.replace('\\', "\\\\").replace('"', "\\\"");
            let json = format!(
                r#"{{"action":"replace","start":{},"end":{},"text":"{}","expected":"{}"}}"#,
                start, end, escaped_text, escaped_find
            );
            if std::fs::write(reply_path(), json.as_bytes()).is_ok() {
                self.update_cached_text(start, end, replace);
                return true;
            }
        }
        false
    }

    fn find_and_replace_in_context(&self, find: &str, replace: &str, context: &str) -> bool {
        let text = self.last_text.borrow().clone();
        if text.is_empty() { return false; }
        let text_lower = text.to_lowercase();
        let ctx_lower = context.to_lowercase();
        let ctx_byte_start = text_lower.find(&ctx_lower).unwrap_or(0);
        let find_lower = find.to_lowercase();
        let search_region = &text_lower[ctx_byte_start..];
        if let Some(rel_byte_pos) = search_region.find(&find_lower) {
            let abs_byte_pos = ctx_byte_start + rel_byte_pos;
            let start = text[..abs_byte_pos].chars().count();
            let end = start + find.chars().count();
            let escaped_text = replace.replace('\\', "\\\\").replace('"', "\\\"");
            let escaped_find = find.replace('\\', "\\\\").replace('"', "\\\"");
            let json = format!(
                r#"{{"action":"replace","start":{},"end":{},"text":"{}","expected":"{}"}}"#,
                start, end, escaped_text, escaped_find
            );
            if std::fs::write(reply_path(), json.as_bytes()).is_ok() {
                self.update_cached_text(start, end, replace);
                return true;
            }
        }
        false
    }

    fn find_and_replace_in_context_at(&self, find: &str, replace: &str, context: &str, char_offset: usize) -> bool {
        let text = self.last_text.borrow().clone();
        if text.is_empty() { return false; }
        log_browser(&format!("REPLACE: find='{}' replace='{}' char_offset={}", find, replace, char_offset));
        log_browser(&format!("  cached text ({} chars): '{}'", text.chars().count(), &text[..text.len().min(200)]));
        // Use the char_offset to find the exact position
        let byte_offset = char_to_byte_offset(&text, char_offset);
        let find_lower = find.to_lowercase();
        // Search near the offset (within ±200 bytes)
        let search_start = byte_offset.saturating_sub(200).min(text.len());
        let search_end = (byte_offset + 200).min(text.len());
        // Ensure we're at char boundaries
        let search_start = (0..=search_start).rev().find(|&i| text.is_char_boundary(i)).unwrap_or(0);
        let search_end = (search_end..=text.len()).find(|&i| text.is_char_boundary(i)).unwrap_or(text.len());
        let region = &text[search_start..search_end];
        let region_lower = region.to_lowercase();
        if let Some(rel_byte_pos) = region_lower.find(&find_lower) {
            let abs_byte_pos = search_start + rel_byte_pos;
            let start = text[..abs_byte_pos].chars().count();
            let end = start + find.chars().count();
            log_browser(&format!("  FOUND at char {}..{}, sending replace JSON", start, end));
            log_browser(&format!("  text BEFORE replace: '{}'", &text[..text.len().min(200)]));
            let escaped = replace.replace('\\', "\\\\").replace('"', "\\\"");
            let find_escaped = find.replace('\\', "\\\\").replace('"', "\\\"");
            let json = format!(
                r#"{{"action":"replace","start":{},"end":{},"text":"{}","expected":"{}"}}"#,
                start, end, escaped, find_escaped
            );
            log_browser(&format!("  reply JSON: {}", json));
            if std::fs::write(reply_path(), json.as_bytes()).is_ok() {
                self.update_cached_text(start, end, replace);
                let new_text = self.last_text.borrow().clone();
                log_browser(&format!("  text AFTER replace: '{}'", &new_text[..new_text.len().min(200)]));
                return true;
            }
            return false;
        }
        log_browser("  NOT FOUND near offset, falling back to context search");
        // Fallback to context-based search
        self.find_and_replace_in_context(find, replace, context)
    }

    fn is_available(&self) -> bool {
        data_path().exists()
    }

    fn read_context(&self) -> Option<CursorContext> {
        let (text, cursor_start, _cursor_end, caret) = self.read_data_file()?;
        if text.is_empty() { return None; }

        // Split text at cursor position (byte-safe)
        let cursor_byte = char_to_byte_offset(&text, cursor_start);
        let before = &text[..cursor_byte];
        let after = &text[cursor_byte..];

        let raw = RawCursorText {
            before: before.to_string(),
            after: after.to_string(),
        };

        let mut ctx = build_context(&raw, caret);
        ctx.cursor_doc_offset = Some(cursor_start);
        Some(ctx)
    }

    fn read_full_document(&self) -> Option<String> {
        // Re-read file to get latest text
        let (text, _, _, _) = self.read_data_file()?;
        if text.is_empty() { None } else { Some(text) }
    }
}

/// Convert a character offset to a byte offset in a UTF-8 string
fn char_to_byte_offset(s: &str, char_offset: usize) -> usize {
    s.char_indices()
        .nth(char_offset)
        .map(|(byte_idx, _)| byte_idx)
        .unwrap_or(s.len())
}

/// Extract a string value from JSON by key (simple parser, no serde needed)
fn extract_json_string(json: &str, key: &str) -> Option<String> {
    let pattern = format!("\"{}\"", key);
    let key_pos = json.find(&pattern)?;
    let after_key = &json[key_pos + pattern.len()..];
    // Skip whitespace and colon
    let after_colon = after_key.trim_start().strip_prefix(':')?;
    let after_ws = after_colon.trim_start();
    if !after_ws.starts_with('"') { return None; }
    let content = &after_ws[1..];
    // Find closing quote (handle escapes)
    let mut result = String::new();
    let mut chars = content.chars();
    loop {
        match chars.next() {
            Some('"') => break,
            Some('\\') => {
                match chars.next() {
                    Some('n') => result.push('\n'),
                    Some('t') => result.push('\t'),
                    Some('r') => result.push('\r'),
                    Some('"') => result.push('"'),
                    Some('\\') => result.push('\\'),
                    Some(c) => { result.push('\\'); result.push(c); }
                    None => break,
                }
            }
            Some(c) => result.push(c),
            None => break,
        }
    }
    Some(result)
}

/// Extract a number value from JSON by key
fn extract_json_number(json: &str, key: &str) -> Option<usize> {
    let pattern = format!("\"{}\"", key);
    let key_pos = json.find(&pattern)?;
    let after_key = &json[key_pos + pattern.len()..];
    let after_colon = after_key.trim_start().strip_prefix(':')?;
    let after_ws = after_colon.trim_start();
    let num_str: String = after_ws.chars().take_while(|c| c.is_ascii_digit()).collect();
    num_str.parse().ok()
}
