/// Browser bridge — reads textarea data from Chrome/Edge extension via native messaging.
///
/// The extension sends text + cursor position to a native messaging host,
/// which writes it to a temp JSON file. This bridge reads that file.

use super::{CursorContext, RawCursorText, TextBridge, build_context, extract_previous_word_before_cursor};
use std::path::PathBuf;
use std::time::Instant;

fn log_browser(msg: &str) {
    use std::io::Write;
    if let Ok(mut f) = crate::logging::LOG_FILE.lock() {
        let _ = writeln!(f, "{}", msg);
        let _ = f.flush();
    }
}

fn data_path() -> PathBuf {
    let name = if cfg!(test) {
        format!("spell-browser-test-{}-data.json", std::process::id())
    } else {
        "spell-browser.json".to_string()
    };
    std::env::temp_dir().join(name)
}

fn reply_path_for(bridge_id: usize) -> PathBuf {
    let name = if cfg!(test) {
        format!("spell-browser-test-{}-reply-{}.json", std::process::id(), bridge_id)
    } else if bridge_id > 0 {
        format!("spell-browser-reply-{}.json", bridge_id)
    } else {
        "spell-browser-reply.json".to_string()
    };
    std::env::temp_dir().join(name)
}

pub struct BrowserBridge {
    last_modified: std::cell::Cell<u64>,
    last_text: std::cell::RefCell<String>,
    // The exact browser-editor selection, supplied by the extension. This
    // must not fall back to an unrelated desktop app's accessibility cache.
    last_selected_text: std::cell::RefCell<Option<String>>,
    last_cursor: std::cell::Cell<usize>,
    /// Absolute start of the active browser paragraph, as published by the
    /// Docs injector. This keeps identical paragraphs distinct.
    last_paragraph_start: std::cell::Cell<usize>,
    last_caret: std::cell::Cell<Option<(i32, i32)>>,
    last_source: std::cell::RefCell<String>,
    last_bridge_id: std::cell::Cell<usize>,
    last_frame_id: std::cell::Cell<usize>,
    pending_completed_word: std::cell::RefCell<Option<String>>,
    last_read: std::cell::Cell<Option<Instant>>,
    /// After sending a replace command, freeze reads from the file until fresh data arrives.
    /// The file still contains pre-replace text; re-reading it would undo the cached fix.
    replace_freeze_modified: std::cell::Cell<u64>,
    /// The word being replaced — used to verify fresh data doesn't still contain it.
    replace_old_word: std::cell::RefCell<String>,
    /// When the freeze was activated — timeout after 5 seconds.
    replace_freeze_time: std::cell::Cell<Option<Instant>>,
}

impl BrowserBridge {
    pub fn new() -> Self {
        // Delete stale data file from previous session to avoid processing old text
        let _ = std::fs::remove_file(data_path());
        BrowserBridge {
            last_modified: std::cell::Cell::new(0),
            last_text: std::cell::RefCell::new(String::new()),
            last_selected_text: std::cell::RefCell::new(None),
            last_cursor: std::cell::Cell::new(0),
            last_paragraph_start: std::cell::Cell::new(0),
            last_caret: std::cell::Cell::new(None),
            last_source: std::cell::RefCell::new(String::new()),
            last_bridge_id: std::cell::Cell::new(0),
            last_frame_id: std::cell::Cell::new(0),
            pending_completed_word: std::cell::RefCell::new(None),
            last_read: std::cell::Cell::new(None),
            replace_freeze_modified: std::cell::Cell::new(0),
            replace_old_word: std::cell::RefCell::new(String::new()),
            replace_freeze_time: std::cell::Cell::new(None),
        }
    }

    fn cached_data(&self) -> Option<(String, usize, usize, Option<(i32, i32)>)> {
        if self.last_modified.get() == 0 {
            return None;
        }
        let text = self.last_text.borrow().clone();
        let cursor = self.last_cursor.get();
        Some((text, cursor, cursor, self.last_caret.get()))
    }

    fn source_tab_id(&self) -> usize {
        self.last_source
            .borrow()
            .split_once('|')
            .and_then(|(tab_id, _url)| tab_id.parse().ok())
            .unwrap_or(0)
    }

    fn reply_path(&self) -> PathBuf {
        reply_path_for(self.last_bridge_id.get())
    }

    fn source_frame_id(&self) -> usize {
        self.last_frame_id.get()
    }

    fn read_data_file(&self) -> Option<(String, usize, usize, Option<(i32, i32)>)> {
        // Rate-limit file reads to every 100ms
        if let Some(last) = self.last_read.get() {
            if last.elapsed().as_millis() < 100 {
                return self.cached_data();
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
            return self.cached_data();
        }

        // After a replace, the file still contains pre-replace text until the
        // extension writes fresh data. Skip re-reading stale file — use cached
        // (post-replace) text instead.
        let freeze = self.replace_freeze_modified.get();
        let freeze_timed_out = freeze > 0 && self.replace_freeze_time.get()
            .map(|t| t.elapsed().as_secs() >= 5)
            .unwrap_or(false);
        if freeze > 0 && freeze_timed_out {
            crate::debug_log!("read_data_file: freeze timed out after 5s — accepting file data");
            self.replace_freeze_modified.set(0);
            self.replace_old_word.borrow_mut().clear();
            self.replace_freeze_time.set(None);
            // Fall through to read the file normally
        } else if freeze > 0 && modified <= freeze {
            return self.cached_data();
        } else if freeze > 0 {
            // File is newer than freeze — but verify the old word is actually gone.
            let old_word = self.replace_old_word.borrow().clone();
            if !old_word.is_empty() {
                let content = std::fs::read_to_string(&path).ok();
                if let Some(ref c) = content {
                    if let Some(file_text) = extract_json_string(c, "text") {
                        let file_lower = file_text.to_lowercase();
                        if has_whole_word(&file_lower, &old_word) {
                            crate::debug_log!("read_data_file: 'fresh' data still has '{}' — keeping freeze", old_word);
                            return self.cached_data();
                        }
                    }
                }
            }
            crate::debug_log!("read_data_file: fresh data confirmed (old word gone), clearing freeze");
            self.replace_freeze_modified.set(0);
            self.replace_old_word.borrow_mut().clear();
            self.replace_freeze_time.set(None);
        }

        self.last_modified.set(modified);

        let content = std::fs::read_to_string(&path).ok()?;

        // Parse JSON: { "text": "...", "cursorStart": N, "cursorEnd": N, ... }
        // Minimal JSON parsing to avoid adding serde dependency
        let text = extract_json_string(&content, "text")?;
        let selected_text = extract_json_string(&content, "selectedText")
            .filter(|selection| !selection.trim().is_empty());
        let cursor_start = extract_json_number(&content, "cursorStart").unwrap_or(text.len());
        let cursor_end = extract_json_number(&content, "cursorEnd").unwrap_or(cursor_start);
        let paragraph_start = extract_json_number(&content, "paragraphStart").unwrap_or(0);
        let caret = match (extract_json_number(&content, "caretX"), extract_json_number(&content, "caretY")) {
            (Some(x), Some(y)) => Some((x as i32, y as i32)),
            _ => None,
        };
        let source = format!(
            "{}|{}",
            extract_json_number(&content, "tabId").unwrap_or(0),
            extract_json_string(&content, "url").unwrap_or_default(),
        );
        let bridge_id = extract_json_number(&content, "bridgeId").unwrap_or(0);
        let frame_id = extract_json_number(&content, "frameId").unwrap_or(0);

        let previous_source = self.last_source.borrow().clone();
        if !previous_source.is_empty()
            && previous_source == source
            && self.last_bridge_id.get() == bridge_id
            && self.last_frame_id.get() == frame_id
            && self.last_paragraph_start.get() == paragraph_start
        {
            let previous_text = self.last_text.borrow();
            if let Some(word) = completed_word_from_transition(
                &previous_text,
                self.last_cursor.get(),
                &text,
                cursor_start,
            ) {
                log_browser(&format!("Browser TTS boundary: '{}'", word));
                *self.pending_completed_word.borrow_mut() = Some(word);
            }
        } else {
            self.pending_completed_word.borrow_mut().take();
        }

        *self.last_text.borrow_mut() = text.clone();
        *self.last_selected_text.borrow_mut() = selected_text;
        self.last_cursor.set(cursor_start);
        self.last_paragraph_start.set(paragraph_start);
        self.last_caret.set(caret);
        *self.last_source.borrow_mut() = source;
        self.last_bridge_id.set(bridge_id);
        self.last_frame_id.set(frame_id);

        Some((text, cursor_start, cursor_end, caret))
    }

    /// Freeze file reads — the on-disk file still has pre-replace text.
    /// Only allow reads again when fresh data arrives without the old word.
    fn activate_replace_freeze(&self, old_word: &str) {
        *self.replace_old_word.borrow_mut() = old_word.to_lowercase();
        let modified = std::fs::metadata(data_path())
            .ok()
            .and_then(|m| m.modified().ok())
            .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
            .map(|d| d.as_millis() as u64)
            .unwrap_or(0);
        log_browser(&format!("activate_replace_freeze: frozen at modified={}, word='{}'", modified, old_word));
        self.replace_freeze_modified.set(modified);
        self.replace_freeze_time.set(Some(Instant::now()));
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
        let old_word: String = text.chars().skip(start).take(end - start).collect();
        let replacement = new_text
            .split_once('|')
            .map(|(_prefix, word)| word)
            .unwrap_or(new_text);

        // Escape the replacement text for JSON
        let escaped = replacement.replace('\\', "\\\\").replace('"', "\\\"");
        let escaped_old_word = old_word.replace('\\', "\\\\").replace('"', "\\\"");
        let json = format!(
            r#"{{"action":"replace","tabId":{},"frameId":{},"start":{},"end":{},"text":"{}","expected":"{}","paragraphStart":{}}}"#,
            self.source_tab_id(), self.source_frame_id(), start, end, escaped, escaped_old_word, self.last_paragraph_start.get()
        );
        if std::fs::write(self.reply_path(), json.as_bytes()).is_ok() {
            self.update_cached_text(start, end, replacement);
            self.activate_replace_freeze(&old_word);
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
                r#"{{"action":"replace","tabId":{},"frameId":{},"start":{},"end":{},"text":"{}","expected":"{}","paragraphStart":{}}}"#,
                self.source_tab_id(), self.source_frame_id(), start, end, escaped_text, escaped_find, self.last_paragraph_start.get()
            );
            if std::fs::write(self.reply_path(), json.as_bytes()).is_ok() {
                self.update_cached_text(start, end, replace);
                self.activate_replace_freeze(find);
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
                r#"{{"action":"replace","tabId":{},"frameId":{},"start":{},"end":{},"text":"{}","expected":"{}","paragraphStart":{}}}"#,
                self.source_tab_id(), self.source_frame_id(), start, end, escaped_text, escaped_find, self.last_paragraph_start.get()
            );
            if std::fs::write(self.reply_path(), json.as_bytes()).is_ok() {
                self.update_cached_text(start, end, replace);
                self.activate_replace_freeze(find);
                return true;
            }
        }
        false
    }

    fn find_and_replace_in_context_at(&self, find: &str, replace: &str, context: &str, char_offset: usize) -> bool {
        let text = self.last_text.borrow().clone();
        if text.is_empty() { return false; }
        log_browser(&format!("REPLACE: find='{}' replace='{}' char_offset={}", find, replace, char_offset));
        log_browser(&format!("  cached text ({} chars): '{}'", text.chars().count(), text));
        // Use the char_offset to find the exact position
        // Find ALL occurrences and pick the one closest to char_offset
        let text_lower = text.to_lowercase();
        let find_lower = find.to_lowercase();
        let mut best_match: Option<(usize, usize)> = None; // (char_start, distance)
        let mut search_from = 0usize;
        while let Some(byte_pos) = text_lower[search_from..].find(&find_lower) {
            let abs_byte = search_from + byte_pos;
            let char_start = text[..abs_byte].chars().count();
            let dist = (char_start as isize - char_offset as isize).unsigned_abs();
            if best_match.is_none() || dist < best_match.unwrap().1 {
                best_match = Some((char_start, dist));
            }
            search_from = abs_byte + 1;
            while search_from < text_lower.len() && !text_lower.is_char_boundary(search_from) {
                search_from += 1;
            }
        }
        if let Some((start, _dist)) = best_match {
            let end = start + find.chars().count();
            log_browser(&format!("  FOUND at char {}..{}, sending replace JSON", start, end));
            log_browser(&format!("  text BEFORE replace: '{}'", text));
            let escaped = replace.replace('\\', "\\\\").replace('"', "\\\"");
            let find_escaped = find.replace('\\', "\\\\").replace('"', "\\\"");
            let json = format!(
                r#"{{"action":"replace","tabId":{},"frameId":{},"start":{},"end":{},"text":"{}","expected":"{}","paragraphStart":{}}}"#,
                self.source_tab_id(), self.source_frame_id(), start, end, escaped, find_escaped, self.last_paragraph_start.get()
            );
            log_browser(&format!("  reply JSON: {}", json));
            if std::fs::write(self.reply_path(), json.as_bytes()).is_ok() {
                self.update_cached_text(start, end, replace);
                self.activate_replace_freeze(find);
                let new_text = self.last_text.borrow().clone();
                log_browser(&format!("  text AFTER replace: '{}'", new_text));
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

        // Empty text is a meaningful state — the user just cleared the
        // editor (Cmd+A + Backspace in a Gmail compose / Reddit comment
        // box).  Previously we bailed with `None` here, which made the
        // desktop's update_grammar_errors + prune_resolved_errors run
        // never see the empty doc → writing_errors kept showing the old
        // misspellings over an empty input field. Reported 2026-05-19.
        //
        // Return an empty CursorContext (word/sentence/masked all blank)
        // so the desktop's "no word, no context" branch at main.rs:~6622
        // runs and triggers prune_resolved_errors, whose empty-doc
        // branch at main.rs:3335 clears writing_errors + queues.
        if text.is_empty() {
            crate::debug_log!("read_context: empty text — returning empty CursorContext to trigger desktop clear");
            return Some(CursorContext {
                word: String::new(),
                sentence: String::new(),
                masked_sentence: None,
                caret_pos: caret,
                cursor_doc_offset: Some(0),
                paragraph_id: String::new(),
            });
        }

        // Split text at cursor position (byte-safe)
        let cursor_byte = char_to_byte_offset(&text, cursor_start);
        let before = &text[..cursor_byte];
        let after = &text[cursor_byte..];

        let raw = RawCursorText {
            before: before.to_string(),
            after: after.to_string(),
        };

        crate::debug_log!("read_context: cursor_start={} text_len={} before_len={} after_len={}",
            cursor_start, text.len(), before.len(), after.len());
        let mut ctx = build_context(&raw, caret);
        ctx.cursor_doc_offset = Some(cursor_start);
        Some(ctx)
    }

    fn take_completed_word_for_tts(&self) -> Option<String> {
        // Pull a fresh payload before consuming the event. Google Docs
        // publishes its text after the OS key event, so reading here is what
        // makes speak-on-space follow the updated browser cursor.
        let _ = self.read_data_file();
        self.pending_completed_word.borrow_mut().take()
    }

    fn read_selected_text(&self) -> Option<String> {
        // Prefer the browser's own editor selection. On Windows and macOS,
        // platform selection caches can otherwise still contain the most
        // recent selection from Word or another native application.
        let _ = self.read_data_file();
        self.last_selected_text.borrow().clone()
    }

    fn read_full_document(&self) -> Option<String> {
        // Re-read file to get latest text. Pass empty text through (do not
        // collapse to None) so the desktop's try_update_doc_text sets
        // last_doc_text = "" and the empty-doc branch of
        // prune_resolved_errors clears stale writing_errors. Without
        // this the user saw old errors hover over a cleared text box
        // (reported 2026-05-19).
        let (text, _, _, _) = self.read_data_file()?;
        Some(text)
    }

    fn read_paragraph_at(&self, cursor_offset: usize) -> Option<(String, String, usize)> {
        let (text, cursor_start, _, _) = self.read_data_file()?;
        if text.trim().is_empty() {
            return None;
        }
        let cursor = if cursor_offset > 0 { cursor_offset } else { cursor_start };
        let (para_text, para_start) = paragraph_window_around_cursor(&text, cursor, 6000);
        if para_text.trim().is_empty() {
            return None;
        }
        Some((
            format!("browser:{}", self.last_paragraph_start.get()),
            para_text,
            para_start,
        ))
    }

    fn find_and_replace_in_paragraph(
        &self,
        find: &str,
        replace: &str,
        paragraph_id: &str,
        context: &str,
        char_offset: usize,
    ) -> bool {
        let current_paragraph_id = format!("browser:{}", self.last_paragraph_start.get());
        if paragraph_id != current_paragraph_id {
            log_browser(&format!(
                "REPLACE: ignoring stale browser paragraph '{}' (active '{}')",
                paragraph_id, current_paragraph_id
            ));
            return false;
        }
        self.find_and_replace_in_context_at(find, replace, context, char_offset)
    }

    /// Browser/textarea: skip the word at cursor if user is typing at end of document.
    /// When editing in the middle (cursor not at end), always check — user changed existing text.
    fn should_skip_word_spelling(&self, cursor_off: usize, word_start: usize, word_end: usize, doc_char_len: usize, word_at_cursor: &str) -> bool {
        // Only skip if cursor is at the very end of the document AND mid-word
        let at_end = cursor_off >= doc_char_len.saturating_sub(1);
        let mid_word = !word_at_cursor.is_empty() && word_at_cursor.chars().last().map(|c| c.is_alphanumeric()).unwrap_or(false);
        at_end && mid_word && cursor_off >= word_start && cursor_off <= word_end
    }

    /// Browser/textarea: skip grammar for the sentence at cursor if typing
    /// at end of an unpunctuated sentence — BUT ONLY when the mid-word IS
    /// the entire sentence content (i.e., user is typing the first word of
    /// a fresh sentence).  Once there are any completed words before the
    /// cursor we MUST dispatch the sentence to the grammar actor,
    /// otherwise the user's already-completed misspellings disappear from
    /// the pencil panel mid-typing.
    ///
    /// Reported 2026-05-19: "While writing in Reddit comment section, if I
    /// am writing a word and lets say I stopped typing our app does not
    /// show anything, not even the previous errors until I press space."
    /// Root cause: process_spelling_queue is a no-op (line ~4005 returns
    /// immediately) — spelling is delivered via the grammar actor's
    /// unknown-words response.  Skipping the grammar dispatch for the
    /// whole sentence whenever the cursor was mid-word blocked that path,
    /// so misspellings for completed words never reached writing_errors
    /// until the user pressed space and the cursor left mid-word state.
    fn should_skip_sentence_grammar(&self, cursor_off: usize, sent_start: usize, _sent_end: usize, ends_with_punct: bool, doc_char_len: usize, word_at_cursor: &str) -> bool {
        let at_end = cursor_off >= doc_char_len.saturating_sub(1);
        let mid_word = !word_at_cursor.is_empty()
            && word_at_cursor.chars().last().map(|c| c.is_alphanumeric()).unwrap_or(false);
        // Position where the current mid-word started.
        let word_start_pos = cursor_off.saturating_sub(word_at_cursor.chars().count());
        // True only if there is no completed text before the mid-word in
        // this sentence (user is partway through the FIRST word of a
        // fresh sentence).  When this is false there are real words
        // before the cursor whose errors the user wants to see — we must
        // not skip.
        let is_only_word_in_sentence = word_start_pos <= sent_start;
        at_end && mid_word && !ends_with_punct && is_only_word_in_sentence
    }
}

/// Convert a character offset to a byte offset in a UTF-8 string
fn char_to_byte_offset(s: &str, char_offset: usize) -> usize {
    s.char_indices()
        .nth(char_offset)
        .map(|(byte_idx, _)| byte_idx)
        .unwrap_or(s.len())
}

fn completed_word_from_transition(
    previous_text: &str,
    previous_cursor: usize,
    current_text: &str,
    current_cursor: usize,
) -> Option<String> {
    if previous_text == current_text || current_cursor == 0 {
        return None;
    }

    let current_byte = char_to_byte_offset(current_text, current_cursor);
    let current_before = &current_text[..current_byte];
    let inserted_space = matches!(current_before.chars().next_back(), Some(' ' | '\u{00a0}'));
    if !inserted_space {
        return None;
    }

    let previous_byte = char_to_byte_offset(previous_text, previous_cursor);
    let previous_before = &previous_text[..previous_byte];
    if matches!(previous_before.chars().next_back(), Some(' ' | '\u{00a0}')) {
        return None;
    }

    let word = extract_previous_word_before_cursor(current_before);
    (!word.is_empty()).then_some(word)
}

fn paragraph_window_around_cursor(text: &str, cursor_char: usize, max_chars: usize) -> (String, usize) {
    let cursor_char = cursor_char.min(text.chars().count());
    let cursor_byte = char_to_byte_offset(text, cursor_char);
    let para_start_byte = text[..cursor_byte]
        .rfind('\n')
        .map(|pos| pos + 1)
        .unwrap_or(0);
    let para_end_byte = text[cursor_byte..]
        .find('\n')
        .map(|pos| cursor_byte + pos)
        .unwrap_or(text.len());
    let para_start_char = text[..para_start_byte].chars().count();
    let para_end_char = text[..para_end_byte].chars().count();

    if para_end_char.saturating_sub(para_start_char) <= max_chars {
        return (text[para_start_byte..para_end_byte].to_string(), para_start_char);
    }

    let half = max_chars / 2;
    let mut start_char = cursor_char.saturating_sub(half).max(para_start_char);
    let mut end_char = start_char.saturating_add(max_chars).min(para_end_char);
    if end_char.saturating_sub(start_char) < max_chars {
        start_char = end_char.saturating_sub(max_chars).max(para_start_char);
    }
    end_char = end_char.max(start_char);
    let start_byte = char_to_byte_offset(text, start_char);
    let end_byte = char_to_byte_offset(text, end_char);
    (text[start_byte..end_byte].to_string(), start_char)
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
                    Some('u') => {
                        // Parse \uXXXX unicode escape
                        let hex: String = (0..4).filter_map(|_| chars.next()).collect();
                        if let Ok(code) = u32::from_str_radix(&hex, 16) {
                            if let Some(ch) = char::from_u32(code) {
                                if !ch.is_control() {
                                    result.push(ch);
                                }
                                // Skip control characters silently
                            }
                        }
                    }
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

/// Check if a whole word exists in text (not as substring of a longer word)
fn has_whole_word(text: &str, word: &str) -> bool {
    let mut pos = 0;
    while let Some(idx) = text[pos..].find(word) {
        let abs = pos + idx;
        let before_ok = abs == 0 || !text[..abs].ends_with(|c: char| c.is_alphanumeric());
        let after_pos = abs + word.len();
        let after_ok = after_pos >= text.len() || !text[after_pos..].starts_with(|c: char| c.is_alphanumeric());
        if before_ok && after_ok {
            return true;
        }
        pos = abs + 1;
    }
    false
}

#[cfg(test)]
mod tests {
    use super::{BrowserBridge, completed_word_from_transition, extract_json_string, reply_path_for};
    use crate::bridge::TextBridge;

    #[test]
    fn completed_word_is_emitted_after_fresh_space_payload() {
        assert_eq!(
            completed_word_from_transition("hello", 5, "hello ", 6),
            Some("hello".to_string())
        );
        assert_eq!(
            completed_word_from_transition("hei", 3, "hei\u{00a0}", 4),
            Some("hei".to_string())
        );
    }

    #[test]
    fn completed_word_handles_google_docs_poll_skips() {
        assert_eq!(
            completed_word_from_transition("h", 1, "hello world ", 12),
            Some("world".to_string())
        );
    }

    #[test]
    fn completed_word_ignores_keepalives_and_cursor_moves() {
        assert_eq!(
            completed_word_from_transition("hello ", 6, "hello ", 6),
            None
        );
        assert_eq!(
            completed_word_from_transition("hello world", 5, "hello world", 6),
            None
        );
        assert_eq!(
            completed_word_from_transition("hello ", 6, "hello  ", 7),
            None
        );
    }

    #[test]
    fn initialized_empty_browser_payload_remains_readable() {
        let bridge = BrowserBridge::new();
        bridge.last_modified.set(1);
        bridge.last_cursor.set(0);
        bridge.last_text.borrow_mut().clear();

        let (text, start, end, _) = bridge.cached_data().expect("empty payload is valid");
        assert!(text.is_empty());
        assert_eq!((start, end), (0, 0));
    }

    #[test]
    fn browser_selected_text_uses_extension_payload() {
        let json = r#"{"text":"Jeg liker piza.","selectedText":"liker piza","cursorStart":14}"#;
        assert_eq!(
            extract_json_string(json, "selectedText"),
            Some("liker piza".to_string())
        );

        let bridge = BrowserBridge::new();
        *bridge.last_selected_text.borrow_mut() = Some("liker piza".to_string());
        assert_eq!(bridge.read_selected_text(), Some("liker piza".to_string()));
    }

    #[test]
    fn replacement_replies_target_source_tab() {
        let bridge = BrowserBridge::new();
        *bridge.last_text.borrow_mut() = "Han gik til skolen.".to_string();
        *bridge.last_source.borrow_mut() = "42|https://docs.google.com/document/d/test".to_string();
        bridge.last_bridge_id.set(1234);
        bridge.last_frame_id.set(7);
        bridge.last_cursor.set(7);
        bridge.last_paragraph_start.set(19);
        let reply = reply_path_for(1234);
        let _ = std::fs::remove_file(&reply);

        assert!(bridge.replace_word("gik|gikk"));
        assert_eq!(
            std::fs::read_to_string(&reply).expect("browser replacement reply"),
            r#"{"action":"replace","tabId":42,"frameId":7,"start":4,"end":7,"text":"gikk","expected":"gik","paragraphStart":19}"#,
        );
        assert_eq!(&*bridge.last_text.borrow(), "Han gikk til skolen.");

        *bridge.last_text.borrow_mut() = "Jeg liker piza.".to_string();
        assert!(bridge.find_and_replace_in_paragraph(
            "piza",
            "pipa",
            "browser:19",
            "Jeg liker piza.",
            10,
        ));
        assert_eq!(
            std::fs::read_to_string(&reply).expect("browser correction reply"),
            r#"{"action":"replace","tabId":42,"frameId":7,"start":10,"end":14,"text":"pipa","expected":"piza","paragraphStart":19}"#,
        );
        assert_eq!(&*bridge.last_text.borrow(), "Jeg liker pipa.");
        assert!(!bridge.find_and_replace_in_paragraph(
            "pipa",
            "pizza",
            "browser:0",
            "Jeg liker pipa.",
            10,
        ));
        assert_eq!(&*bridge.last_text.borrow(), "Jeg liker pipa.");
        assert!(reply.file_name().unwrap().to_string_lossy().contains("reply-1234"));

        let _ = std::fs::remove_file(reply);
    }
}
