/// Text context around the cursor in the active application.
#[derive(Debug, Clone, Default)]
pub struct CursorContext {
    pub word: String,
    pub sentence: String,
    /// When clicking mid-word: full sentence with the word replaced by <mask>
    /// Used for BERT fill-in-the-blank prediction
    pub masked_sentence: Option<String>,
    /// Screen coordinates of the caret (x, y below the caret)
    pub caret_pos: Option<(i32, i32)>,
    /// Character offset of the cursor in the document (for error range detection)
    pub cursor_doc_offset: Option<usize>,
    /// Paragraph ID from Word Add-in (stable identifier for error tracking)
    pub paragraph_id: String,
}

/// Underline color for error marking.
#[derive(Clone, Copy)]
pub enum ErrorUnderlineColor {
    Red,   // Spelling errors
    Blue,  // Grammar errors
}

/// Abstraction for reading/writing text in the focused application.
///
/// Implementations:
/// - WordComBridge (Windows) — Word COM automation, best experience
/// - WordApplescriptBridge (Mac) — Word AppleScript (future)
/// - AccessibilityBridge — fallback for any app via OS accessibility API
pub trait TextBridge {
    /// Human-readable name for this bridge (e.g. "Word COM", "Accessibility")
    fn name(&self) -> &str;

    /// Can this bridge connect to the currently focused application?
    fn is_available(&self) -> bool;

    /// Read the current word, sentence, and caret position.
    fn read_context(&self) -> Option<CursorContext>;

    /// Replace the current word at the cursor with new text.
    fn replace_word(&self, new_text: &str) -> bool;

    /// Find a word in the document and replace it. Returns true if replaced.
    fn find_and_replace(&self, _find: &str, _replace: &str) -> bool { false }

    /// Find a word within a specific sentence context and replace it.
    fn find_and_replace_in_context(&self, _find: &str, _replace: &str, _context: &str) -> bool { false }

    /// Find a word within a specific sentence at a known char offset in the document.
    /// Default: falls back to find_and_replace_in_context (ignoring offset).
    fn find_and_replace_in_context_at(&self, find: &str, replace: &str, context: &str, _char_offset: usize) -> bool {
        self.find_and_replace_in_context(find, replace, context)
    }

    /// Read a larger text window for context (e.g. 5000 chars before cursor).
    /// Used for sentence embeddings / topic extraction.
    fn read_document_context(&self) -> Option<String> { None }

    /// Read the full document text (for grammar checking all sentences).
    fn read_full_document(&self) -> Option<String> { None }

    /// Select/highlight a character range in the document (for navigating to errors).
    /// Returns optional (x, y) screen position of the selection.
    fn select_range(&self, _char_start: usize, _char_end: usize) -> Option<(i32, i32)> { None }

    /// Select/highlight a word within a specific paragraph (for navigating to errors in add-in).
    fn select_word_in_paragraph(&self, _word: &str, _paragraph_id: &str) -> bool { false }

    /// Set the target window handle for cross-process reads (no-op by default).
    fn set_target_hwnd(&self, _hwnd: isize) {}

    /// Set the foreground window handle (called by BridgeManager before read_context).
    fn set_fg_hwnd(&self, _hwnd: isize) {}

    /// Mark a character range with wavy underline. Color: red for spelling, blue for grammar.
    fn mark_error_underline(&self, _char_start: usize, _char_end: usize, _color: ErrorUnderlineColor) -> bool { false }

    /// Remove red wavy underline from a character range.
    fn clear_error_underline(&self, _char_start: usize, _char_end: usize) -> bool { false }

    /// Clear ALL error underlines in the document (cleanup on exit).
    fn clear_all_error_underlines(&self) -> bool { false }

    /// Read all paragraphs with stable IDs. Returns (para_id, text, char_start).
    /// Used for incremental document scanning — only changed paragraphs need processing.
    fn read_paragraphs(&self) -> Option<Vec<(String, String, usize)>> { None }

    /// Read the single paragraph at cursor position. Returns (para_id, text, char_start).
    fn read_paragraph_at(&self, _cursor_offset: usize) -> Option<(String, String, usize)> { None }

    /// Read cached selected text (polled while external app had focus).
    fn read_selected_text(&self) -> Option<String> { None }

    /// Mark a word with colored wavy underline by searching within a paragraph.
    fn underline_word(&self, _word: &str, _paragraph_id: &str, _color: &str) -> bool { false }

    /// Clear underline from a word by searching within a paragraph.
    fn clear_underline_word(&self, _word: &str, _paragraph_id: &str) -> bool { false }
    /// Clear ALL underlines in a paragraph (used when paragraph content changes).
    fn clear_paragraph_underlines(&self, _paragraph_id: &str) -> bool { false }

    /// Push a raw JSON command to the add-in reply queue.
    fn push_reply(&self, _json: &str) {}

    /// Push a command to the FRONT of the reply queue (priority).
    fn push_reply_urgent(&self, _json: &str) {}

    // ── Per-bridge error scanning behavior ──
    // Each bridge controls WHEN spelling/grammar checks run.
    // Default: always check (safe for Word COM, Accessibility).
    // Browser overrides to skip while typing at end of document.

    /// Should this word be skipped for spelling checks?
    /// Called for each word during document scanning.
    /// `cursor_off`: cursor position in document (char offset)
    /// `word_start`/`word_end`: word position in document (char offsets)
    /// `doc_char_len`: total document length in chars
    fn should_skip_word_spelling(&self, _cursor_off: usize, _word_start: usize, _word_end: usize, _doc_char_len: usize, _word_at_cursor: &str) -> bool { false }

    /// Should this sentence be skipped for grammar checks?
    /// `cursor_off`: cursor position in document (char offset)
    /// `sent_start`/`sent_end`: sentence position in document (char offsets)
    /// `ends_with_punct`: whether the sentence ends with .!?
    /// `doc_char_len`: total document length in chars
    fn should_skip_sentence_grammar(&self, _cursor_off: usize, _sent_start: usize, _sent_end: usize, _ends_with_punct: bool, _doc_char_len: usize, _word_at_cursor: &str) -> bool { false }

    /// Drain changed sentences detected by the bridge (e.g. Word Add-in paragraph events).
    /// Default: empty (bridges that don't track sentence changes return nothing).
    fn drain_changed_paragraphs(&self) -> Vec<word_addin::ChangedParagraph> { vec![] }

    /// Drain deleted paragraph IDs detected by the bridge (e.g. Word Add-in paragraph delete events).
    /// Default: empty (bridges that don't track paragraph deletions return nothing).
    fn drain_deleted_paragraphs(&self) -> Vec<String> { vec![] }

    /// Check if a reset was requested (new document opened). Default: false.
    fn take_reset(&self) -> bool { false }

    /// Update errors JSON for /errors test endpoint
    fn update_errors_json(&self, _json: &str) {}
}

#[cfg(target_os = "windows")]
pub mod word_com;

#[cfg(target_os = "windows")]
pub mod accessibility_win;

pub mod word_addin;

pub mod browser;

// Future:
// #[cfg(target_os = "macos")]
// pub mod accessibility_mac;

/// Create platform-specific bridges (excluding Browser, which is added separately).
pub fn create_bridges(lang_word_id: i32) -> Vec<Box<dyn TextBridge>> {
    let mut bridges: Vec<Box<dyn TextBridge>> = Vec::new();

    #[cfg(target_os = "windows")]
    {
        if let Some(word) = word_com::WordComBridge::try_connect(lang_word_id) {
            crate::log!("Word COM bridge connected");
            let ok = word.disable_word_proofing();
            crate::log!("Word proofing disabled: {}", ok);
            bridges.push(Box::new(word));
        }
        bridges.push(Box::new(accessibility_win::AccessibilityBridge::new()));
    }

    #[cfg(target_os = "macos")]
    {
        // Word Add-in bridge: HTTP server that Word's JS add-in connects to.
        // Always start — the add-in connects when it's ready.
        let addin_bridge = word_addin::WordAddinBridge::new();
        crate::log!("Word Add-in bridge started (HTTP port {})", 52525);
        bridges.push(Box::new(addin_bridge));
    }

    bridges
}

/// Try to connect a Word bridge (for late connection when Word opens after app startup).
pub fn try_connect_word_bridge(lang_word_id: i32) -> Vec<Box<dyn TextBridge>> {
    let mut bridges: Vec<Box<dyn TextBridge>> = Vec::new();

    #[cfg(target_os = "windows")]
    {
        if let Some(word) = word_com::WordComBridge::try_connect(lang_word_id) {
            let ok = word.disable_word_proofing();
            crate::log!("Word proofing disabled (late): {}", ok);
            bridges.push(Box::new(word));
        }
    }

    // On macOS, Word Add-in bridge is always running (HTTP server).
    // No late connection needed — add-in connects when loaded.

    bridges
}

// ── Shared text processing used by all bridges ──

/// Raw text around cursor, provided by each bridge in its own way.
/// The shared functions below operate on this.
pub struct RawCursorText {
    /// Text before cursor (up to ~2000 chars)
    pub before: String,
    /// Text after cursor (up to ~2000 chars)
    pub after: String,
}

/// Extract the word being typed by scanning backwards from cursor.
pub fn extract_word_before_cursor(before: &str) -> String {
    before
        .chars()
        .rev()
        .take_while(|c| c.is_alphanumeric() || *c == '-' || *c == '\'')
        .collect::<Vec<_>>()
        .into_iter()
        .rev()
        .collect::<String>()
        .trim()
        .to_string()
}

/// Extract the rest of the word after cursor (when caret is mid-word).
pub fn extract_word_after_cursor(after: &str) -> String {
    after
        .chars()
        .take_while(|c| c.is_alphanumeric() || *c == '-' || *c == '\'')
        .collect::<String>()
}

/// Find the sentence around the cursor from before+after text.
pub fn find_sentence_around_cursor(before: &str, after: &str) -> String {
    let sentence_end_chars = ['.', '!', '?'];

    // Scan backwards for sentence start
    let sent_start = before
        .rfind(|c: char| sentence_end_chars.contains(&c))
        .map(|pos| {
            // Skip past the punctuation + following whitespace
            let rest = &before[pos + 1..];
            let skip = rest.len() - rest.trim_start().len();
            pos + 1 + skip
        })
        .unwrap_or(0);
    let before_part = &before[sent_start..];

    // Scan forwards for sentence end
    let sent_end = after
        .find(|c: char| sentence_end_chars.contains(&c))
        .map(|pos| pos + 1) // include the punctuation
        .unwrap_or(after.len());
    let after_part = &after[..sent_end];

    format!("{}{}", before_part, after_part)
        .replace('\r', " ")
        .replace('\n', " ")
        .trim()
        .to_string()
}

/// Build masked sentence for BERT fill-in-the-blank.
/// `before` = text before cursor, `after` = text after cursor, `word` = word at cursor.
/// Result: context with word replaced by `<mask>`.
pub fn build_masked_sentence(raw: &RawCursorText, word: &str) -> Option<String> {
    if word.is_empty() {
        // No word at cursor (e.g. just pressed space) — place mask at cursor position
        let before = raw.before.trim_end();
        if before.is_empty() {
            return None;
        }
        let after = raw.after.trim_start();
        let after_part = if after.is_empty() { ".".to_string() } else { after.to_string() };
        return Some(format!("{} <mask> {}", before, after_part));
    }
    // Strip the word from before text (it's the suffix being typed)
    let before_trimmed = if raw.before.ends_with(word) {
        &raw.before[..raw.before.len() - word.len()]
    } else {
        // Partial word — strip matching suffix
        let word_before = extract_word_before_cursor(&raw.before);
        &raw.before[..raw.before.len() - word_before.len()]
    };
    let word_after = extract_word_after_cursor(&raw.after);
    let after_trimmed = if raw.after.starts_with(&word_after) {
        &raw.after[word_after.len()..]
    } else {
        &raw.after
    };

    let masked = format!(
        "{}<mask> {}",
        before_trimmed.trim_end_matches(|c: char| c.is_whitespace()),
        after_trimmed.trim_start()
    );
    Some(masked)
}

/// Build CursorContext from raw text around cursor.
pub fn build_context(raw: &RawCursorText, caret_pos: Option<(i32, i32)>) -> CursorContext {
    let word_before = extract_word_before_cursor(&raw.before);
    let word_after = extract_word_after_cursor(&raw.after);
    let word = format!("{}{}", word_before, word_after);

    let mut sentence = find_sentence_around_cursor(&raw.before, &raw.after);

    // If sentence is empty but we have text (e.g. cursor right after final period),
    // use the last complete sentence from the before text.
    if sentence.is_empty() && !raw.before.trim().is_empty() {
        let trimmed = raw.before.trim_end();
        // Strip trailing punctuation to find the previous sentence boundary
        let without_final = trimmed.trim_end_matches(|c: char| c == '.' || c == '!' || c == '?');
        if !without_final.is_empty() {
            let prev_end = without_final.rfind(|c: char| c == '.' || c == '!' || c == '?');
            let start = prev_end.map(|p| p + 1).unwrap_or(0);
            sentence = trimmed[start..].trim().to_string();
        } else {
            sentence = trimmed.to_string();
        }
    }

    let masked = build_masked_sentence(raw, &word);

    CursorContext {
        word,
        sentence,
        masked_sentence: masked,
        caret_pos,
        cursor_doc_offset: None,
        paragraph_id: String::new(),
    }
}
