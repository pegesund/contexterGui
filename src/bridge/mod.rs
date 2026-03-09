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

    /// Set the target window handle for cross-process reads (no-op by default).
    fn set_target_hwnd(&self, _hwnd: isize) {}
}

#[cfg(target_os = "windows")]
pub mod word_com;

#[cfg(target_os = "windows")]
pub mod accessibility_win;

// Future:
// #[cfg(target_os = "macos")]
// pub mod word_applescript;
// #[cfg(target_os = "macos")]
// pub mod accessibility_mac;

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
        return None;
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

    let sentence = find_sentence_around_cursor(&raw.before, &raw.after);

    let masked = if !word.is_empty() {
        build_masked_sentence(raw, &word)
    } else {
        None
    };

    CursorContext {
        word,
        sentence,
        masked_sentence: masked,
        caret_pos,
    }
}
