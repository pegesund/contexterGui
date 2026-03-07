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

    /// Read a larger text window for context (e.g. 5000 chars before cursor).
    /// Used for sentence embeddings / topic extraction.
    fn read_document_context(&self) -> Option<String> { None }
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
