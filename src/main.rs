#[macro_use]
pub mod logging;

mod bert_worker;
mod bridge;
mod grammar_actor;
mod ocr;
mod platform;
mod stt;
mod tts;

use bridge::{CursorContext, TextBridge};
use std::collections::HashMap;
use std::hash::{Hash, Hasher};
use std::io::Write;
use std::path::PathBuf;
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

/// Truncate a string to at most `max` bytes, backing up to the nearest char boundary.
fn trunc(s: &str, max: usize) -> &str {
    if s.len() <= max { return s; }
    let mut end = max;
    while end > 0 && !s.is_char_boundary(end) { end -= 1; }
    &s[..end]
}

use nostos_cognio::baseline::{compute_baseline, Baselines};
use nostos_cognio::complete::{complete_word, grammar_filter, GrammarCheckResult, Completion};
use nostos_cognio::embeddings::EmbeddingStore;
use nostos_cognio::grammar::GrammarChecker;
use nostos_cognio::grammar::swipl_checker::SwiGrammarChecker;
use nostos_cognio::grammar::types::GrammarError;
use nostos_cognio::model::Model;
use nostos_cognio::prefix_index::{self, PrefixIndex};
use nostos_cognio::wordfreq;

// --- Grammar checker abstraction ---

pub(crate) enum AnyChecker {
    Neo(GrammarChecker),
    Swi(SwiGrammarChecker),
}

// SAFETY: AnyChecker is only ever accessed from one thread at a time.
// SWI-Prolog's raw pointers (PredicateT) are !Send, but the grammar actor
// ensures single-threaded access via mpsc channel serialization.
unsafe impl Send for AnyChecker {}

impl AnyChecker {
    fn has_word(&self, word: &str) -> bool {
        match self {
            AnyChecker::Neo(c) => c.has_word(word),
            AnyChecker::Swi(c) => c.has_word(word),
        }
    }

    fn prefix_lookup(&self, prefix: &str, limit: usize) -> Vec<String> {
        match self {
            AnyChecker::Neo(c) => c.prefix_lookup(prefix, limit),
            AnyChecker::Swi(c) => c.prefix_lookup(prefix, limit),
        }
    }

    fn check_sentence(&mut self, text: &str) -> Vec<GrammarError> {
        match self {
            AnyChecker::Neo(c) => c.check_sentence(text),
            AnyChecker::Swi(c) => c.check_sentence(text),
        }
    }

    fn has_error(&mut self, text: &str) -> bool {
        match self {
            AnyChecker::Neo(c) => !c.check_sentence(text).is_empty(),
            AnyChecker::Swi(c) => c.has_error(text),
        }
    }

    fn check_sentence_full(&mut self, text: &str) -> nostos_cognio::grammar::types::CheckResult {
        match self {
            AnyChecker::Neo(c) => c.check_sentence_full(text),
            AnyChecker::Swi(c) => c.check_sentence_full(text),
        }
    }

    fn fuzzy_lookup(&self, word: &str, max_distance: u32) -> Vec<(String, u32)> {
        match self {
            AnyChecker::Neo(c) => c.fuzzy_lookup(word, max_distance),
            AnyChecker::Swi(c) => c.fuzzy_lookup(word, max_distance),
        }
    }

    fn suggest_compound(&self, word: &str) -> Option<String> {
        match self {
            AnyChecker::Swi(c) => c.suggest_compound(word),
            AnyChecker::Neo(_) => None,
        }
    }

    /// Get the set of POS tags for a word from the dictionary
    fn pos_set(&self, word: &str) -> std::collections::HashSet<String> {
        let analyzer = match self {
            AnyChecker::Neo(c) => c.analyzer().clone(),
            AnyChecker::Swi(c) => c.analyzer().clone(),
        };
        let mut pos = std::collections::HashSet::new();
        if let Some(readings) = analyzer.dict_lookup(word) {
            for r in &readings {
                pos.insert(r.pos.to_string());
            }
        }
        pos
    }

    /// Split unpunctuated text into sentences using Prolog sentence boundary detection.
    /// Returns None if not using SWI checker or no boundaries found.
    /// Validates that each resulting sub-sentence has at least one likely verb —
    /// rejects splits that produce verbless fragments like "I huset." or "Kaker og brus."
    fn split_by_prolog(&mut self, text: &str) -> Option<Vec<String>> {
        match self {
            AnyChecker::Swi(c) => {
                let sentences = nostos_cognio::punctuation::split_by_prolog(c, text);
                if sentences.len() <= 1 {
                    return None;
                }
                // Validate: every sub-sentence must have at least one likely verb.
                let analyzer = c.analyzer().clone();
                for sent in &sentences {
                    let stripped = sent.trim_end_matches(|ch: char| ch == '.' || ch == '!' || ch == '?').trim();
                    if !nostos_cognio::punctuation::has_likely_verb_in_sentence(&analyzer, stripped) {
                        eprintln!("Grammar: rejecting Prolog split — '{}' has no likely verb", stripped);
                        return None;
                    }
                }
                // Validate: every sub-sentence must pass grammar check.
                // Rejects splits like "Jeg liker." + "Å gikk tur." where sub-sentences have errors.
                for sent in &sentences {
                    let errors = c.check_sentence(sent);
                    if !errors.is_empty() {
                        eprintln!("Grammar: rejecting Prolog split — '{}' has {} grammar errors", sent, errors.len());
                        return None;
                    }
                }
                Some(sentences)
            }
            _ => None,
        }
    }

}

// --- Error list for spelling and grammar ---

#[derive(Clone, Debug)]
pub(crate) enum ErrorCategory {
    Spelling,
    Grammar,
    SentenceBoundary,
}

#[derive(Clone, Debug)]
pub(crate) struct WritingError {
    pub(crate) category: ErrorCategory,
    pub(crate) word: String,
    pub(crate) suggestion: String,
    pub(crate) explanation: String,
    pub(crate) rule_name: String,
    /// The sentence text containing the error
    pub(crate) sentence_context: String,
    /// Character offset of the sentence in the document (for position-aware duplicate handling)
    pub(crate) doc_offset: usize,
    /// Alternative index (0 = primary, >0 = secondary alternatives for grammar)
    pub(crate) position: usize,
    /// true if user clicked "Ignorer"
    pub(crate) ignored: bool,
    /// Absolute char offset of the error word in the document (for underline marking)
    pub(crate) word_doc_start: usize,
    /// Absolute char end offset of the error word in the document
    pub(crate) word_doc_end: usize,
    /// Whether we've applied a red wavy underline for this error
    pub(crate) underlined: bool,
    /// Pinned to top of error list (newly found at word boundary)
    pub(crate) pinned: bool,
    /// Paragraph ID from Word Add-in (for error removal when sentence changes)
    pub(crate) paragraph_id: String,
}

#[derive(Clone)]
struct SpellingQueueItem {
    word: String,
    sentence_ctx: String,
    paragraph_id: String,
}

/// Find a word within a sentence and return (doc_start, doc_end) in absolute char offsets.
/// `error_word` = the word to find, `sentence_ctx` = the sentence text,
/// `doc_offset` = char offset of the sentence in the document.
fn find_word_doc_range(error_word: &str, sentence_ctx: &str, doc_offset: usize) -> (usize, usize) {
    let word_lower = error_word.to_lowercase();
    let sent_lower = sentence_ctx.to_lowercase();
    // Find word at a word boundary in the sentence
    let chars: Vec<char> = sent_lower.chars().collect();
    let word_chars: Vec<char> = word_lower.chars().collect();
    let wlen = word_chars.len();
    for i in 0..chars.len().saturating_sub(wlen.saturating_sub(1)) {
        if i > 0 && chars[i - 1].is_alphanumeric() { continue; }
        let end = i + wlen;
        if end < chars.len() && chars[end].is_alphanumeric() { continue; }
        if &chars[i..end] == &word_chars[..] {
            return (doc_offset + i, doc_offset + end);
        }
    }
    // Fallback: whole sentence range
    (doc_offset, doc_offset + sentence_ctx.chars().count())
}

// --- Data paths ---

fn data_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../contexter-repo/training-data")
}

fn dict_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../rustSpell/mtag-rs/data/fullform_bm.mfst")
}

fn compound_data_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../syntaxer/compound_data.pl")
}

fn grammar_rules_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../syntaxer/grammar_rules.pl")
}

fn syntaxer_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../syntaxer")
}

// --- Bridge manager: picks the best available bridge ---

struct BridgeManager {
    bridges: Vec<Box<dyn TextBridge>>,
    last_check: Instant,
    /// Index of the bridge that last successfully read context
    active_idx: usize,
    /// PID of the last app we successfully read text from (to avoid switching to terminals etc.)
    last_user_pid: u32,
    /// True when the last user-focused app was a browser — NEVER activate Word COM in this state
    last_user_was_browser: bool,
    /// Last successfully read context (returned when our window is foreground)
    last_context: Option<CursorContext>,
    /// Set when bridge switches — main loop should clear stale errors
    bridge_switched: bool,
    /// Platform abstraction for OS-specific services
    platform: Box<dyn platform::PlatformServices>,
}

impl BridgeManager {
    fn new(platform: Box<dyn platform::PlatformServices>) -> Self {
        let mut bridges: Vec<Box<dyn TextBridge>> = bridge::create_bridges();
        // Browser bridge (via Chrome/Edge extension) — highest priority for browser textareas
        bridges.push(Box::new(bridge::browser::BrowserBridge::new()));

        // If the browser data file exists and is recent, assume user was in a browser.
        // This prevents Word COM from being read on the very first frame before
        // foreground detection has a chance to set the flag.
        let browser_file_exists = {
            let path = std::env::temp_dir().join("norsktale-browser.json");
            std::fs::metadata(&path).ok()
                .and_then(|m| m.modified().ok())
                .and_then(|t| t.elapsed().ok())
                .map(|age| age.as_secs() < 120) // file modified in last 2 minutes
                .unwrap_or(false)
        };
        if browser_file_exists {
            log!("Browser data file found (< 2min old) — defaulting to browser mode");
        }

        BridgeManager {
            bridges,
            last_check: Instant::now(),
            active_idx: 0,
            last_user_pid: 0,
            last_user_was_browser: browser_file_exists,
            last_context: None,
            bridge_switched: false,
            platform,
        }
    }

    fn read_context(&mut self) -> Option<CursorContext> {
        if self.last_check.elapsed() > Duration::from_secs(5) {
            self.last_check = Instant::now();
            let has_word = self.bridges.iter().any(|b| b.name().contains("Word"));
            if !has_word {
                for new_bridge in bridge::try_connect_word_bridge() {
                    log!("{} bridge connected (late)", new_bridge.name());
                    self.bridges.insert(0, new_bridge);
                }
            }
        }

        // Detect which app the user clicked on via platform abstraction.
        // When our own always-on-top window is foreground, keep reading from the
        // last known user app (user is just looking at our UI, caret is still
        // in the other app).
        let fg = self.platform.foreground_app();
        let fg_hwnd_raw = fg.handle;
        let fg_pid = fg.pid;
        let fg_title = fg.title.clone();
        let app_kind = self.platform.classify_app(&fg);
        let our_window_focused = app_kind == platform::AppKind::OurApp;
        let word_is_foreground = app_kind == platform::AppKind::Word;
        let is_browser = app_kind == platform::AppKind::Browser;

        // Log every focus change (not just every 3 seconds)
        static LAST_FG_PID: std::sync::atomic::AtomicU32 = std::sync::atomic::AtomicU32::new(0);
        let prev_fg = LAST_FG_PID.load(std::sync::atomic::Ordering::Relaxed);
        if fg_pid != prev_fg {
            LAST_FG_PID.store(fg_pid, std::sync::atomic::Ordering::Relaxed);
            log!("FG: '{}' pid={} exe='{}' our={} word={} browser={} last_user={}", trunc(&fg_title, 40), fg_pid, fg.exe_name, our_window_focused, word_is_foreground, is_browser, self.last_user_pid);
        }

        // Word Add-in bridge is data-driven (HTTP POST), not foreground-driven.
        // Check it first — if it has fresh data, use it regardless of foreground.
        for (i, bridge) in self.bridges.iter().enumerate() {
            if bridge.name() == "Word Add-in" {
                if let Some(ctx) = bridge.read_context() {
                    if !ctx.word.is_empty() || !ctx.sentence.is_empty() {
                        if self.active_idx != i {
                            log!("Bridge switch: {} → Word Add-in", self.bridges[self.active_idx].name());
                            self.bridge_switched = true;
                        }
                        self.active_idx = i;
                        self.last_user_pid = fg_pid;
                        self.last_context = Some(ctx.clone());
                        return Some(ctx);
                    }
                }
                break;
            }
        }

        // Our window is foreground — keep using whatever bridge was already active.
        // NEVER switch bridges when our window is focused.
        if our_window_focused {
            return self.last_context.clone();
        }

        // Only activate for supported programs: Word, Edge, Chrome, Notepad
        let is_notepad = app_kind == platform::AppKind::Notepad;
        let is_supported = word_is_foreground || is_browser || is_notepad;
        if !is_supported {
            return self.last_context.clone();
        }

        // --- BROWSER: ONLY use Browser bridge. No fallbacks. Ever. ---
        if is_browser {
            self.last_user_was_browser = true;
            if let Some(browser_idx) = self.bridges.iter().position(|b| b.name() == "Browser") {
                if let Some(ctx) = self.bridges[browser_idx].read_context() {
                    if !ctx.word.is_empty() || !ctx.sentence.is_empty() {
                        if self.active_idx != browser_idx {
                            log!("Bridge switch: {} → Browser", self.bridges[self.active_idx].name());
                            self.bridge_switched = true;
                        }
                        self.active_idx = browser_idx;
                        self.last_user_pid = fg_pid;
                        self.last_context = Some(ctx.clone());
                        return Some(ctx);
                    }
                }
                // Browser bridge has no data — just return last context. NO fallback.
                return self.last_context.clone();
            }
            // No Browser bridge exists — return last context. NO fallback.
            return self.last_context.clone();
        }

        // --- WORD ---
        if word_is_foreground {
            self.last_user_was_browser = false;
            for (i, bridge) in self.bridges.iter().enumerate() {
                if bridge.name().contains("Word") {
                    if let Some(ctx) = bridge.read_context() {
                        if !ctx.word.is_empty() || !ctx.sentence.is_empty() {
                            if self.active_idx != i {
                                log!("Bridge switch: {} → Word COM", self.bridges[self.active_idx].name());
                                self.bridge_switched = true;
                            }
                            self.active_idx = i;
                            self.last_user_pid = fg_pid;
                            self.last_context = Some(ctx.clone());
                            return Some(ctx);
                        }
                    }
                    break;
                }
            }
            return self.last_context.clone();
        }

        // --- NOTEPAD: uses Accessibility bridge ---
        if is_notepad {
            self.last_user_was_browser = false;
            for bridge in self.bridges.iter() {
                if bridge.name() == "Accessibility" {
                    bridge.set_fg_hwnd(fg_hwnd_raw);
                }
            }
            for (i, bridge) in self.bridges.iter().enumerate() {
                if bridge.name() == "Accessibility" {
                    if let Some(ctx) = bridge.read_context() {
                        if !ctx.word.is_empty() || !ctx.sentence.is_empty() {
                            if self.active_idx != i {
                                log!("Bridge switch: {} → Accessibility", self.bridges[self.active_idx].name());
                                self.bridge_switched = true;
                            }
                            self.active_idx = i;
                            self.last_user_pid = fg_pid;
                            self.last_context = Some(ctx.clone());
                            return Some(ctx);
                        }
                    }
                    break;
                }
            }
        }

        self.last_context.clone()
    }

    fn active_bridge(&self) -> Option<&dyn TextBridge> {
        self.bridges.get(self.active_idx).map(|b| b.as_ref())
    }

    fn active_bridge_name(&self) -> &str {
        self.effective_bridge().map(|b| b.name()).unwrap_or("none")
    }

    #[allow(dead_code)]
    fn replace_word(&self, new_text: &str) -> bool {
        let bridge_name = self.effective_bridge().map(|b| b.name()).unwrap_or("none");
        log!("replace_word('{}') via bridge '{}' (idx={})", new_text, bridge_name, self.active_idx);
        let result = self.effective_bridge().map(|b| b.replace_word(new_text)).unwrap_or(false);
        log!("replace_word result: {}", result);
        result
    }

    fn effective_bridge(&self) -> Option<&dyn TextBridge> {
        if self.last_user_was_browser {
            self.bridges.iter().find(|b| b.name() == "Browser").map(|b| b.as_ref())
        } else {
            self.bridges.get(self.active_idx).map(|b| b.as_ref())
        }
    }

    fn find_and_replace(&self, find: &str, replace: &str) -> bool {
        self.effective_bridge().map(|b| b.find_and_replace(find, replace)).unwrap_or(false)
    }

    fn find_and_replace_in_context(&self, find: &str, replace: &str, context: &str) -> bool {
        self.effective_bridge().map(|b| b.find_and_replace_in_context(find, replace, context)).unwrap_or(false)
    }

    fn find_and_replace_in_context_at(&self, find: &str, replace: &str, context: &str, char_offset: usize) -> bool {
        self.effective_bridge().map(|b| b.find_and_replace_in_context_at(find, replace, context, char_offset)).unwrap_or(false)
    }

    fn read_document_context(&self) -> Option<String> {
        self.effective_bridge().and_then(|b| b.read_document_context())
    }

    fn read_full_document(&self) -> Option<String> {
        // When last user app was a browser, ONLY read from Browser bridge
        if self.last_user_was_browser {
            if let Some(browser_idx) = self.bridges.iter().position(|b| b.name() == "Browser") {
                return self.bridges[browser_idx].read_full_document();
            }
            return None;
        }
        self.effective_bridge().and_then(|b| b.read_full_document())
    }

    fn select_range(&self, char_start: usize, char_end: usize) -> Option<(i32, i32)> {
        // Try all bridges — goto must work even when our window is foreground
        for bridge in &self.bridges {
            if let Some(pos) = bridge.select_range(char_start, char_end) {
                return Some(pos);
            }
        }
        None
    }

    fn set_target_hwnd(&self, hwnd: isize) {
        for bridge in &self.bridges {
            bridge.set_target_hwnd(hwnd);
        }
    }

    fn mark_error_underline(&self, char_start: usize, char_end: usize) -> bool {
        self.effective_bridge().map(|b| b.mark_error_underline(char_start, char_end)).unwrap_or(false)
    }

    fn clear_error_underline(&self, char_start: usize, char_end: usize) -> bool {
        self.effective_bridge().map(|b| b.clear_error_underline(char_start, char_end)).unwrap_or(false)
    }

    fn clear_all_error_underlines(&self) -> bool {
        self.effective_bridge().map(|b| b.clear_all_error_underlines()).unwrap_or(false)
    }

    fn should_skip_word_spelling(&self, cursor_off: usize, word_start: usize, word_end: usize, doc_char_len: usize, word_at_cursor: &str) -> bool {
        self.effective_bridge().map(|b| b.should_skip_word_spelling(cursor_off, word_start, word_end, doc_char_len, word_at_cursor)).unwrap_or(false)
    }

    fn should_skip_sentence_grammar(&self, cursor_off: usize, sent_start: usize, sent_end: usize, ends_with_punct: bool, doc_char_len: usize, word_at_cursor: &str) -> bool {
        self.effective_bridge().map(|b| b.should_skip_sentence_grammar(cursor_off, sent_start, sent_end, ends_with_punct, doc_char_len, word_at_cursor)).unwrap_or(false)
    }
}

// --- Detect if cursor is mid-word or at a word boundary ---

fn is_mid_word(word: &str) -> bool {
    if word.is_empty() {
        return false;
    }
    let last = word.chars().last().unwrap();
    last.is_alphanumeric() || last == '-' || last == '\''
}

/// Extract the prefix being typed (partial word for completion).
fn extract_prefix(word: &str) -> &str {
    word.trim()
}

fn escape_json_str(s: &str) -> String {
    s.replace('\\', "\\\\").replace('"', "\\\"").replace('\n', "\\n").replace('\r', "\\r").replace('\t', "\\t")
}

// --- Pending BERT state types ---

/// Pending spelling BERT re-ranking
struct PendingSpellingBert {
    request_id: bert_worker::RequestId,
    error_idx_word: String,       // word to match in writing_errors
    error_doc_offset: usize,
    candidates: Vec<(String, f32)>, // (candidate, ortho_sim)
}

/// Pending grammar correction BERT ranking
struct PendingGrammarBert {
    request_id: bert_worker::RequestId,
    sentence_context: String,
    doc_offset: usize,
    candidates: Vec<(String, String, String)>, // (corrected_sentence, explanation, rule_name)
}

/// Pending consonant confusion BERT scoring
struct PendingConsonantBert {
    request_id: bert_worker::RequestId,
    word: String,
    variants: Vec<String>,
    sentence_ctx: String,
    doc_offset: usize,
    // sentences[0] = original, sentences[1..] = variants
}

// --- egui app ---

/// Items delivered by background startup threads
enum StartupItem {
    Completer {
        model: Option<Model>,
        prefix_index: Option<PrefixIndex>,
        baselines: Option<Arc<Baselines>>,
        wordfreq: Option<Arc<HashMap<String, u64>>>,
        embedding_store: Option<Arc<EmbeddingStore>>,
        errors: Vec<String>,
    },
}

/// Lazy-loaded STT engine items
enum WhisperLoadItem {
    Final(Result<Box<dyn stt::SttEngine>, String>),
    Streaming(Result<Box<dyn stt::SttEngine>, String>),
}

struct ContextApp {
    manager: BridgeManager,
    context: CursorContext,
    last_poll: Instant,
    poll_interval: Duration,
    follow_cursor: bool,
    last_caret_pos: Option<(i32, i32)>,
    /// After goto, freeze window position for a few seconds so it doesn't jump back
    goto_freeze_until: Option<Instant>,
    // Grammar checker (kept for main-thread dictionary lookups; SWI grammar ops go through actor)
    checker: Option<AnyChecker>,
    /// Direct analyzer reference for dictionary lookups (cloned from checker before actor takes it)
    analyzer: Option<std::sync::Arc<mtag::Analyzer>>,
    /// Grammar actor: runs grammar checking on background thread
    grammar_actor: Option<grammar_actor::GrammarActorHandle>,
    grammar_errors: Vec<GrammarError>,
    last_checked_sentence: String,
    // Word completer — BERT model lives in dedicated worker thread (no lock contention)
    bert_worker: Option<bert_worker::BertWorkerHandle>,
    bert_ready: bool,
    completion_cancel: Arc<std::sync::atomic::AtomicBool>,
    /// Last time context changed — for debouncing completion dispatch
    last_context_change: Instant,
    /// The cache key we last dispatched (avoid re-dispatching same)
    dispatched_key: String,
    prefix_index: Option<PrefixIndex>,
    baselines: Option<Arc<Baselines>>,
    wordfreq: Option<Arc<HashMap<String, u64>>>,
    embedding_store: Option<Arc<EmbeddingStore>>,
    completions: Vec<Completion>,
    /// Open suggestions (any word) for fill-in-the-blank mode
    open_completions: Vec<Completion>,
    last_completed_prefix: String,
    /// Keep window large briefly after replacement so it doesn't shrink instantly
    last_replace_time: Instant,
    /// Cache: (masked_sentence, logits) from single_forward — reused when only prefix changes
    cached_forward: Option<(String, Vec<f32>)>,
    /// Cache: (masked_sentence, right_column) — right column only depends on logits, not prefix
    cached_right_column: Option<(String, Vec<(String, f32)>)>,
    /// Cache: (masked_sentence, scored_words) — mtag supplement BERT-ranked results
    cached_mtag_supplement: Option<(String, Vec<(String, f32)>)>,
    // Embedding sync
    last_embedding_sync: Instant,
    embedding_sync_interval: Duration,
    // Settings
    grammar_completion: bool,
    quality: u8, // 0=fast, 1=balanced, 2=full
    // Debounce: wait before running completion
    last_prefix_change: Instant,
    debounce_ms: u64,
    pending_completion: bool,
    // Completion selection mode (Ctrl+Space to enter, arrows to navigate, Enter to accept)
    selected_completion: Option<usize>,
    selection_mode: bool,
    /// Target app's HWND to return focus to (Word, Notepad, etc.)
    app_handle: Option<isize>,
    /// Platform abstraction for OS-specific services
    platform: Box<dyn platform::PlatformServices>,
    /// Track Ctrl+Space held to prevent repeated activation
    ctrl_space_held: bool,
    /// Which column is selected: 0=left (completions), 1=right (open_completions)
    selected_column: u8,
    // Status
    load_errors: Vec<String>,
    // Tab navigation
    selected_tab: usize, // 0=Innhold, 1=Grammatikk, 2=Innstillinger, 3=Debug
    show_debug_tab: bool,
    /// Error index to scroll to when cursor clicks on an underlined word
    focused_error_idx: Option<usize>,
    focused_error_set_time: Instant,
    focused_error_scroll_done: bool,
    /// Error index pinned to top of list (persists until explicitly cleared)
    // Error list (spelling + grammar)
    writing_errors: Vec<WritingError>,
    /// Words the user has chosen to ignore (spelling)
    ignored_words: std::collections::HashSet<String>,
    /// Last word that was spell-checked (to avoid re-checking)
    last_spell_checked_word: String,
    /// Authoritative document text from paragraph reads — used for grammar/spelling/pruning
    last_doc_text: String,
    /// Approximate doc length from masked_sentence — used only for change detection
    last_doc_approx_len: usize,
    /// Word just replaced — reject doc text updates that still contain this word (stale)
    last_replaced_word: Option<String>,
    /// Hash of last document text — skip entire update if doc unchanged
    last_doc_hash: u64,
    /// Number of sentences in last scan — detect paste vs fix
    last_sentence_count: usize,
    /// Hashes of sentences already checked for Prolog sub-splitting (expensive, persists across doc changes)
    prolog_checked_hashes: std::collections::HashSet<u64>,
    /// Hashes of sentences grammar-checked and found clean (no errors)
    processed_sentence_hashes: std::collections::HashSet<u64>,
    /// Track which sentence hashes belong to which paragraph (for cleanup on paragraph change/delete)
    paragraph_sentence_hashes: HashMap<String, Vec<u64>>,
    /// Pending spelling work — checked incrementally
    spelling_queue: Vec<SpellingQueueItem>,
    /// Pending grammar work: sentences still to check (incremental, one per frame)
    grammar_queue: Vec<(String, usize)>,
    /// Total sentences when grammar scan started (for progress bar)
    grammar_queue_total: usize,
    /// Whether a grammar scan is in progress (shows indicator in UI)
    grammar_scanning: bool,
    /// Cooldown after a fix — don't prune errors until canvas has repainted with new text
    /// Recently fixed words — don't re-detect these as errors (cleared when fresh text arrives from extension)
    /// Deferred find-and-replace (word, replacement, optional sentence context, doc char offset) — executed next frame
    pending_fix: Option<(String, String, String, usize)>,
    /// Pending consonant confusion candidates — validated with grammar checker after check_spelling
    pending_consonant_checks: Vec<WritingError>,
    /// Pending async BERT scoring for spelling re-ranking
    pending_spelling_bert: Vec<PendingSpellingBert>,
    /// Pending async BERT scoring for grammar correction ranking
    pending_grammar_bert: Vec<PendingGrammarBert>,
    /// Pending async BERT scoring for consonant confusion
    pending_consonant_bert: Vec<PendingConsonantBert>,
    /// Suggestion window: (misspelled_word, candidates)
    suggestion_window: Option<(String, Vec<(String, f32)>)>,
    suggestion_selection: std::sync::Arc<std::sync::Mutex<Option<usize>>>,
    /// Rule info popup: (rule_name, explanation, sentence_context, fix_idx, suggestion)
    rule_info_window: Option<(String, String, String, usize, String)>,
    // OCR clipboard monitoring
    ocr: Option<ocr::OcrClipboard>,
    ocr_receiver: Option<std::sync::mpsc::Receiver<Result<String, String>>>,
    ocr_text: Option<String>,
    ocr_copy_mode: bool, // true = copy to clipboard, false = speak
    // Microphone / Whisper
    whisper_engine: Option<Arc<Mutex<Box<dyn stt::SttEngine>>>>,       // final model (medium-q5 or tiny)
    whisper_streaming: Option<Arc<Mutex<Box<dyn stt::SttEngine>>>>,    // streaming model (base; None in tiny mode)
    mic_handle: Option<stt::MicHandle>,
    mic_transcribing: bool,
    mic_result_text: Option<String>,
    /// 0 = Rask (tiny only, ~75MB), 1 = Beste (base streaming + medium-q5 final, ~690MB)
    whisper_mode: u8,
    /// Receiver for lazy-loaded whisper engines
    whisper_load_rx: Option<std::sync::mpsc::Receiver<WhisperLoadItem>>,
    /// True while whisper models are being loaded
    whisper_loading: bool,
    /// Status message during whisper model loading
    whisper_load_status: String,
    /// Start recording as soon as whisper finishes loading
    whisper_pending_record: bool,
    // Voice selection window
    show_voice_window: bool,
    ui_scale: f32,
    voice_list: Vec<tts::VoiceInfo>,
    // Startup loading
    startup_rx: Option<std::sync::mpsc::Receiver<StartupItem>>,
    startup_done: Vec<String>,    // labels of completed items
    startup_total: usize,         // total items to load
}

/// Build left completions via BPE extension (when prefix_index has matches).
/// Runs on background thread — no access to self.
fn build_bpe_completions(
    model: &mut Model,
    masked: &str,
    prefix_lower: &str,
    matches: &[(u32, String)],
    logits: &[f32],
    wordfreq: Option<&HashMap<String, u64>>,
    nearby_words: &std::collections::HashSet<String>,
    mtag_valid: &std::collections::HashSet<String>,
    capitalize: bool,
    cancel: &std::sync::atomic::AtomicBool,
) -> Vec<Completion> {
    use std::sync::atomic::Ordering;

    let is_valid = |w: &str| -> bool {
        let key = w.to_lowercase();
        if nearby_words.contains(&key) { return false; }
        wordfreq.map_or(true, |wf| wf.contains_key(&key))
    };
    let cap = |s: &str| -> String {
        let mut c = s.chars();
        match c.next() {
            None => String::new(),
            Some(f) => f.to_uppercase().to_string() + c.as_str(),
        }
    };

    let mut token_scored: Vec<(String, f32)> = matches.iter()
        .map(|(tid, word)| (word.clone(), logits[*tid as usize]))
        .collect();
    token_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

    struct Candidate {
        token_ids: Vec<u32>,
        word: String,
        score: f32,
        done: bool,
    }
    let mut candidate_set: std::collections::HashSet<String> = std::collections::HashSet::new();
    let mut candidates: Vec<Candidate> = Vec::new();

    // Top 20 by logit
    for (tok_word, tok_score) in token_scored.iter().take(20) {
        if candidate_set.insert(tok_word.clone()) {
            if let Some((tid, _)) = matches.iter().find(|(_, w)| w == tok_word) {
                candidates.push(Candidate {
                    token_ids: vec![*tid],
                    word: tok_word.clone(),
                    score: *tok_score,
                    done: false,
                });
            }
        }
    }
    // Top 20 long tokens (≥5 chars)
    let mut long_tokens: Vec<&(String, f32)> = token_scored.iter()
        .filter(|(w, s)| w.len() >= 5 && *s > 0.0)
        .collect();
    long_tokens.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
    for (tok_word, tok_score) in long_tokens.iter().take(20) {
        if candidate_set.insert(tok_word.clone()) {
            if let Some((tid, _)) = matches.iter().find(|(_, w)| w == tok_word) {
                candidates.push(Candidate {
                    token_ids: vec![*tid],
                    word: tok_word.clone(),
                    score: *tok_score,
                    done: false,
                });
            }
        }
    }

    let mask_parts: Vec<&str> = masked.splitn(2, "<mask>").collect();
    let ctx_before = mask_parts[0].trim_end();
    let ctx_after = mask_parts.get(1).map(|s| s.trim_start()).unwrap_or(".");

    let max_steps = if prefix_lower.len() <= 3 { 3 } else { 1 };
    for _step in 0..max_steps {
        if cancel.load(Ordering::Acquire) { return vec![]; }
        let best_score = candidates.iter()
            .filter(|c| !c.done)
            .map(|c| c.score)
            .fold(f32::NEG_INFINITY, f32::max);
        let threshold = best_score - 15.0;
        let mut to_extend: Vec<usize> = candidates.iter().enumerate()
            .filter(|(_, c)| !c.done && c.score >= threshold)
            .map(|(i, _)| i)
            .collect();
        for c in candidates.iter_mut() {
            if !c.done && c.score < threshold { c.done = true; }
        }
        let batch_cap = if prefix_lower.len() <= 2 { 5 } else { 10 };
        to_extend.truncate(batch_cap);
        if to_extend.is_empty() { break; }

        let batch_texts: Vec<String> = to_extend.iter()
            .map(|&i| {
                let accumulated = model.tokenizer
                    .decode(&candidates[i].token_ids, false)
                    .unwrap_or_default();
                let accumulated = accumulated.trim();
                format!("{} {}<mask> {}", ctx_before, accumulated, ctx_after)
            })
            .collect();

        match model.batched_forward_argmax(&batch_texts) {
            Ok((argmaxes, _)) => {
                for (k, &i) in to_extend.iter().enumerate() {
                    let best_id = argmaxes[k];
                    let best_token = &model.id_to_token[best_id as usize];
                    let is_continuation = !best_token.starts_with('Ġ')
                        && !matches!(best_token.as_str(), "." | "," | "!" | "?" | ";" | ":");
                    if is_continuation {
                        candidates[i].token_ids.push(best_id);
                        candidates[i].word = model.tokenizer
                            .decode(&candidates[i].token_ids, false)
                            .unwrap_or_default().trim().to_string();
                    } else {
                        candidates[i].done = true;
                    }
                }
            }
            Err(_) => break,
        }
    }

    let mut left_scored: Vec<(String, f32)> = Vec::new();
    let mut seen_words: std::collections::HashSet<String> = std::collections::HashSet::new();
    for c in &candidates {
        let key = c.word.to_lowercase();
        if is_valid(&key) && seen_words.insert(key.clone()) {
            left_scored.push((key, c.score));
        }
    }
    eprintln!("BPE candidates: [{}]",
        candidates.iter().take(10).map(|c| format!("{}({:.1},done={})", c.word, c.score, c.done))
        .collect::<Vec<_>>().join(", "));

    left_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
    left_scored.into_iter()
        .take(25)
        .map(|(w, s)| Completion {
            word: if capitalize { cap(&w) } else { w },
            score: s,
            elapsed_ms: 0.0,
        })
        .collect()
}

/// Build left completions from pre-fetched mtag candidates scored by BERT.
/// Runs on background thread — no access to self.
fn build_mtag_completions(
    model: &mut Model,
    masked: &str,
    mtag_candidates: &[String],
    logits: &[f32],
    capitalize: bool,
    cancel: &std::sync::atomic::AtomicBool,
) -> Vec<Completion> {
    use std::sync::atomic::Ordering;
    if mtag_candidates.is_empty() { return vec![]; }

    let cap = |s: &str| -> String {
        let mut c = s.chars();
        match c.next() {
            None => String::new(),
            Some(f) => f.to_uppercase().to_string() + c.as_str(),
        }
    };

    let mask_parts: Vec<&str> = masked.splitn(2, "<mask>").collect();
    let ctx_before = mask_parts[0].trim_end();
    let ctx_after = mask_parts.get(1).map(|s| s.trim_start()).unwrap_or(".");

    let candidates_with_tokens: Vec<(String, Vec<u32>)> = mtag_candidates.iter()
        .filter_map(|w| {
            let enc = model.tokenizer.encode(format!(" {}", w).as_str(), false).ok()?;
            let ids: Vec<u32> = enc.get_ids().to_vec();
            if ids.is_empty() { return None; }
            Some((w.clone(), ids))
        })
        .collect();

    // First-token score
    let mut scores: Vec<f32> = candidates_with_tokens.iter()
        .map(|(_, ids)| logits[ids[0] as usize])
        .collect();

    // Multi-token scoring
    let max_tokens = candidates_with_tokens.iter().map(|(_, ids)| ids.len()).max().unwrap_or(1);
    for t in 1..max_tokens {
        if cancel.load(Ordering::Acquire) { return vec![]; }
        let to_score: Vec<usize> = candidates_with_tokens.iter().enumerate()
            .filter(|(_, (_, ids))| ids.len() > t)
            .map(|(i, _)| i)
            .collect();
        if to_score.is_empty() { break; }

        let mut unique_prefixes: Vec<Vec<u32>> = Vec::new();
        let mut prefix_to_idx: std::collections::HashMap<Vec<u32>, usize> = std::collections::HashMap::new();
        let mut candidate_to_prefix: Vec<usize> = Vec::new();
        for &i in &to_score {
            let token_prefix = candidates_with_tokens[i].1[..t].to_vec();
            let pidx = if let Some(&existing) = prefix_to_idx.get(&token_prefix) {
                existing
            } else {
                let idx = unique_prefixes.len();
                prefix_to_idx.insert(token_prefix.clone(), idx);
                unique_prefixes.push(token_prefix);
                idx
            };
            candidate_to_prefix.push(pidx);
        }

        let batch_texts: Vec<String> = unique_prefixes.iter()
            .map(|ids| {
                let partial = model.tokenizer.decode(ids, false).unwrap_or_default();
                format!("{} {}<mask> {}", ctx_before, partial.trim(), ctx_after)
            })
            .collect();

        if let Ok((batch_logits, _)) = model.batched_forward(&batch_texts) {
            for (k, &i) in to_score.iter().enumerate() {
                let pidx = candidate_to_prefix[k];
                scores[i] += batch_logits[pidx][candidates_with_tokens[i].1[t] as usize];
            }
        }
    }

    let mut scored: Vec<(String, f32)> = candidates_with_tokens.iter().enumerate()
        .map(|(i, (w, ids))| (w.clone(), scores[i] / ids.len() as f32))
        .collect();
    scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
    scored.into_iter()
        .take(25)
        .map(|(w, s)| Completion {
            word: if capitalize { cap(&w) } else { w },
            score: s,
            elapsed_ms: 0.0,
        })
        .collect()
}

/// Build right-column completions from logits (no model call needed).
fn build_right_completions(
    model: &Model,
    logits: &[f32],
    wordfreq: Option<&HashMap<String, u64>>,
    nearby_words: &std::collections::HashSet<String>,
    left_words: &std::collections::HashSet<String>,
) -> Vec<Completion> {
    let is_valid = |w: &str| -> bool {
        let key = w.to_lowercase();
        if nearby_words.contains(&key) { return false; }
        wordfreq.map_or(true, |wf| wf.contains_key(&key))
    };

    let mut seen: std::collections::HashSet<String> = std::collections::HashSet::new();
    let mut all_scored: Vec<(String, f32)> = model.id_to_token.iter()
        .enumerate()
        .filter(|(_, tok)| tok.starts_with('Ġ'))
        .filter_map(|(i, _)| {
            let decoded = model.tokenizer
                .decode(&[i as u32], false)
                .unwrap_or_default().trim().to_lowercase();
            if decoded.is_empty() || decoded.len() <= 1 { return None; }
            if !is_valid(&decoded) || left_words.contains(&decoded) { return None; }
            if !seen.insert(decoded.clone()) { return None; }
            Some((decoded, logits[i]))
        })
        .collect();
    all_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
    all_scored.into_iter()
        .take(10)
        .map(|(w, s)| Completion { word: w, score: s, elapsed_ms: 0.0 })
        .collect()
}

impl ContextApp {
    fn new(grammar_completion: bool, use_swipl: bool, quality: u8, show_debug_tab: bool) -> Self {
        let platform = platform::create_platform();
        platform.init_runtime();

        let mut load_errors = Vec::new();

        // Load dictionary (analyzer) on main thread for fast lookups.
        // SWI-Prolog grammar checker loads on the grammar actor thread (later).
        let analyzer: Option<std::sync::Arc<mtag::Analyzer>> = match mtag::Analyzer::new(&dict_path()) {
            Ok(a) => {
                eprintln!("Loaded dictionary with {} entries", a.dict_size());
                Some(std::sync::Arc::new(a))
            }
            Err(e) => {
                let msg = format!("Dictionary: {}", e);
                eprintln!("{}", msg);
                load_errors.push(msg);
                None
            }
        };

        // Spawn heavy model loading on background threads
        let (startup_tx, startup_rx) = std::sync::mpsc::channel();

        // Thread 1: NorBERT4 + completer
        let tx2 = startup_tx.clone();
        std::thread::spawn(move || {
            let data = data_dir();
            let onnx_path = data.join("onnx/norbert4_base_int8.onnx");
            let tokenizer_path = data.join("onnx/tokenizer.json");
            let wordfreq_path = data.join("wordfreq.tsv");
            let minilm_onnx = data.join("minilm-onnx/model_optimized.onnx");
            let minilm_tok = data.join("minilm-onnx/tokenizer.json");
            let embed_cache = data.join("word_embeddings.bin");

            let mut errors = Vec::new();
            let (model_opt, prefix_index, baselines, wf, embedding_store) =
                match ContextApp::load_completer(
                    &onnx_path, &tokenizer_path, &wordfreq_path,
                    &minilm_onnx, &minilm_tok, &embed_cache,
                ) {
                    Ok(parts) => parts,
                    Err(e) => {
                        let msg = format!("Completer: {}", e);
                        eprintln!("{}", msg);
                        errors.push(msg);
                        (None, None, None, None, None)
                    }
                };
            let _ = tx2.send(StartupItem::Completer {
                model: model_opt, prefix_index,
                baselines: baselines.map(|b| Arc::new(b)),
                wordfreq: wf,
                embedding_store: embedding_store.map(|e| Arc::new(e)),
                errors,
            });
        });

        // Whisper models are lazy-loaded on first mic press (saves ~650MB+ RAM)
        drop(startup_tx);

        ContextApp {
            manager: BridgeManager::new(platform::create_platform()),
            context: CursorContext::default(),
            last_poll: Instant::now(),
            poll_interval: Duration::from_millis(300),
            follow_cursor: true,
            goto_freeze_until: None,
            last_caret_pos: None,
            checker: None,
            analyzer,
            grammar_actor: None,
            grammar_errors: Vec::new(),
            last_checked_sentence: String::new(),
            bert_worker: None,
            bert_ready: false,
            completion_cancel: Arc::new(std::sync::atomic::AtomicBool::new(false)),
            last_context_change: Instant::now(),
            dispatched_key: String::new(),
            prefix_index: None,
            baselines: None,
            wordfreq: None,
            embedding_store: None,
            completions: Vec::new(),
            open_completions: Vec::new(),
            last_completed_prefix: String::new(),
            last_replace_time: Instant::now() - Duration::from_secs(10),
            cached_forward: None,
            cached_right_column: None,
            cached_mtag_supplement: None,
            last_embedding_sync: Instant::now(),
            embedding_sync_interval: Duration::from_secs(3),
            grammar_completion,
            quality,
            last_prefix_change: Instant::now(),
            debounce_ms: if quality == 0 { 100 } else { 150 },
            pending_completion: false,
            selected_completion: None,
            selection_mode: false,
            app_handle: None,
            platform,
            ctrl_space_held: false,
            selected_column: 0,
            load_errors,
            selected_tab: 0,
            show_debug_tab,
            focused_error_idx: None,
            focused_error_set_time: Instant::now() - Duration::from_secs(10),
            focused_error_scroll_done: false,
            writing_errors: Vec::new(),
            ignored_words: std::collections::HashSet::new(),
            last_spell_checked_word: String::new(),
            last_doc_text: String::new(),
            last_doc_approx_len: 0,
            last_replaced_word: None,
            last_doc_hash: 0,
            last_sentence_count: 0,
            prolog_checked_hashes: std::collections::HashSet::new(),
            processed_sentence_hashes: std::collections::HashSet::new(),
            paragraph_sentence_hashes: HashMap::new(),
            spelling_queue: Vec::new(),
            grammar_queue: Vec::new(),
            grammar_queue_total: 0,
            grammar_scanning: false,
            pending_fix: None,
            pending_consonant_checks: Vec::new(),
            pending_spelling_bert: Vec::new(),
            pending_grammar_bert: Vec::new(),
            pending_consonant_bert: Vec::new(),
            suggestion_window: None,
            suggestion_selection: std::sync::Arc::new(std::sync::Mutex::new(None)),
            rule_info_window: None,
            ocr: match ocr::OcrClipboard::new() {
                Ok(o) => { eprintln!("OCR clipboard monitor ready"); Some(o) }
                Err(e) => { eprintln!("OCR not available: {}", e); None }
            },
            ocr_receiver: None,
            ocr_text: None,
            ocr_copy_mode: false,
            whisper_engine: None,
            whisper_streaming: None,
            mic_handle: None,
            mic_transcribing: false,
            mic_result_text: None,
            whisper_mode: 1, // default: Beste (base+medium-q5)
            whisper_load_rx: None,
            whisper_loading: false,
            whisper_load_status: String::new(),
            whisper_pending_record: false,
            show_voice_window: false,
            ui_scale: 1.0,
            voice_list: Vec::new(),
            startup_rx: Some(startup_rx),
            startup_done: Vec::new(),
            startup_total: 1, // completer only
        }
    }

    fn load_checker() -> Result<GrammarChecker, Box<dyn std::error::Error>> {
        let compound_data = std::fs::read_to_string(compound_data_path())
            .unwrap_or_else(|_| {
                eprintln!("compound_data.pl not found, using empty");
                String::new()
            });
        GrammarChecker::new(dict_path().to_str().unwrap(), &compound_data)
    }

    fn load_swipl_checker(swipl_path: &str) -> Result<SwiGrammarChecker, Box<dyn std::error::Error>> {
        SwiGrammarChecker::new(
            swipl_path,
            dict_path().to_str().unwrap(),
            grammar_rules_path().to_str().unwrap(),
            syntaxer_dir().to_str().unwrap(),
        )
    }

    fn load_completer(
        onnx_path: &PathBuf, tokenizer_path: &PathBuf, wordfreq_path: &PathBuf,
        minilm_onnx: &PathBuf, minilm_tok: &PathBuf, embed_cache: &PathBuf,
    ) -> anyhow::Result<(
        Option<Model>,
        Option<PrefixIndex>,
        Option<Baselines>,
        Option<Arc<HashMap<String, u64>>>,
        Option<EmbeddingStore>,
    )> {
        eprintln!("Loading NorBERT4 from {}...", onnx_path.display());
        let mut model = Model::load(
            onnx_path.to_str().unwrap(),
            tokenizer_path.to_str().unwrap(),
        )?;
        eprintln!("NorBERT4 loaded. Vocab: {}, backend: {}", model.vocab_size(), model.backend_name());

        eprintln!("Building prefix index...");
        let pi = prefix_index::build_prefix_index(&model.tokenizer);
        eprintln!("Prefix index: {} prefixes", pi.len());

        eprintln!("Computing baselines...");
        let baselines = compute_baseline(&mut model)?;

        let wf = wordfreq::load_wordfreq(wordfreq_path.as_path(), 10);
        eprintln!("WordFreq: {} words", wf.len());

        // MiniLM embedding store disabled — saves ~500 MB RAM.
        // PMI topic words (via NorBERT4 baselines) still active.
        let embedding_store: Option<EmbeddingStore> = None;

        Ok((Some(model), Some(pi), Some(baselines), Some(Arc::new(wf)), embedding_store))
    }

    fn run_grammar_check(&mut self) {
        let sentence = self.context.sentence.trim().to_string();
        if sentence.is_empty() || sentence == self.last_checked_sentence {
            return;
        }
        self.last_checked_sentence = sentence.clone();

        // Grammar checking happens via the actor/queue path, not this synchronous path.
        // (checker has been moved to the grammar actor)
    }

    /// Check spelling of a word. `sentence_ctx` is the sentence it appears in.
    fn check_spelling(&mut self, word: &str, sentence_ctx: &str, paragraph_id: &str, doc_offset: usize) {
        let clean = word.trim().to_lowercase();
        if clean.is_empty() || clean.len() < 2 || clean == self.last_spell_checked_word {
            return;
        }
        self.last_spell_checked_word = clean.clone();

        // Skip if word is in ignore list
        if self.ignored_words.contains(&clean) {
            return;
        }

        // Skip punctuation-only, numbers, or words containing digits
        if clean.chars().all(|c| c.is_ascii_punctuation() || c.is_ascii_digit()) {
            return;
        }
        if !clean.chars().any(|c| c.is_alphabetic()) {
            return;
        }
        if clean.chars().any(|c| c.is_ascii_digit()) {
            return;
        }
        // Skip email addresses
        if clean.contains('@') {
            return;
        }
        // Split on slash and check each part separately (oppdrag/prosjekt → oppdrag, prosjekt)
        if clean.contains('/') {
            for part in clean.split('/') {
                let part = part.trim();
                if !part.is_empty() && part.len() >= 2 {
                    self.check_spelling(part, sentence_ctx, paragraph_id, doc_offset);
                }
            }
            return;
        }

        // Skip capitalized words — likely proper nouns (Tekna, Oslo, etc.)
        if word.trim().chars().next().map_or(false, |c| c.is_uppercase()) {
            return;
        }
        // Skip words with apostrophe/symbol + Norwegian ending (api'er, pdf'en)
        {
            let separators = ['\'', '\u{00b4}', '\u{2019}', '\u{2018}', '\u{0060}', '-'];
            let split = clean.char_indices().find(|(_, c)| separators.contains(c));
            if let Some((byte_pos, sep_char)) = split {
                let before = &clean[..byte_pos];
                let after = &clean[byte_pos + sep_char.len_utf8()..];
                let endings = ["er", "en", "ene", "ens", "et", "ets", "a"];
                if before.len() >= 2 && endings.contains(&after) {
                    return;
                }
            }
        }

        // Common modal verb misspellings — BERT can't distinguish "vil" vs "ville" in context
        let modal_fixes: &[(&str, &str)] = &[
            ("vile", "ville"), ("skule", "skulle"), ("kune", "kunne"), ("måte", "måtte"),
            ("bure", "burde"), ("tore", "torde"), ("gide", "gidde"),
        ];
        if let Some((_, correct)) = modal_fixes.iter().find(|(wrong, _)| *wrong == clean) {
            if !self.writing_errors.iter().any(|e| e.word == clean && e.sentence_context == sentence_ctx && e.doc_offset == doc_offset && !e.ignored) {
                log!("modal fix: '{}' → '{}'", clean, correct);
                self.writing_errors.push(WritingError {
                    category: ErrorCategory::Spelling,
                    word: clean.clone(),
                    suggestion: correct.to_string(),
                    explanation: format!("«{}» → «{}»", clean, correct),
                    rule_name: "stavefeil_modal".to_string(),
                    sentence_context: sentence_ctx.to_string(),
                    doc_offset,
                    position: 0,
                    ignored: false,
                    word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: paragraph_id.to_string(),
                });
            }
            return;
        }

        // Phase 1: Dictionary lookups (immutable borrow on checker)
        let mut found;
        let kt_gt_valid_alt: Option<String>;
        let consonant_alts: Vec<String>;
        let original_found;

        {
            let analyzer = match &self.analyzer {
                Some(a) => a,
                None => return,
            };

            found = analyzer.has_word(&clean);
            if !found {
                // Check if it's a valid compound word (e.g. maskinlæringsalgoritmene)
                // Try splitting at every position and check if both parts are known
                for i in 3..clean.len().saturating_sub(2) {
                    if !clean.is_char_boundary(i) { continue; }
                    let left = &clean[..i];
                    let right = &clean[i..];
                    // Allow 's' binding letter: maskinlæring-s-algoritmene
                    if analyzer.has_word(left) && analyzer.has_word(right) {
                        found = true;
                        break;
                    }
                    if right.starts_with('s') && right.len() > 3 {
                        let right_after_s = &right[1..];
                        if analyzer.has_word(left) && analyzer.has_word(right_after_s) {
                            found = true;
                            break;
                        }
                    }
                }
                if !found {
                    log!("spell: '{}' NOT in dict (sentence: '{}')", clean, trunc(sentence_ctx, 50));
                }
            }

            // kt/gt confusion: check if alt form exists
            kt_gt_valid_alt = if found {
                let alt = if clean.ends_with("kt") {
                    Some(format!("{}gt", &clean[..clean.len()-2]))
                } else if clean.ends_with("gt") {
                    Some(format!("{}kt", &clean[..clean.len()-2]))
                } else {
                    None
                };
                alt.filter(|a| analyzer.has_word(a))
            } else {
                None
            };

            // Consonant confusion: find valid alternatives with shared POS
            consonant_alts = if found && clean.len() >= 4 {
                let orig_pos = {
                    let mut pos = std::collections::HashSet::new();
                    if let Some(readings) = analyzer.dict_lookup(&clean) {
                        for r in &readings { pos.insert(r.pos.to_string()); }
                    }
                    pos
                };
                consonant_variants(&clean).into_iter()
                    .filter(|v| {
                        if !analyzer.has_word(v) { return false; }
                        let v_pos = {
                            let mut pos = std::collections::HashSet::new();
                            if let Some(readings) = analyzer.dict_lookup(v) {
                                for r in &readings { pos.insert(r.pos.to_string()); }
                            }
                            pos
                        };
                        let shared = orig_pos.intersection(&v_pos).count() > 0;
                        if !shared {
                            log!("  consonant skip '{}' → '{}' (no shared POS: {:?} vs {:?})", clean, v, orig_pos, v_pos);
                        }
                        shared
                    })
                    .collect()
            } else {
                Vec::new()
            };

            original_found = if !found { analyzer.has_word(word.trim()) } else { false };
        } // analyzer borrow dropped

        // Phase 2: BERT scoring via worker (async) + writing errors
        if found {
            // kt/gt confusion — send to BERT worker for sentence scoring
            if let Some(ref alt) = kt_gt_valid_alt {
                let mut all_variants = vec![alt.clone()];
                // Merge consonant alts if any
                for a in &consonant_alts {
                    if !all_variants.contains(a) {
                        all_variants.push(a.clone());
                    }
                }
                self.send_consonant_bert(&clean, all_variants, sentence_ctx, doc_offset);
            } else if !consonant_alts.is_empty() {
                // Consonant confusion only (no kt/gt)
                self.send_consonant_bert(&clean, consonant_alts.clone(), sentence_ctx, doc_offset);
            }

            return;
        }

        if original_found {
            return;
        }

        // Word not found — unified suggestion pipeline
        // Dedup: skip if this word already has a spelling error in this paragraph
        if self.writing_errors.iter().any(|e| {
            matches!(e.category, ErrorCategory::Spelling) && e.word.to_lowercase() == clean && e.paragraph_id == paragraph_id && !e.ignored
        }) {
            return;
        }

        // BERT must be available — no fallback without it
        // Don't set last_spell_checked_word so we retry when BERT loads
        if !self.bert_ready {
            self.last_spell_checked_word.clear();
            return;
        }

        let suggestions = self.find_spelling_suggestions(&clean, sentence_ctx);
        // Fix up doc_offset on any pending BERT re-ranking request
        if let Some(pending) = self.pending_spelling_bert.last_mut() {
            if pending.error_doc_offset == 0 {
                pending.error_doc_offset = doc_offset;
            }
        }
        let best = suggestions.first().map(|(w, _)| w.clone()).unwrap_or_default();
        let rule = "stavefeil_bert";

        // Show best orthographic suggestion immediately; BERT will re-rank and update later.
        let has_pending_bert = self.pending_spelling_bert.last()
            .map(|p| p.error_idx_word == clean.to_lowercase())
            .unwrap_or(false);
        let shown_suggestion = best.clone();

        self.writing_errors.push(WritingError {
            category: ErrorCategory::Spelling,
            word: clean.clone(),
            suggestion: shown_suggestion,
            explanation: format!("«{}» finnes ikke i ordboken.", clean),
            rule_name: rule.to_string(),
            sentence_context: sentence_ctx.to_string(),
            doc_offset,
            position: 0,
            ignored: false,
            word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: paragraph_id.to_string(),
        });
        if !best.is_empty() {
            log!("Spelling: '{}' → '{}' (unified pipeline, bert_pending={})", clean, best, has_pending_bert);
        } else {
            log!("Spelling: '{}' not found, no valid suggestion", clean);
        }
    }

    /// Validate pending consonant confusion candidates with grammar checker + word frequency.
    /// Promotes to writing_errors if:
    /// 1. The variant sentence has fewer grammar errors (variant fixes something), OR
    /// 2. Grammar is equal for both BUT the variant is much more frequent (≥10x in wordfreq)
    ///    — catches rare/dialectal forms like "spile" when "spille" is the standard form.
    fn validate_consonant_checks(&mut self) {
        if self.pending_consonant_checks.is_empty() {
            return;
        }
        let pending = std::mem::take(&mut self.pending_consonant_checks);
        // checker has been moved to grammar actor; consonant validation
        // that requires check_sentence is skipped. Accept all BERT-preferred candidates.
        let budget_start = std::time::Instant::now();
        let mut deferred: Vec<WritingError> = Vec::new();
        for mut candidate in pending {
            // Time budget: defer remaining candidates to next frame
            if budget_start.elapsed().as_millis() > 15 {
                deferred.push(candidate);
                continue;
            }
            // Already flagged for this sentence occurrence?
            if self.writing_errors.iter().any(|e| e.sentence_context == candidate.sentence_context && e.doc_offset == candidate.doc_offset && !e.ignored) {
                continue;
            }
            // Extract orig_word and variant_word from rule_name "consonant_confusion:spile:spille"
            let parts: Vec<&str> = candidate.rule_name.splitn(3, ':').collect();
            let (orig_word, variant_word) = if parts.len() == 3 {
                (parts[1].to_string(), parts[2].to_string())
            } else {
                continue;
            };

            // Grammar check skipped (checker moved to actor). Accept BERT-preferred candidates.
            log!("consonant validate (no grammar): '{}' → '{}'", orig_word, variant_word);

            // Clean up rule_name for display
            candidate.rule_name = "consonant_confusion".to_string();

            // BERT already decided the variant is better — accept it
            log!("consonant confirmed: '{}' → '{}' (BERT preferred)", orig_word, variant_word);
            self.writing_errors.push(candidate);
        }
        // Re-queue deferred candidates for next frame
        if !deferred.is_empty() {
            self.pending_consonant_checks.extend(deferred);
        }
    }

    /// Send consonant confusion check to BERT worker for async sentence scoring.
    fn send_consonant_bert(&mut self, word: &str, variants: Vec<String>, sentence_ctx: &str, doc_offset: usize) {
        if !self.bert_ready { return; }
        // Skip if already flagged for this position
        if self.writing_errors.iter().any(|e| e.sentence_context == sentence_ctx && e.doc_offset == doc_offset && !e.ignored) {
            return;
        }
        let worker = match &mut self.bert_worker {
            Some(w) => w,
            None => return,
        };
        // Extract context_before/context_after around the word
        let word_lower = word.to_lowercase();
        let sentence_lower = sentence_ctx.to_lowercase();
        let word_pos = sentence_lower.find(&word_lower);
        let (context_before, context_after) = if let Some(pos) = word_pos {
            (sentence_ctx[..pos].trim_end().to_string(), sentence_ctx[pos + word_lower.len()..].trim_start().to_string())
        } else {
            (sentence_ctx.to_string(), String::new())
        };
        // Include original word + variants as candidates
        let mut candidates = vec![word.to_string()];
        candidates.extend(variants.iter().cloned());
        log!("consonant BERT send: '{}' variants={:?}", word, variants);
        let request_id = worker.send(|id| bert_worker::BertRequest::SpellingScore { id, context_before, context_after, candidates });
        self.pending_consonant_bert.push(PendingConsonantBert {
            request_id,
            word: word.to_string(),
            variants,
            sentence_ctx: sentence_ctx.to_string(),
            doc_offset,
        });
    }

    /// Poll BERT worker for all response types and handle them.
    fn poll_bert_responses(&mut self, ctx: &egui::Context) {
        // Collect all available responses first (avoids borrow conflicts)
        let mut responses: Vec<bert_worker::BertResponse> = Vec::new();
        if let Some(worker) = &mut self.bert_worker {
            while let Some(resp) = worker.try_recv() {
                responses.push(resp);
            }
        }

        for resp in responses {
            match resp {
                bert_worker::BertResponse::Completion { id: _, cache_key, left, right } => {
                    // Only accept results for the current context — ignore stale responses
                    let current_key = format!("{}|{}",
                        self.context.masked_sentence.as_deref().unwrap_or(""),
                        extract_prefix(&self.context.word));
                    if cache_key != current_key {
                        log!("Ignoring stale completion (key mismatch)");
                        continue;
                    }
                    log!("BERT completion received: {} left, {} right: [{}]", left.len(), right.len(),
                        left.iter().take(10).map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "));
                    // Apply grammar filter on main thread
                    if self.grammar_completion {
                        if let Some(analyzer) = &self.analyzer {
                            // Step 1: Dictionary-filter completions
                            let left_filtered: Vec<Completion> = left.into_iter()
                                .filter(|c| analyzer.has_word(&c.word.to_lowercase()))
                                .take(15)
                                .collect();
                            let right_filtered: Vec<Completion> = right.into_iter()
                                .filter(|c| analyzer.has_word(&c.word.to_lowercase()))
                                .take(15)
                                .collect();

                            // Step 2: Grammar-filter via actor (synchronous)
                            log!("Grammar filter: actor={} left={} right={}", self.grammar_actor.is_some(), left_filtered.len(), right_filtered.len());
                            if let Some(actor) = &self.grammar_actor {
                                let sentence = &self.context.sentence;
                                let prefix = extract_prefix(&self.context.word);
                                let ctx_for_grammar = sentence.strip_suffix(prefix)
                                    .unwrap_or(sentence).trim_end().to_string();

                                // Build all test sentences for batch grammar check
                                let last_fragment = {
                                    let start = ctx_for_grammar.rfind(|c: char| ".!?".contains(c))
                                        .map(|i| i + 1).unwrap_or(0);
                                    ctx_for_grammar[start..].trim().to_string()
                                };
                                let all_candidates: Vec<&Completion> = left_filtered.iter()
                                    .chain(right_filtered.iter()).collect();
                                let test_sentences: Vec<String> = all_candidates.iter()
                                    .map(|c| format!("{} {}.", last_fragment, c.word))
                                    .collect();

                                // One batch call to grammar actor
                                let t_grammar = std::time::Instant::now();
                                let batch_results = actor.check_sentences_batch(&test_sentences);
                                log!("Grammar batch: {} sentences in {:?}", test_sentences.len(), t_grammar.elapsed());

                                // Build check results map
                                let mut grammar_ok: std::collections::HashMap<String, bool> = std::collections::HashMap::new();
                                for (i, c) in all_candidates.iter().enumerate() {
                                    let errors = &batch_results[i];
                                    grammar_ok.insert(c.word.to_lowercase(), errors.is_empty());
                                }

                                // Filter left and right by grammar results
                                self.completions = left_filtered.into_iter()
                                    .filter(|c| *grammar_ok.get(&c.word.to_lowercase()).unwrap_or(&false))
                                    .take(5).collect();
                                self.open_completions = right_filtered.into_iter()
                                    .filter(|c| *grammar_ok.get(&c.word.to_lowercase()).unwrap_or(&false))
                                    .take(5).collect();
                            } else {
                                self.completions = left_filtered.into_iter().take(5).collect();
                                self.open_completions = right_filtered.into_iter().take(5).collect();
                            }

                            eprintln!("completions (grammar-filtered): [{}]",
                                self.completions.iter().map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "));
                            eprintln!("open_completions (grammar-filtered): [{}]",
                                self.open_completions.iter().map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "));
                        } else {
                            self.completions = left.into_iter().take(5).collect();
                            self.open_completions = right.into_iter().take(5).collect();
                        }
                    } else {
                        self.completions = left.into_iter().take(5).collect();
                        self.open_completions = right.into_iter().take(5).collect();
                    }
                    self.last_completed_prefix = cache_key;
                }

                bert_worker::BertResponse::SpellingScore { id, scored_candidates } => {
                    if let Some(idx) = self.pending_spelling_bert.iter().position(|p| p.request_id == id) {
                        let pending = self.pending_spelling_bert.remove(idx);
                        self.handle_spelling_bert_response(pending, &scored_candidates);
                    } else if let Some(idx) = self.pending_grammar_bert.iter().position(|p| p.request_id == id) {
                        let pending = self.pending_grammar_bert.remove(idx);
                        self.handle_grammar_bert_response(pending, &scored_candidates);
                    } else if let Some(idx) = self.pending_consonant_bert.iter().position(|p| p.request_id == id) {
                        let pending = self.pending_consonant_bert.remove(idx);
                        self.handle_consonant_bert_response(pending, &scored_candidates);
                    }
                }

                bert_worker::BertResponse::MlmForward { .. } => {
                    // MLM results currently unused in async flow
                }
            }
        }
        // Request repaint so new results are rendered.
        // Always schedule periodic repaints — egui may sleep when our window
        // is not focused (e.g. user is typing in Word), but we still need to
        // poll for BERT responses and update the UI.
        if !self.pending_spelling_bert.is_empty()
            || !self.pending_grammar_bert.is_empty()
            || !self.pending_consonant_bert.is_empty()
            || !self.completions.is_empty()
            || !self.open_completions.is_empty()
        {
            ctx.request_repaint_after(Duration::from_millis(50));
        }
    }

    /// Handle BERT spelling score response for spelling re-ranking.
    fn handle_spelling_bert_response(&mut self, pending: PendingSpellingBert, scored_candidates: &[(String, f32)]) {
        if scored_candidates.is_empty() {
            log!("spelling BERT response: empty scored_candidates");
            return;
        }
        // scored_candidates is already sorted best-first by the worker
        // Re-rank using sqrt(ortho) weighting from the original candidates
        let ortho_map: std::collections::HashMap<String, f32> = pending.candidates.iter().cloned().collect();
        let mut rescored: Vec<(String, f32)> = scored_candidates.iter()
            .map(|(candidate, bert_score)| {
                let ortho_sim = ortho_map.get(candidate).copied().unwrap_or(0.5);
                let final_score = bert_score * ortho_sim.sqrt();
                log!("  spelling BERT: '{}' bert={:.3} × sqrt(ortho {:.2}) = {:.3}", candidate, bert_score, ortho_sim, final_score);
                (candidate.clone(), final_score)
            })
            .collect();
        rescored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

        // Update the matching writing error's suggestion
        if let Some((best, _)) = rescored.first() {
            for e in &mut self.writing_errors {
                if matches!(e.category, ErrorCategory::Spelling)
                    && e.word.to_lowercase() == pending.error_idx_word
                    && !e.ignored
                {
                    if e.suggestion != *best {
                        log!("spelling BERT upgrade: '{}' → '{}' (was '{}')", e.word, best, e.suggestion);
                        e.suggestion = best.clone();
                    }
                    break;
                }
            }
        }
    }

    /// Handle BERT spelling score response for grammar correction ranking.
    fn handle_grammar_bert_response(&mut self, pending: PendingGrammarBert, scored_candidates: &[(String, f32)]) {
        if scored_candidates.is_empty() { return; }
        // scored_candidates is sorted best-first; find the best that matches our candidates
        // Map scored candidate back to the full (sentence, explanation, rule_name) tuple
        let candidate_map: std::collections::HashMap<&str, &(String, String, String)> = pending.candidates.iter()
            .map(|c| (c.0.as_str(), c))
            .collect();
        // The best scored candidate that matches one of our grammar candidates
        let best = scored_candidates.iter()
            .find_map(|(candidate, score)| {
                candidate_map.get(candidate.as_str()).map(|c| (c, *score))
            });
        if let Some(((best_sentence, best_expl, best_rule), best_score)) = best {
            log!("grammar BERT: best='{}' (score={:.3})", best_sentence, best_score);

            // Update matching grammar error
            for e in &mut self.writing_errors {
                if matches!(e.category, ErrorCategory::Grammar)
                    && e.sentence_context == pending.sentence_context
                    && !e.ignored
                {
                    e.suggestion = best_sentence.clone();
                    e.explanation = best_expl.clone();
                    e.rule_name = best_rule.clone();
                    break;
                }
            }
        }
    }

    /// Handle BERT spelling score response for consonant confusion.
    fn handle_consonant_bert_response(&mut self, pending: PendingConsonantBert, scored_candidates: &[(String, f32)]) {
        if scored_candidates.is_empty() { return; }
        // scored_candidates is sorted best-first; first entry with the original word is the orig score
        let orig_score = scored_candidates.iter()
            .find(|(c, _)| c == &pending.word)
            .map(|(_, s)| *s)
            .unwrap_or(f32::NEG_INFINITY);
        // Find best variant (not the original word)
        let best_alt = scored_candidates.iter()
            .find(|(c, _)| c != &pending.word && pending.variants.contains(c));
        for (candidate, score) in scored_candidates {
            log!("consonant BERT: '{}' orig={:.2}, '{}' score={:.2}", pending.word, orig_score, candidate, score);
        }
        if let Some((best, s_best)) = best_alt {
            if *s_best > orig_score {
                let best = best.clone();
                let corrected_sentence = pending.sentence_ctx.replacen(&pending.word, &best, 1);
                self.pending_consonant_checks.push(WritingError {
                    category: ErrorCategory::Grammar,
                    word: pending.sentence_ctx.clone(),
                    suggestion: corrected_sentence,
                    explanation: format!("«{}» → «{}» (enkel/dobbel konsonant)", pending.word, best),
                    rule_name: format!("consonant_confusion:{}:{}", pending.word, best),
                    sentence_context: pending.sentence_ctx,
                    doc_offset: pending.doc_offset,
                    position: 0,
                    ignored: false,
                    word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: String::new(),
                });
            }
        }
    }

    /// Re-run unified suggestion pipeline for spelling errors that were created before
    /// BERT was available. Once BERT loads, this upgrades ortho-only suggestions to BERT-ranked ones.
    fn upgrade_spelling_suggestions(&mut self) {
        if !self.bert_ready { return; }

        let to_upgrade: Vec<(usize, String, String)> = self.writing_errors.iter().enumerate()
            .filter(|(_, e)| {
                matches!(e.category, ErrorCategory::Spelling)
                    && !e.ignored
                    && e.rule_name == "stavefeil" // not yet BERT-ranked
                    && !e.sentence_context.is_empty()
            })
            .map(|(i, e)| (i, e.word.clone(), e.sentence_context.clone()))
            .collect();

        // Process only 1 upgrade per frame to keep GUI responsive
        if let Some((idx, word, sentence_ctx)) = to_upgrade.into_iter().next() {
            let suggestions = self.find_spelling_suggestions(&word, &sentence_ctx);
            if let Some((best, score)) = suggestions.first() {
                log!("Spelling upgrade: '{}' → '{}' score={:.2}", word, best, score);
                self.writing_errors[idx].suggestion = best.clone();
            }
            self.writing_errors[idx].rule_name = "stavefeil_bert".to_string();
        }
    }

    /// UNIFIED spelling suggestion pipeline. ALL spelling suggestions go through this function.
    /// 1. Generate candidates from ALL sources (fuzzy, prefix, truncated fuzzy, compound, split, wordfreq)
    /// 2. Filter: must be real dictionary word AND produce valid grammar in context
    /// 3. Rank survivors by BERT (or ortho-only if BERT unavailable)
    /// 4. Walk up to 50 candidates for grammar checking
    /// 5. Fallback: BERT MLM predictions filtered by ortho similarity to misspelled word
    fn find_spelling_suggestions(&mut self, word: &str, sentence_ctx: &str) -> Vec<(String, f32)> {
        let word_lower = word.to_lowercase();
        let word_trigrams = Self::trigrams(&word_lower);
        let word_first = word_lower.chars().next().unwrap_or(' ');

        // ── Phase 1: Collect candidates from all sources (immutable checker) ──
        let mut candidates: Vec<String> = Vec::new();
        let mut seen = std::collections::HashSet::new();
        let mut edit_distances: HashMap<String, u32> = HashMap::new();

        {
            let analyzer = match &self.analyzer {
                Some(a) => a,
                None => return Vec::new(),
            };

            // Source 1: Fuzzy Levenshtein matches (distance 2)
            for (w, dist) in analyzer.fuzzy_lookup(&word_lower, 2) {
                let wl = w.to_lowercase();
                if wl == word_lower || wl.len() < 2 { continue; }
                edit_distances.insert(wl.clone(), dist);
                if seen.insert(wl.clone()) { candidates.push(wl); }
            }

            // Source 2: Prefix lookup (missing-letter typos: "fotbal" → "fotball")
            for w in analyzer.prefix_lookup(&word_lower, 20) {
                let wl = w.to_lowercase();
                let extra = wl.len() as i32 - word_lower.len() as i32;
                if extra >= 1 && extra <= 3 {
                    edit_distances.entry(wl.clone()).or_insert(extra as u32);
                    if wl != word_lower && wl.len() >= 2 && seen.insert(wl.clone()) {
                        candidates.push(wl);
                    }
                }
            }

            // Source 3: Prefix with last char removed (typo in final position)
            let char_count = word_lower.chars().count();
            if char_count >= 3 {
                let end_byte = word_lower.char_indices().rev().next().map(|(i, _)| i).unwrap_or(0);
                let shorter = &word_lower[..end_byte];
                for w in analyzer.prefix_lookup(shorter, 20) {
                    let wl = w.to_lowercase();
                    let diff = (wl.len() as i32 - word_lower.len() as i32).unsigned_abs() + 1;
                    edit_distances.entry(wl.clone()).or_insert(diff);
                    if wl != word_lower && wl.len() >= 2 && seen.insert(wl.clone()) {
                        candidates.push(wl);
                    }
                }
            }

            // Source 4: Truncated fuzzy (strip 1-2 trailing chars then fuzzy)
            for strip in 1..=2u32 {
                let chars: Vec<char> = word_lower.chars().collect();
                if chars.len() <= 3 + strip as usize { continue; }
                let truncated: String = chars[..chars.len() - strip as usize].iter().collect();
                for (w, dist) in analyzer.fuzzy_lookup(&truncated, 2) {
                    let wl = w.to_lowercase();
                    edit_distances.entry(wl.clone()).or_insert(dist + strip);
                    if wl != word_lower && wl.len() >= 2 && seen.insert(wl.clone()) {
                        candidates.push(wl);
                    }
                }
            }

            // Source 5: Compound suggestion (skipped — only available on SWI checker, not on Analyzer)

            // Source 6: Split function word ("tilbutikken" → "til butikken")
            if let Some(split) = try_split_function_word(&word_lower, analyzer) {
                let sl = split.to_lowercase();
                if seen.insert(sl.clone()) { candidates.push(sl); }
            }
        } // analyzer borrow dropped

        // Source 7: Wordfreq — common words with trigram overlap
        if let Some(wf) = &self.wordfreq {
            for (w, _freq) in wf.iter() {
                let wl = w.to_lowercase();
                if wl == word_lower || seen.contains(&wl) { continue; }
                if wl.chars().next().unwrap_or(' ') != word_first { continue; }
                let w_trigrams = Self::trigrams(&wl);
                let common = word_trigrams.iter().filter(|t| w_trigrams.contains(t)).count();
                if common >= 2 && seen.insert(wl.clone()) {
                    candidates.push(wl);
                }
            }
        }

        log!("find_spelling_suggestions: {} raw candidates for '{}'", candidates.len(), word_lower);

        // ── Phase 2: Ortho score all candidates ──
        let mut ortho_scored: Vec<(String, f32)> = Vec::new();
        for w in &candidates {
            let w_trigrams = Self::trigrams(w);
            let common = word_trigrams.iter().filter(|t| w_trigrams.contains(t)).count();
            let max_t = word_trigrams.len().max(w_trigrams.len()).max(1);
            let trigram_sim = common as f32 / max_t as f32;

            let prefix_len = word_lower.chars().zip(w.chars())
                .take_while(|(a, b)| a == b).count();
            let max_len = word_lower.chars().count().max(w.chars().count()).max(1);
            let prefix_sim = prefix_len as f32 / max_len as f32;

            let edit_sim = match edit_distances.get(w.as_str()) {
                Some(1) => 0.85,
                Some(2) => 0.65,
                Some(3) => 0.45,
                _ => 0.0,
            };
            let mut ortho_sim = trigram_sim.max(edit_sim).max(prefix_sim);
            // First-char bonus: misspellings usually preserve the initial letter
            if w.chars().next() == Some(word_first) {
                ortho_sim += 0.15;
            }
            ortho_scored.push((w.clone(), ortho_sim));
        }
        ortho_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

        // ── Phase 3: MLM fallback via BERT worker (async) ──
        // MLM forward is sent to the worker. If Phase 4 finds no grammar-valid candidates,
        // Phase 5 will use the MLM predictions when they arrive. For now, we skip MLM
        // fallback in the synchronous path — it's handled async via pending_spelling_bert.
        let mut mlm_fallback_candidates: Vec<(String, f32)> = Vec::new();

        // Send MlmForward request to worker if available (results handled async)
        if let Some(worker) = &mut self.bert_worker {
            let masked_context = if let Some(pos) = sentence_ctx.to_lowercase().find(&word_lower) {
                let before = &sentence_ctx[..pos];
                let after = &sentence_ctx[pos + word_lower.len()..];
                format!("{} <mask>{}", before.trim_end(), after)
            } else {
                format!("{} <mask>", sentence_ctx)
            };
            let _mlm_id = worker.send(|id| bert_worker::BertRequest::MlmForward {
                id, masked_text: masked_context, top_k: 100,
            });
            // MLM results will be picked up in poll_bert_responses() if needed
        }

        log!("find_spelling_suggestions: top 5 ortho for '{}': {:?}", word_lower,
            ortho_scored.iter().take(5).collect::<Vec<_>>());

        // ── Phase 4: Dictionary filter (walk up to 100 by ortho score) ──
        // Grammar filter skipped (checker moved to actor). Dictionary check via analyzer.
        let mut passing: Vec<(String, f32)> = Vec::new();
        {
            let analyzer = match &self.analyzer {
                Some(a) => a,
                None => return ortho_scored, // no analyzer available, return unfiltered
            };

            let mut checked = 0;
            for (candidate, score) in &ortho_scored {
                if checked >= 8 { break; } // keep low to avoid GUI freeze

                // Skip hyphenated candidates when misspelled word has no hyphen
                if !word_lower.contains('-') && candidate.contains('-') {
                    continue;
                }

                // Dictionary check: every word in the candidate must exist
                let words: Vec<&str> = candidate.split_whitespace().collect();
                if words.iter().any(|w| !analyzer.has_word(w)) {
                    continue;
                }

                checked += 1;
                passing.push((candidate.clone(), *score));
            }
        }

        // ── Phase 6: Send async BERT sentence re-ranking (results come via poll_bert_responses) ──
        // Return first grammar-valid candidate immediately. BERT worker will re-rank async.
        if passing.len() > 1 {
            if let Some(worker) = &mut self.bert_worker {
                let sentence_lower = sentence_ctx.to_lowercase();
                let word_pos = sentence_lower.find(&word_lower);
                let (context_before, context_after) = if let Some(pos) = word_pos {
                    (sentence_lower[..pos].trim_end().to_string(), sentence_lower[pos + word_lower.len()..].trim_start().to_string())
                } else {
                    (sentence_lower.clone(), String::new())
                };
                let candidates: Vec<String> = passing.iter().take(30).map(|(c, _)| c.clone()).collect();
                let request_id = worker.send(|id| bert_worker::BertRequest::SpellingScore { id, context_before, context_after, candidates });
                self.pending_spelling_bert.push(PendingSpellingBert {
                    request_id,
                    error_idx_word: word_lower.clone(),
                    error_doc_offset: 0, // will be set by caller
                    candidates: passing.iter().take(30).cloned().collect(),
                });
                log!("  sent {} candidates for BERT spelling score (id={})", passing.len().min(30), request_id);
            }
        }

        log!("find_spelling_suggestions: {} grammar-valid for '{}'", passing.len(), word_lower);
        passing
    }

    /// Generate candidate corrections for a sentence with grammar errors,
    /// score each with BERT, and return the top candidates (up to 3).
    fn best_sentence_corrections(&mut self, sentence: &str, errors: &[GrammarError]) -> Vec<(String, String, String, f32)> {
        let mut candidates: Vec<(String, String, String)> = Vec::new(); // (corrected_sentence, explanation, rule_name)

        // 1. Apply each individual grammar suggestion
        //    If suggestion contains '|' (multiple alternatives), try each one separately
        //    Also generate single/double consonant variants (common dyslexia error)
        for e in errors {
            if !e.suggestion.is_empty() {
                let alternatives: Vec<&str> = e.suggestion.split('|').collect();
                for alt in &alternatives {
                    let fixed = replace_word_at_position(sentence, &e.word, alt);
                    let expl = format!("«{}» -> «{}»: {}", e.word, alt, e.explanation);
                    candidates.push((fixed, expl, e.rule_name.clone()));

                    // Try double/single consonant variants of the suggestion
                    for variant in consonant_variants(alt) {
                        if let Some(analyzer) = &self.analyzer {
                            if analyzer.has_word(&variant) {
                                let vfixed = replace_word_at_position(sentence, &e.word, &variant);
                                let vexpl = format!("«{}» -> «{}»: {}", e.word, variant, e.explanation);
                                candidates.push((vfixed, vexpl, e.rule_name.clone()));
                            }
                        }
                    }
                }
            }
        }

        // 2a. Try removing "å" before the error word, and also removing å + applying substitution
        for e in errors {
            let words: Vec<&str> = sentence.split_whitespace().collect();
            if let Some(pos) = words.iter().position(|w| {
                w.trim_matches(|c: char| c.is_ascii_punctuation()).eq_ignore_ascii_case(&e.word)
            }) {
                if pos > 0 {
                    let prev = words[pos - 1].trim_matches(|c: char| c.is_ascii_punctuation());
                    if prev == "å" {
                        // Try just removing å
                        let removed_aa = remove_word_from_sentence(sentence, "å");
                        if removed_aa != sentence {
                            candidates.push((removed_aa.clone(), format!("Fjernet «å» foran «{}».", e.word), e.rule_name.clone()));
                        }
                        // Try removing å AND applying the substitution
                        // e.g. "har å gikk" → "har gått" (remove å, replace gikk with suggestion)
                        if !e.suggestion.is_empty() {
                            let first_alt = e.suggestion.split('|').next().unwrap_or(&e.suggestion);
                            let combined = replace_word_at_position(&removed_aa, &e.word, first_alt);
                            if combined != sentence && combined != removed_aa {
                                candidates.push((combined, format!("«å {}» → «{}»", e.word, first_alt), e.rule_name.clone()));
                            }
                        }
                    }
                }
            }
        }

        // 2b. Try removing each error word — only if no substitution suggestion exists
        let has_substitution = errors.iter().any(|e| !e.suggestion.is_empty());
        if !has_substitution {
            for e in errors {
                let removed = remove_word_from_sentence(sentence, &e.word);
                if removed != sentence {
                    candidates.push((removed, format!("Fjernet «{}».", e.word), e.rule_name.clone()));
                }
            }
        }

        // 3. Apply all suggestions together (use first alternative for each)
        if errors.len() > 1 {
            let mut all_fixed = sentence.to_string();
            let mut all_expl = Vec::new();
            let mut all_rules = Vec::new();
            for e in errors {
                if !e.suggestion.is_empty() {
                    let first_alt = e.suggestion.split('|').next().unwrap_or(&e.suggestion);
                    all_fixed = replace_word_at_position(&all_fixed, &e.word, first_alt);
                    all_expl.push(format!("«{}» -> «{}»", e.word, first_alt));
                    all_rules.push(e.rule_name.clone());
                }
            }
            if all_fixed != sentence {
                candidates.push((all_fixed, all_expl.join(", "), all_rules.join(",")));
            }
        }

        // Deduplicate by corrected sentence
        candidates.dedup_by(|a, b| a.0 == b.0);
        {
            let mut seen = std::collections::HashSet::new();
            candidates.retain(|(c, _, _)| seen.insert(c.clone()));
        }

        if candidates.is_empty() {
            return Vec::new();
        }

        // Grammar-check skipped (checker moved to actor). Dictionary-filter via analyzer.
        let valid_candidates: Vec<(String, String, String)> = if let Some(analyzer) = &self.analyzer {
            candidates.into_iter()
                .filter(|(c, _, _)| {
                    // Check all words exist in dictionary (catches misspelled suggestions like "spile")
                    for word in c.split_whitespace() {
                        let clean = word.trim_matches(|ch: char| ch.is_ascii_punctuation() || ch == '\u{00ab}' || ch == '\u{00bb}').to_lowercase();
                        if clean.len() >= 2 && !clean.chars().all(|ch| ch.is_ascii_digit()) && !analyzer.has_word(&clean) {
                            eprintln!("    REJECTED (spelling): '{}' — '{}' not in dictionary", c, clean);
                            return false;
                        }
                    }
                    true
                })
                .collect()
        } else {
            candidates
        };

        if valid_candidates.is_empty() {
            log!("  No grammatically valid candidates found");
            return Vec::new();
        }

        // Use first valid candidate immediately. If multiple candidates, send async BERT re-ranking.
        if valid_candidates.len() > 1 {
            if let Some(worker) = &mut self.bert_worker {
                let candidates: Vec<String> = valid_candidates.iter().map(|(c, _, _)| c.clone()).collect();
                let request_id = worker.send(|id| bert_worker::BertRequest::SpellingScore {
                    id,
                    context_before: String::new(),
                    context_after: String::new(),
                    candidates,
                });
                self.pending_grammar_bert.push(PendingGrammarBert {
                    request_id,
                    sentence_context: sentence.to_string(),
                    doc_offset: 0, // set by caller
                    candidates: valid_candidates.clone(),
                });
                log!("  Grammar: sent {} candidates for BERT spelling score (id={})", valid_candidates.len(), request_id);
            }
        }

        log!("  Grammar correction (first valid): '{}'", valid_candidates[0].0);
        vec![(valid_candidates[0].0.clone(), valid_candidates[0].1.clone(), valid_candidates[0].2.clone(), 1.0)]
    }

    fn trigrams(word: &str) -> Vec<String> {
        let chars: Vec<char> = word.chars().collect();
        if chars.len() < 3 {
            return vec![word.to_string()];
        }
        (0..chars.len() - 2)
            .map(|i| chars[i..i+3].iter().collect())
            .collect()
    }

    /// Update last_doc_text, rejecting stale reads that still contain a just-replaced word.
    fn try_update_doc_text(&mut self, doc: String) {
        // Skip when Word Add-in is active — error management handled by add-in
        if self.manager.bridges.iter().any(|b| b.name() == "Word Add-in") {
            return;
        }
        if let Some(ref old_word) = self.last_replaced_word {
            if doc.to_lowercase().split(|c: char| !c.is_alphanumeric()).any(|w| w == old_word.as_str()) {
                return; // stale — still has the old word
            }
            self.last_replaced_word = None;
        }
        self.last_doc_text = doc;
    }

    /// Remove errors whose word has been corrected in the document.
    fn prune_resolved_errors(&mut self) {
        // Use cached text — never read COM here (race condition causes garbled text)
        if self.last_doc_text.is_empty() { return; }
        let doc_text = self.last_doc_text.to_lowercase();
        // Clear underlines for errors that will be removed
        for e in &mut self.writing_errors {
            let should_remove = if e.ignored {
                true
            } else {
                match e.category {
                    ErrorCategory::Grammar => !doc_text.contains(&e.sentence_context.to_lowercase()),
                    ErrorCategory::Spelling => {
                        let word_lower = e.word.to_lowercase();
                        !doc_text.split(|c: char| !c.is_alphanumeric()).any(|w| w == word_lower)
                    }
                    ErrorCategory::SentenceBoundary => !doc_text.contains(&e.word.to_lowercase()),
                }
            };
            if should_remove && e.underlined {
                self.manager.clear_error_underline(e.word_doc_start, e.word_doc_end);
                e.underlined = false;
            }
        }
        self.writing_errors.retain(|e| {
            if e.ignored {
                log!("Pruning ignored: {:?} '{}'", e.category, trunc(&e.word, 40));
                return false;
            }
            let still_present = match e.category {
                ErrorCategory::Grammar => {
                    doc_text.contains(&e.sentence_context.to_lowercase())
                }
                ErrorCategory::Spelling => {
                    let word_lower = e.word.to_lowercase();
                    doc_text.split(|c: char| !c.is_alphanumeric())
                        .any(|w| w == word_lower)
                }
                ErrorCategory::SentenceBoundary => {
                    doc_text.contains(&e.word.to_lowercase())
                }
            };
            if !still_present {
                log!("Error resolved: {:?} '{}' no longer in document", e.category, trunc(&e.word, 40));
            }
            still_present
        });
    }

    /// Prepare grammar scan: read document, split sentences, compute offsets, fill queue.
    /// This is fast (no SWI/BERT calls) and runs every poll when document changes.
    /// The actual per-sentence grammar checking happens incrementally in process_grammar_queue().
    fn update_grammar_errors(&mut self) {
        // When Word Add-in is active, sentence detection and error management
        // is handled by the add-in (process_addin_changed_paragraphs).
        // Skip the old full-doc scanning approach.
        if self.manager.bridges.iter().any(|b| b.name() == "Word Add-in") {
            return;
        }
        // Called on paste/cut/move only — not on every keystroke.
        // Queue processing happens at word boundaries in the main poll loop.

        // Use cached document text — never read COM here (race condition causes garbled text)
        // last_doc_text is updated in poll_context() when Word is confirmed foreground
        let doc_text = if self.last_doc_text.is_empty() { return; } else { self.last_doc_text.clone() };

        // Quick check: if document hasn't changed at all, skip everything
        let doc_hash = hash_str(&doc_text);
        if doc_hash == self.last_doc_hash {
            // Error removal is handled by the add-in's sentence change detection
            // (process_addin_changed_paragraphs clears errors when a sentence changes).
            // No pruning based on full doc text — the add-in bridge doesn't have it.
            if false {
                // Force rescan by invalidating doc hash
                self.last_doc_hash = 0;
            }
            return;
        }
        log!("Doc hash changed ({} → {}), rescanning {} chars", self.last_doc_hash, doc_hash, doc_text.len());
        log!("  Full doc text: '{}'", trunc(&doc_text, 300));
        self.last_doc_hash = doc_hash;

        // Count sentences to detect paste (large change) vs fix (small change)
        let new_sentence_count = split_sentences(&doc_text).len();
        let old_sentence_count = self.last_sentence_count;
        self.last_sentence_count = new_sentence_count;
        let is_major_change = (new_sentence_count as isize - old_sentence_count as isize).unsigned_abs() > 2;

        if is_major_change {
            // Major change (window switch, paste, etc.) — clear all stale state
            log!("Major doc change: {} → {} sentences, clearing all queues + clean hashes",
                old_sentence_count, new_sentence_count);
            self.writing_errors.clear();
            self.spelling_queue.clear();
            self.pending_spelling_bert.clear();
            self.grammar_queue.clear();
            self.grammar_queue_total = 0;
            self.processed_sentence_hashes.clear();
        }

        {
            // On any doc change, prune errors whose text is no longer in the document
            // and clear sentence hashes so those sentences get re-scanned
            let doc_lower = doc_text.to_lowercase();
            let mut pruned_contexts2: Vec<String> = Vec::new();
            self.writing_errors.retain(|e| {
                let keep = doc_lower.contains(&e.word.to_lowercase());
                if !keep { pruned_contexts2.push(e.sentence_context.clone()); }
                keep
            });
            for ctx in &pruned_contexts2 {
                self.processed_sentence_hashes.remove(&hash_str(ctx));
            }
        }

        let mut sentences = split_sentences(&doc_text);
        // Track which original sentences were sub-split by Prolog
        // (original_text → split_sentences) for boundary suggestions
        let mut prolog_splits: Vec<(String, Vec<String>)> = Vec::new();

        // If no punctuated sentences but text exists, try Prolog sentence splitting
        if sentences.is_empty() && nostos_cognio::punctuation::needs_punctuation_check(&doc_text) {
            let doc_h = hash_str(&doc_text);
            if !self.prolog_checked_hashes.contains(&doc_h) {
                // Prolog splitting skipped (checker moved to actor, split_by_prolog not on Analyzer)

                // BERT sentence splitting skipped — Prolog handles most cases.
                // (BERT model is owned by worker thread, not accessible here)
                self.prolog_checked_hashes.insert(doc_h);
            }
        }

        // Also check each punctuated sentence for internal boundaries
        // e.g. "Jeg spiller fotball jeg går tur." — has final period but missing internal one
        // Time-budgeted: process max 10ms of Prolog splits per frame to avoid freezing.
        // Unprocessed sentences pass through without splitting — they'll be split next frame.
        // Prolog sub-splitting skipped (checker moved to actor, split_by_prolog not on Analyzer)

        if sentences.is_empty() {
            return;
        }

        // All sentences in the current document with their char offsets.
        // For duplicate sentences, each occurrence gets its own offset.
        let doc_lower = doc_text.to_lowercase();
        let new_sentences: Vec<(String, usize)> = {
            let trimmed_list: Vec<String> = sentences.iter()
                .map(|s| s.trim().to_string())
                .filter(|s| !s.is_empty())
                .collect();
            let mut result = Vec::new();
            let mut claimed_offsets: Vec<usize> = Vec::new();
            for s in &trimmed_list {
                let s_lower = s.to_lowercase();
                // Find the next unclaimed occurrence of this sentence in the doc
                let mut search_from = 0usize;
                let mut found_offset = None;
                while let Some(byte_pos) = doc_lower[search_from..].find(&s_lower) {
                    let abs_byte = search_from + byte_pos;
                    // Use doc_lower for char_offset since abs_byte is from doc_lower
                    let char_offset = if abs_byte <= doc_lower.len() && doc_lower.is_char_boundary(abs_byte) {
                        doc_lower[..abs_byte].chars().count()
                    } else {
                        0
                    };
                    if !claimed_offsets.contains(&char_offset) {
                        found_offset = Some(char_offset);
                        claimed_offsets.push(char_offset);
                        break;
                    }
                    // Advance past this match — must land on a char boundary
                    search_from = abs_byte + 1;
                    while search_from < doc_lower.len() && !doc_lower.is_char_boundary(search_from) {
                        search_from += 1;
                    }
                }
                result.push((s.clone(), found_offset.unwrap_or(0)));
            }
            result
        };

        // --- Step 0: Re-map existing errors to new offsets, remove stale ones ---
        let mut stale_sentences: Vec<String> = Vec::new();
        log!("  Step 0: re-mapping {} errors to {} sentences", self.writing_errors.len(), new_sentences.len());
        for e in &self.writing_errors {
            if !e.ignored {
                log!("    error: '{}' in '{}' off={}", trunc(&e.word, 20), trunc(&e.sentence_context, 40), e.doc_offset);
            }
        }
        {
            let mut available_offsets: std::collections::HashMap<String, Vec<usize>> = std::collections::HashMap::new();
            for (s, off) in &new_sentences {
                available_offsets.entry(s.clone()).or_default().push(*off);
            }
            let mut claimed: std::collections::HashMap<String, Vec<usize>> = std::collections::HashMap::new();
            for e in &mut self.writing_errors {
                if e.ignored { continue; }
                let key = e.sentence_context.clone();
                // Try 1: exact sentence match with offset claiming (handles duplicate sentences)
                if let Some(offsets) = available_offsets.get(&key) {
                    let already_claimed = claimed.entry(key.clone()).or_default();
                    if let Some(&off) = offsets.iter().find(|o| !already_claimed.contains(o)) {
                        e.doc_offset = off;
                        already_claimed.push(off);
                        continue; // success
                    }
                    // All offsets claimed — fall through to word search
                }
                // Try 2: find any new sentence containing this error word
                // But if the sentence changed, the error may no longer be valid — mark stale
                let word_lower = e.word.to_lowercase();
                let mut relocated = false;
                for (s, off) in &new_sentences {
                    if s.to_lowercase().contains(&word_lower) {
                        log!("Relocated error '{}' to sentence '{}' off={} — clearing hash for rescan", trunc(&e.word, 20), trunc(s, 40), off);
                        // Sentence changed — error may be stale. Remove it and force rescan.
                        stale_sentences.push(e.sentence_context.clone());
                        stale_sentences.push(s.clone());
                        e.word.clear(); // mark for removal
                        relocated = true;
                        break;
                    }
                }
                if !relocated {
                    let word_copy = e.word.clone();
                    stale_sentences.push(e.sentence_context.clone());
                    e.word.clear(); // mark for removal below
                    log!("Stale error: '{}' (not found in any sentence) — will rescan", trunc(&word_copy, 40));
                }
            }
        }
        // Remove errors marked for removal and clear their sentence hashes for rescan
        self.writing_errors.retain(|e| !e.word.is_empty());
        for ctx in &stale_sentences {
            self.processed_sentence_hashes.remove(&hash_str(ctx));
        }

        // --- Step 1: Sentence boundary suggestions (shown first, highest priority) ---
        for (original_text, split_sents) in &prolog_splits {
            // Only suggest if we haven't already suggested for this exact text
            if self.writing_errors.iter().any(|e| {
                matches!(e.category, ErrorCategory::SentenceBoundary)
                    && e.word == *original_text
                    && !e.ignored
            }) {
                continue;
            }
            // Build the punctuated version from the split sentences
            let punctuated = split_sents.join(" ");
            // Skip if suggestion is same as original
            if punctuated.trim() == original_text.trim() {
                continue;
            }
            eprintln!("Sentence boundary suggestion: '{}' -> '{}'",
                trunc(&original_text, 60),
                trunc(&punctuated, 60));

            self.writing_errors.push(WritingError {
                category: ErrorCategory::SentenceBoundary,
                word: original_text.clone(),
                suggestion: punctuated,
                explanation: format!("Setningsgrense: teksten ser ut til å inneholde {} setninger uten punktum.", split_sents.len()),
                rule_name: "setningsgrense".to_string(),
                sentence_context: original_text.clone(),
                doc_offset: 0,
                position: 0,
                ignored: false,
                word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: String::new(),
            });
        }

        // --- Step 2: Check each sentence — skip already-known ones, only scan new ---
        // Clear spell-check dedup so words get re-checked with correct sentence/offset
        self.last_spell_checked_word.clear();
        log!("  {} sentences to process, {} existing errors:", new_sentences.len(), self.writing_errors.len());
        for (s, off) in &new_sentences {
            log!("    [off={}] '{}'", off, trunc(s, 80));
        }
        let mut queue: Vec<(String, usize)> = Vec::new();
        for (trimmed, doc_offset) in &new_sentences {
            let sent_h = hash_str(trimmed);

            // Also skip if this occurrence already has errors recorded (re-mapped in Step 0)
            let has_errors = self.writing_errors.iter().any(|e| {
                e.sentence_context == *trimmed && e.doc_offset == *doc_offset && !e.ignored
            });

            // Skip sentences already checked and clean (both grammar AND spelling)
            if self.processed_sentence_hashes.contains(&sent_h) {
                log!("  SKIP (clean hash): '{}'", trunc(trimmed, 60));
                continue;
            }

            let cursor_off = self.context.cursor_doc_offset.unwrap_or(0);
            let doc_char_len = doc_text.chars().count();

            // Spelling: queue words for sentences not yet known-clean
            // Bridge decides whether to skip the word at cursor
            if !has_errors {
                let mut char_pos = *doc_offset;
                for word in trimmed.split_whitespace() {
                    let clean = word.trim_matches(|c: char| c.is_ascii_punctuation() || c == '\u{00ab}' || c == '\u{00bb}');
                    let word_start = char_pos;
                    let word_end = char_pos + word.chars().count();
                    char_pos = word_end + 1; // +1 for the space
                    if clean.is_empty() { continue; }
                    if self.manager.should_skip_word_spelling(cursor_off, word_start, word_end, doc_char_len, &self.context.word) {
                        continue;
                    }
                    self.spelling_queue.push(SpellingQueueItem { word: clean.to_string(), sentence_ctx: trimmed.clone(), paragraph_id: String::new() });
                }
            }
            if has_errors {
                log!("  SKIP (has errors): '{}'", trunc(trimmed, 60));
                continue;
            }

            // Grammar: bridge decides whether to skip this sentence
            if !is_major_change && old_sentence_count > 0 {
                let sent_end = *doc_offset + trimmed.chars().count();
                let ends_with_punct = trimmed.ends_with('.') || trimmed.ends_with('!') || trimmed.ends_with('?');
                if self.manager.should_skip_sentence_grammar(cursor_off, *doc_offset, sent_end, ends_with_punct, doc_char_len, &self.context.word) {
                    log!("  SKIP (bridge): '{}' cursor={} range={}..{}", trunc(trimmed, 60), cursor_off, doc_offset, sent_end);
                    continue;
                }
            }

            let has_spelling_errors = self.writing_errors.iter().any(|e| {
                e.sentence_context == *trimmed
                    && e.doc_offset == *doc_offset
                    && !e.ignored
                    && matches!(e.category, ErrorCategory::Spelling)
            });
            if has_spelling_errors {
                log!("  Skipping grammar check — spelling errors pending in '{}'", trimmed);
                continue;
            }

            queue.push((trimmed.clone(), *doc_offset));
        }

        if !queue.is_empty() {
            log!("Grammar queue: {} sentences to check", queue.len());
            self.grammar_queue_total = queue.len();
            self.grammar_queue = queue;
            self.grammar_scanning = true;
            // Process first one immediately
            self.process_grammar_queue();
        }
    }

    /// Process a few words from the spelling queue per call.
    /// Keeps UI responsive when many words need checking (e.g. English text).
    fn process_spelling_queue(&mut self) {
        let mut processed = 0;
        while let Some(item) = self.spelling_queue.first().cloned() {
            self.spelling_queue.remove(0);
            self.check_spelling(&item.word, &item.sentence_ctx, &item.paragraph_id, 0);
            processed += 1;
            if processed >= 1 { break; } // max 1 word per frame — keeps GUI responsive
        }
        if processed > 0 {
            self.validate_consonant_checks();
        }
    }

    /// Drain changed paragraphs from Word Add-in, split into sentences, and send to grammar actor + spelling queue.
    fn process_addin_changed_paragraphs(&mut self) {
        let actor = match &self.grammar_actor {
            Some(a) => a,
            None => return,
        };

        // Check for reset (new document opened)
        for bridge in &self.manager.bridges {
            if bridge.take_reset() {
                log!("Reset: clearing ALL state");
                self.writing_errors.clear();
                self.processed_sentence_hashes.clear();
                self.paragraph_sentence_hashes.clear();
                self.spelling_queue.clear();
                self.grammar_queue.clear();
                self.grammar_scanning = false;
                self.grammar_errors.clear();
                self.pending_spelling_bert.clear();
                self.pending_grammar_bert.clear();
                self.pending_consonant_bert.clear();
                self.pending_consonant_checks.clear();
                self.last_spell_checked_word.clear();
                self.last_doc_text.clear();
                self.last_doc_hash = 0;
                self.last_doc_approx_len = 0;
                self.last_sentence_count = 0;
                self.prolog_checked_hashes.clear();
                self.completions.clear();
                self.open_completions.clear();
                self.last_checked_sentence.clear();
            }
        }

        // Drain from all bridges that support it
        for bridge in &self.manager.bridges {
            let changed = bridge.drain_changed_paragraphs();
            for p in changed {
                // Clear errors only for CHANGED sentences in this paragraph.
                // Errors for unchanged (clean hash) sentences are kept.

                log!("Addin changed paragraph: '{}' (para={})", trunc(&p.text, 50), trunc(&p.paragraph_id, 10));

                // Split paragraph into sentences
                let sentences = split_sentences(&p.text);
                let new_hashes: Vec<u64> = sentences.iter().map(|s| hash_str(s)).collect();

                // Remove old sentence hashes for this paragraph from clean set
                // and clear errors for sentences that no longer exist
                if let Some(old_hashes) = self.paragraph_sentence_hashes.get(&p.paragraph_id) {
                    for old_h in old_hashes {
                        if !new_hashes.contains(old_h) {
                            // Sentence no longer exists in this paragraph — remove hash + errors
                            self.processed_sentence_hashes.remove(old_h);
                        }
                    }
                }
                // Clear errors for sentences that are no longer in the paragraph
                let new_sentence_set: std::collections::HashSet<String> = sentences.iter().map(|s| s.to_lowercase()).collect();
                let before_count = self.writing_errors.len();
                self.writing_errors.retain(|e| {
                    if e.paragraph_id != p.paragraph_id { return true; }
                    // Keep if sentence still exists in paragraph
                    let keep = new_sentence_set.contains(&e.sentence_context.to_lowercase());
                    if !keep {
                        log!("  Removing stale error: word='{}' sentence='{}'", e.word, trunc(&e.sentence_context, 40));
                    }
                    keep
                });
                if self.writing_errors.len() < before_count {
                    log!("  Cleared {} stale errors for para={}", before_count - self.writing_errors.len(), trunc(&p.paragraph_id, 10));
                }

                // Check each sentence: skip if already processed (hash unchanged)
                for sentence_text in &sentences {
                    let sent_h = hash_str(sentence_text);
                    if self.processed_sentence_hashes.contains(&sent_h) {
                        continue; // Already processed, skip
                    }

                    // This sentence changed — clear its errors
                    let sentence_lower = sentence_text.to_lowercase();
                    self.writing_errors.retain(|e| {
                        !(e.paragraph_id == p.paragraph_id && e.sentence_context.to_lowercase() == sentence_lower)
                    });

                    // Grammar check (async via actor)
                    actor.check_sentence(sentence_text, 0, &p.paragraph_id, 0);

                    // Spelling handled by grammar actor's check_sentence_full — no separate queue needed
                }

                // Store new sentence hashes for this paragraph
                self.paragraph_sentence_hashes.insert(p.paragraph_id.clone(), new_hashes);
            }
        }

        // Handle deleted paragraphs
        for bridge in &self.manager.bridges {
            let deleted = bridge.drain_deleted_paragraphs();
            for para_id in deleted {
                let before = self.writing_errors.len();
                self.writing_errors.retain(|e| e.paragraph_id != para_id);
                if self.writing_errors.len() < before {
                    log!("Cleared {} errors for deleted para={}", before - self.writing_errors.len(), trunc(&para_id, 10));
                }
                // Remove sentence hashes for deleted paragraph
                if let Some(hashes) = self.paragraph_sentence_hashes.remove(&para_id) {
                    for h in hashes {
                        self.processed_sentence_hashes.remove(&h);
                    }
                }
            }
        }

        // Process spelling queue (1 word per call, same as Windows)
        if !self.spelling_queue.is_empty() {
            self.process_spelling_queue();
        }
    }

    /// Send grammar queue items to the grammar actor (non-blocking).
    /// Results come back via poll_grammar_responses().
    fn process_grammar_queue(&mut self) {
        let actor = match &self.grammar_actor {
            Some(a) => a,
            None => return,
        };

        // Send all queued sentences to the actor
        while let Some((trimmed, doc_offset)) = self.grammar_queue.first().cloned() {
            self.grammar_queue.remove(0);

            // Skip if already has errors
            let has_errors = self.writing_errors.iter().any(|e| {
                e.sentence_context == trimmed && e.doc_offset == doc_offset && !e.ignored
            });
            if has_errors {
                continue;
            }

            // Skip if already checked and clean
            let sent_h = hash_str(&trimmed);
            if self.processed_sentence_hashes.contains(&sent_h) {
                continue;
            }

            log!("Grammar send: '{}' (offset={})", trunc(&trimmed, 60), doc_offset);
            actor.check_sentence(&trimmed, doc_offset, "", 0);
        }
        self.grammar_scanning = false;
    }

    /// Poll grammar actor for results and create WritingErrors.
    fn poll_grammar_responses(&mut self) {
        let actor = match &self.grammar_actor {
            Some(a) => a,
            None => return,
        };

        while let Some(resp) = actor.try_recv() {
            let sent_h = hash_str(&resp.sentence);
            self.processed_sentence_hashes.insert(sent_h); // Mark ALL sentences as processed

            // Handle grammar errors
            if !resp.errors.is_empty() {
                for ge in &resp.errors {
                    log!("  Grammar error: '{}' → '{}' ({})", ge.word, ge.suggestion, ge.rule_name);
                }

                let errors_with_suggestions: Vec<_> = resp.errors.iter()
                    .filter(|e| !e.suggestion.is_empty())
                    .collect();

                if !errors_with_suggestions.is_empty() {
                    for (i, ge) in errors_with_suggestions.iter().enumerate() {
                        let first_alt = ge.suggestion.split('|').next().unwrap_or(&ge.suggestion);
                        let corrected = replace_word_at_position(&resp.sentence, &ge.word, first_alt);
                        if corrected.trim() == resp.sentence.trim() {
                            continue;
                        }
                        log!("  Grammar fix: '{}' → '{}' [{}]", ge.word, first_alt, ge.rule_name);
                        self.writing_errors.push(WritingError {
                            category: ErrorCategory::Grammar,
                            word: resp.sentence.to_string(),
                            suggestion: corrected,
                            explanation: format!("«{}» → «{}»: {}", ge.word, first_alt, ge.explanation),
                            rule_name: ge.rule_name.clone(),
                            sentence_context: resp.sentence.to_string(),
                            doc_offset: resp.doc_offset,
                            position: i,
                            ignored: false,
                            word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: resp.paragraph_id.clone(),
                        });
                        // Blue underline for grammar errors
                        if let Some(b) = self.manager.effective_bridge() {
                            b.underline_word(&ge.word, &resp.paragraph_id, "#0000FF");
                        }
                    }
                } else {
                    let first = &resp.errors[0];
                    log!("  Flagging without correction: '{}' ({})", first.word, first.rule_name);
                    self.writing_errors.push(WritingError {
                        category: ErrorCategory::Grammar,
                        word: resp.sentence.to_string(),
                        suggestion: String::new(),
                        explanation: first.explanation.clone(),
                        rule_name: first.rule_name.clone(),
                        sentence_context: resp.sentence.to_string(),
                        doc_offset: resp.doc_offset,
                        position: 0,
                        ignored: false,
                        word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: resp.paragraph_id.clone(),
                    });
                }
            }

            // Handle unknown words (spelling errors) — from check_sentence_full
            for unk in &resp.unknown_words {
                let best = unk.spelling_suggestions.first().cloned().unwrap_or_default();
                if best.is_empty() && unk.split_suggestions.is_empty() {
                    // No suggestions at all — still flag as unknown
                    log!("  Unknown word: '{}' (no suggestions)", unk.word);
                } else {
                    log!("  Spelling: '{}' → '{}' (from grammar checker)", unk.word, best);
                }
                self.writing_errors.push(WritingError {
                    category: ErrorCategory::Spelling,
                    word: unk.word.clone(),
                    suggestion: best,
                    explanation: format!("«{}» finnes ikke i ordboken.", unk.word),
                    rule_name: "stavefeil".to_string(),
                    sentence_context: resp.sentence.to_string(),
                    doc_offset: resp.doc_offset,
                    position: unk.position,
                    ignored: false,
                    word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: resp.paragraph_id.clone(),
                });
                // Underline the error word in Word
                if let Some(b) = self.manager.effective_bridge() {
                    b.underline_word(&unk.word, &resp.paragraph_id, "#FF0000");
                }
            }
        }
    }

    /// Sync red wavy underlines in Word with current writing errors.
    /// - Computes word positions for new errors (word_doc_start/end == 0)
    /// - Applies underline for errors not yet marked
    /// - Clears underline for ignored/removed errors
    fn sync_error_underlines(&mut self) {
        // Compute positions for errors that don't have them yet
        for e in &mut self.writing_errors {
            if e.word_doc_start == 0 && e.word_doc_end == 0 && !e.ignored {
                // For spelling errors, word = the misspelled word
                // For grammar errors, word = whole sentence — extract error word from explanation
                if matches!(e.category, ErrorCategory::SentenceBoundary) {
                    // Underline the whole sentence for boundary errors
                    e.word_doc_start = e.doc_offset;
                    e.word_doc_end = e.doc_offset + e.word.chars().count();
                    log!("Underline: sentence boundary {}..{} for '{}'",
                        e.word_doc_start, e.word_doc_end, trunc(&e.word, 50));
                } else if matches!(e.category, ErrorCategory::Spelling) {
                    // Spelling: underline just the misspelled word
                    let (s, end) = find_word_doc_range(&e.word, &e.sentence_context, e.doc_offset);
                    e.word_doc_start = s;
                    e.word_doc_end = end;
                    log!("Underline: spelling range {}..{} for '{}' in '{}'",
                        s, end, e.word, trunc(&e.sentence_context, 50));
                } else {
                    // Grammar/consonant: underline the whole sentence
                    e.word_doc_start = e.doc_offset;
                    e.word_doc_end = e.doc_offset + e.sentence_context.chars().count();
                    log!("Underline: grammar range {}..{} for '{}'",
                        e.word_doc_start, e.word_doc_end, trunc(&e.sentence_context, 50));
                }
            }
        }

        // Apply underlines for new errors, clear for ignored ones
        for e in &mut self.writing_errors {
            if e.ignored && e.underlined {
                // Error was ignored — remove underline
                self.manager.clear_error_underline(e.word_doc_start, e.word_doc_end);
                e.underlined = false;
            } else if !e.ignored && !e.underlined && e.word_doc_start < e.word_doc_end {
                // New error — apply underline
                let marked = self.manager.mark_error_underline(e.word_doc_start, e.word_doc_end);
                log!("Underline: marking {}..{} rule={} expl='{}' ok={}",
                    e.word_doc_start, e.word_doc_end, e.rule_name, trunc(&e.explanation, 50), marked);
                // Mark as underlined even if bridge doesn't support it (prevents spam)
                e.underlined = true;
            }
        }
    }

    fn sync_embeddings(&mut self) {
        if self.last_embedding_sync.elapsed() < self.embedding_sync_interval {
            return;
        }
        self.last_embedding_sync = Instant::now();

        if let Some(store) = self.embedding_store.as_mut().and_then(|s| Arc::get_mut(s)) {
            if let Some(doc_text) = self.manager.read_document_context() {
                // split_sentences only returns complete sentences (ending .!?)
                // so partial/in-progress sentences are never embedded.
                let sentences = split_sentences(&doc_text);
                match store.sync_sentences(&sentences) {
                    Ok(n) if n > 0 => {
                        eprintln!("Embedded {} new sentences ({} total):", n, sentences.len());
                        for s in &sentences {
                            eprintln!("  '{}'", s);
                        }
                        // New embeddings available — force re-completion so topic boost applies
                        self.last_completed_prefix.clear();
                        self.cached_forward = None;
                        self.cached_right_column = None;
                        self.cached_mtag_supplement = None;
                    }
                    Err(e) => eprintln!("Embedding sync error: {}", e),
                    _ => {}
                }
            }
        }
    }

    fn run_completion(&mut self) {
        let prefix = extract_prefix(&self.context.word);

        // No prefix and no masked context → nothing to do
        if prefix.is_empty() && self.context.masked_sentence.is_none() {
            self.completions.clear();
            self.open_completions.clear();
            return;
        }

        // Build a cache key from prefix + masked sentence
        let cache_key = if prefix.is_empty() {
            format!("__noprefix__{}", self.context.masked_sentence.as_deref().unwrap_or(""))
        } else {
            prefix.to_string()
        };
        if cache_key == self.last_completed_prefix {
            return;
        }
        self.last_completed_prefix = cache_key;
        let t_total = Instant::now();

        // Fill-in-the-blank completions are handled by the background completion thread.
        // This function only handles the legacy complete_word path (no masked sentence).
        if self.context.masked_sentence.is_some() {
            return;
        }

        // Legacy complete_word path (no masked sentence)

        // Build context and run completion (borrows checker immutably for has_word)
        let raw_results = {
            let fallback_fn: Option<Box<dyn Fn(&str) -> bool + '_>> = self.analyzer.as_ref().map(|a| {
                let a = Arc::clone(a);
                Box::new(move |w: &str| a.has_word(w)) as Box<dyn Fn(&str) -> bool>
            });
            let fallback_ref: Option<&dyn Fn(&str) -> bool> = fallback_fn.as_ref().map(|b| b.as_ref());
            let prefix_fn: Option<Box<dyn Fn(&str, usize) -> Vec<String> + '_>> = self.analyzer.as_ref().map(|a| {
                let a = Arc::clone(a);
                Box::new(move |p: &str, limit: usize| a.prefix_lookup(p, limit)) as Box<dyn Fn(&str, usize) -> Vec<String>>
            });
            let prefix_ref: Option<&dyn Fn(&str, usize) -> Vec<String>> = prefix_fn.as_ref().map(|b| b.as_ref());

            #[allow(unreachable_code)]
            if false {
                // Legacy complete_word path disabled — BERT model owned by worker thread
                let model: &mut Model = unreachable!();
                let pi = self.prefix_index.as_ref().unwrap();
                let ctx = {
                    let sentence = &self.context.sentence;
                    let sentence_ctx = sentence.strip_suffix(prefix).unwrap_or(sentence).trim_end();
                    if let Some(doc_text) = self.manager.read_document_context() {
                        let doc_trimmed = doc_text.trim_end();
                        doc_trimmed
                            .strip_suffix(prefix)
                            .unwrap_or(doc_trimmed)
                            .trim_end()
                            .to_string()
                    } else {
                        sentence_ctx.to_string()
                    }
                };

                // Quality controls BPE extension depth and candidate count
                // 0: single-token only (~200ms), 1: 1 step (~800ms), 2: full (~2s)
                let (top_n, max_steps) = match self.quality {
                    0 => (5, 0),
                    1 => (5, 1),
                    _ => (5, 3),
                };

                let t_bert = Instant::now();
                match complete_word(
                    &mut *model,
                    ctx.as_str(),
                    prefix,
                    pi,
                    self.baselines.as_deref(),
                    self.wordfreq.as_deref(),
                    fallback_ref,
                    prefix_ref,
                    self.embedding_store.as_deref(),
                    1.0,   // pmi_weight
                    10.0,  // topic_boost
                    top_n,
                    max_steps,
                ) {
                    Ok(results) => {
                        let bert_ms = t_bert.elapsed().as_millis();
                        Some((results, ctx, bert_ms))
                    }
                    Err(e) => {
                        eprintln!("Completion error: {}", e);
                        self.completions.clear();
            self.open_completions.clear();
                        None
                    }
                }
            } else {
                None
            }
        };

        // Grammar filter skipped (checker moved to actor, check_sentence not available)
        if let Some((results, _ctx, bert_ms)) = raw_results {
            if self.grammar_completion {
                // Grammar filtering not possible without checker; pass completions through
                self.completions = results.into_iter().take(5).collect();
            } else {
                let total_ms = t_total.elapsed().as_millis();
                eprintln!("complete '...{}' bert={}ms total={}ms -> {}",
                    prefix, bert_ms, total_ms,
                    results.iter().map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "));
                self.completions = results.into_iter().take(5).collect();
            }
        }
    }

    fn return_focus_to_app(&self) {
        if let Some(handle) = self.app_handle {
            self.platform.set_foreground(handle);
        }
    }
}

/// Try splitting an unknown word into a function word + remainder.
/// Returns "til butikken" for "tilbutikken", "i morgen" for "imorgen", etc.
/// Only splits after known prepositions/adverbs/conjunctions — these never
/// form legitimate compounds, so if the remainder is a dictionary word,
/// the split is correct.
fn try_split_function_word(word: &str, analyzer: &mtag::Analyzer) -> Option<String> {
    // Norwegian function words that are commonly glued to the next word.
    // Sorted longest-first so "etter" matches before "et".
    const FUNCTION_WORDS: &[&str] = &[
        "gjennom", "mellom", "under", "etter", "langs", "rundt",
        "foran", "bortover", "innover", "utover",
        "forbi", "siden", "etter", "blant",
        "over", "inne", "borte",
        "uten", "utenfor", "innenfor",
        "med", "mot", "ved", "hos", "fra",
        "for", "som", "men",
        "til", "per", "via",
        "på", "av", "om",
        "en", "et", "ei",
        "og", "at",
        "i",
    ];

    let lower = word.to_lowercase();

    // Phase 1: Try known function word prefixes (high confidence)
    for prefix in FUNCTION_WORDS {
        if lower.len() <= prefix.len() + 1 { continue; } // remainder must be ≥2 chars
        if !lower.starts_with(prefix) { continue; }
        let remainder = &lower[prefix.len()..];
        if remainder.len() < 2 { continue; }
        if analyzer.has_word(remainder) {
            return Some(format!("{} {}", prefix, remainder));
        }
    }

    // Phase 2: General split — try all positions where both parts are dictionary words.
    // Catches "løpsakte" → "løp sakte", "huserstore" → "huser store", etc.
    // Both parts must be ≥3 chars to avoid spurious splits on short prefixes.
    let chars: Vec<char> = lower.chars().collect();
    let mut best_split: Option<(String, usize)> = None;
    for split_at in 3..=(chars.len().saturating_sub(3)) {
        let left: String = chars[..split_at].iter().collect();
        let right: String = chars[split_at..].iter().collect();
        if analyzer.has_word(&left) && analyzer.has_word(&right) {
            // Prefer the most balanced split (both parts as long as possible)
            let balance = left.len().min(right.len());
            if best_split.as_ref().map(|(_, b)| balance > *b).unwrap_or(true) {
                best_split = Some((format!("{} {}", left, right), balance));
            }
        }
    }
    best_split.map(|(s, _)| s)
}

/// Generate single↔double consonant variants of a word.
/// "spile" → ["spille"], "balle" → ["bale"], "skinn" → ["skin"], etc.
fn consonant_variants(word: &str) -> Vec<String> {
    let chars: Vec<char> = word.chars().collect();
    let mut variants = Vec::new();
    let consonants = "bcdfghjklmnpqrstvwxz";

    // Try doubling each single consonant
    for i in 0..chars.len() {
        if consonants.contains(chars[i]) {
            // Only double if not already doubled
            if i + 1 >= chars.len() || chars[i + 1] != chars[i] {
                let mut v: Vec<char> = chars.clone();
                v.insert(i + 1, chars[i]);
                variants.push(v.into_iter().collect());
            }
        }
    }

    // Try removing one from each double consonant
    for i in 0..chars.len().saturating_sub(1) {
        if chars[i] == chars[i + 1] && consonants.contains(chars[i]) {
            let mut v: Vec<char> = chars.clone();
            v.remove(i + 1);
            variants.push(v.into_iter().collect());
        }
    }

    variants
}

/// Split text into sentences for embedding.
/// Replace first occurrence of a word (whole word match) in a sentence.
pub(crate) fn replace_word_at_position(sentence: &str, word: &str, replacement: &str) -> String {
    let words: Vec<&str> = sentence.split_whitespace().collect();
    let mut result = Vec::new();
    let mut replaced = false;
    for w in &words {
        let clean = w.trim_matches(|c: char| c.is_ascii_punctuation() || c == '«' || c == '»');
        if !replaced && clean.eq_ignore_ascii_case(word) {
            // Preserve trailing punctuation
            let suffix: String = w.chars().rev().take_while(|c| c.is_ascii_punctuation()).collect::<String>().chars().rev().collect();
            result.push(format!("{}{}", replacement, suffix));
            replaced = true;
        } else {
            result.push(w.to_string());
        }
    }
    result.join(" ")
}

/// Remove first occurrence of a word from a sentence.
fn remove_word_from_sentence(sentence: &str, word: &str) -> String {
    let words: Vec<&str> = sentence.split_whitespace().collect();
    let mut result = Vec::new();
    let mut removed = false;
    for w in &words {
        let clean = w.trim_matches(|c: char| c.is_ascii_punctuation() || c == '«' || c == '»');
        if !removed && clean.eq_ignore_ascii_case(word) {
            removed = true;
        } else {
            result.push(w.to_string());
        }
    }
    result.join(" ")
}


pub(crate) fn hash_str(s: &str) -> u64 {
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    s.hash(&mut hasher);
    hasher.finish()
}

fn split_sentences(text: &str) -> Vec<String> {
    split_sentences_with_offsets(text).into_iter().map(|(s, _)| s).collect()
}

/// Split text into sentences, returning each sentence with its character offset in the text.
fn split_sentences_with_offsets(text: &str) -> Vec<(String, usize)> {
    let mut sentences = Vec::new();
    let mut current = String::new();
    let mut start_byte = 0usize; // byte offset of current sentence start
    let mut pos = 0usize; // current byte position
    for c in text.chars() {
        if current.is_empty() || current.chars().all(|ch| ch.is_whitespace()) {
            // Haven't started real content yet — track the start
            if current.is_empty() {
                start_byte = pos;
            }
        }
        current.push(c);
        pos += c.len_utf8();
        if c == '.' || c == '!' || c == '?' {
            let trimmed = current.trim().to_string();
            if !trimmed.is_empty() && trimmed.len() > 5 {
                // Find actual start: skip leading whitespace
                let leading_ws = current.len() - current.trim_start().len();
                let char_offset = text[..start_byte + leading_ws].chars().count();
                sentences.push((trimmed, char_offset));
            }
            current.clear();
            start_byte = pos;
        }
    }
    // Include trailing text without punctuation (e.g. user still typing)
    let trimmed = current.trim().to_string();
    if !trimmed.is_empty() && trimmed.split_whitespace().count() >= 2 {
        let leading_ws = current.len() - current.trim_start().len();
        let char_offset = text[..start_byte + leading_ws].chars().count();
        sentences.push((trimmed, char_offset));
    }
    sentences
}

fn get_screen_size(platform: &dyn platform::PlatformServices) -> (f32, f32) {
    platform.screen_size()
}

fn icon_button(ui: &mut egui::Ui, icon: &str, hover: &str) -> bool {
    let btn = egui::Button::new(egui::RichText::new(icon).size(14.0))
        .fill(egui::Color32::WHITE)
        .stroke(egui::Stroke::new(1.0, egui::Color32::from_rgb(160, 160, 160)));
    ui.add(btn).on_hover_text(hover).clicked()
}

/// Returns (category, description, examples_wrong, examples_right) for a grammar rule.
fn rule_info(rule_name: &str) -> (&'static str, &'static str, &'static [&'static str], &'static [&'static str]) {
    match rule_name {
        r if r.starts_with("modalverb") => (
            "Modalverb + verbform",
            "Etter modalverb (kan, vil, skal, må, bør) skal verbet stå i infinitiv, ikke i presens eller preteritum.",
            &["Jeg kan spiser.", "Hun vil gikk.", "Vi skal kommer."],
            &["Jeg kan spise.", "Hun vil gå.", "Vi skal komme."],
        ),
        r if r.starts_with("har_") || r == "har_substantiv_som_verb" => (
            "Har/hadde + verbform",
            "Etter «har» eller «hadde» skal verbet stå i perfektum partisipp (har spist, har gått), ikke i presens, preteritum eller infinitiv.",
            &["Jeg har spiser.", "Vi har gikk.", "De har kom.", "Jeg hadde spilet."],
            &["Jeg har spist.", "Vi har gått.", "De har kommet.", "Jeg hadde spilt."],
        ),
        r if r.starts_with("infinitivsmerke") || r == "aa_ikke_verb" => (
            "Infinitivsmerke + verbform",
            "Etter «å» skal verbet stå i infinitiv. Presens eller preteritum etter «å» er feil.",
            &["Jeg liker å spiser.", "Hun prøvde å gikk."],
            &["Jeg liker å spise.", "Hun prøvde å gå."],
        ),
        "og_skal_vaere_aa" => (
            "«og» skal være «å»",
            "Infinitivsmerket «å» forveksles ofte med konjunksjonen «og». Foran et verb i infinitiv skal det stå «å».",
            &["Jeg prøver og spise.", "Hun liker og lese."],
            &["Jeg prøver å spise.", "Hun liker å lese."],
        ),
        "aa_skal_vaere_og" => (
            "«å» skal være «og»",
            "Konjunksjonen «og» forveksles ofte med infinitivsmerket «å». Mellom to sideordnede ledd skal det stå «og».",
            &["Jeg å du.", "Brød å smør."],
            &["Jeg og du.", "Brød og smør."],
        ),
        r if r.starts_with("ubestemt_artikkel") => (
            "Ubestemt artikkel + bestemt substantiv",
            "Etter ubestemt artikkel (en, ei, et) skal substantivet stå i ubestemt form.",
            &["en bilen", "et huset"],
            &["en bil", "et hus"],
        ),
        r if r.starts_with("artikkel_kjoenn") => (
            "Feil kjønn på artikkel",
            "Artikkelen må ha samme kjønn som substantivet. Hankjønn: en, hunkjønn: ei/en, intetkjønn: et.",
            &["en hus", "et bil"],
            &["et hus", "en bil"],
        ),
        r if r.starts_with("dem_som_subjekt") => (
            "«dem» brukt som subjekt",
            "«Dem» er objektsform. Som subjekt skal man bruke «de».",
            &["Dem spiser.", "Dem er fine."],
            &["De spiser.", "De er fine."],
        ),
        r if r.starts_with("dobbel_bestemthet") => (
            "Dobbel bestemthet",
            "I bokmål bruker man vanligvis dobbel bestemthet: bestemt artikkel + bestemt substantiv. Men noen dialekter dropper den ene.",
            &["den bil", "det hus"],
            &["den bilen", "det huset"],
        ),
        r if r.contains("samsvar") => (
            "Samsvarsbøyning",
            "Adjektivet må bøyes i samsvar med substantivet i kjønn og tall.",
            &["en stor hus", "et rød bil"],
            &["et stort hus", "en rød bil"],
        ),
        r if r.starts_with("eiendomsord") => (
            "Feil kjønn på eiendomsord",
            "Eiendomsordet (min, din, sin, etc.) må bøyes i samsvar med substantivets kjønn.",
            &["min hus", "mitt bil"],
            &["mitt hus", "min bil"],
        ),
        _ => (
            "Grammatikkregel",
            "Setningen har en grammatisk feil som ble oppdaget av grammatikksjekken.",
            &[],
            &[],
        ),
    }
}

fn rule_color(rule_name: &str) -> egui::Color32 {
    match rule_name {
        "saerskrivingsfeil" => egui::Color32::from_rgb(220, 50, 50),
        name if name.starts_with("modalverb") => egui::Color32::from_rgb(200, 120, 0),
        name if name.starts_with("dobbelbestemmelse") => egui::Color32::from_rgb(0, 140, 180),
        name if name.contains("samsvar") => egui::Color32::from_rgb(140, 80, 200),
        _ => egui::Color32::from_rgb(180, 130, 0),
    }
}

impl eframe::App for ContextApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        // Allow copying text from labels
        ctx.style_mut(|s| s.interaction.selectable_labels = true);

        // Spawn grammar actor on first update — loads SWI-Prolog on its own thread
        if self.grammar_actor.is_none() && self.analyzer.is_some() {
            self.grammar_actor = Some(grammar_actor::spawn_grammar_actor_with_loader(
                self.platform.swipl_path().to_string(),
                dict_path().to_str().unwrap().to_string(),
                grammar_rules_path().to_str().unwrap().to_string(),
                syntaxer_dir().to_str().unwrap().to_string(),
                std::fs::read_to_string(compound_data_path()).unwrap_or_default(),
                ctx.clone(),
            ));
            log!("Grammar actor spawning (SWI-Prolog loads on actor thread)");
        }

        // Poll grammar actor for results (non-blocking)
        self.poll_grammar_responses();

        // Update errors JSON for /errors endpoint
        {
            let json = format!("[{}]", self.writing_errors.iter()
                .filter(|e| !e.ignored)
                .map(|e| {
                    let cat = match e.category {
                        ErrorCategory::Spelling => "spelling",
                        ErrorCategory::Grammar => "grammar",
                        ErrorCategory::SentenceBoundary => "sentence_boundary",
                    };
                    format!(r#"{{"category":"{}","word":"{}","suggestion":"{}","rule":"{}","sentence":"{}"}}"#,
                        cat, escape_json_str(&e.word), escape_json_str(&e.suggestion),
                        escape_json_str(&e.rule_name), escape_json_str(&e.sentence_context))
                })
                .collect::<Vec<_>>()
                .join(","));
            for bridge in &self.manager.bridges {
                bridge.update_errors_json(&json);
            }
        }

        // Drain changed paragraphs from Word Add-in and send to grammar actor
        self.process_addin_changed_paragraphs();

        // Execute deferred find-and-replace
        if let Some((find, replace, context, doc_offset)) = self.pending_fix.take() {
            log!("pending_fix: bridge='{}' find='{}' replace='{}' offset={}",
                self.manager.active_bridge_name(),
                trunc(&find, 60), trunc(&replace, 60), doc_offset);
            // Clear underline BEFORE replacement (positions are still valid in original doc)
            let find_lower_pre = find.to_lowercase();
            for e in &mut self.writing_errors {
                if e.underlined
                    && (e.word.to_lowercase() == find_lower_pre || e.sentence_context.to_lowercase() == find_lower_pre)
                    && e.doc_offset == doc_offset
                {
                    self.manager.clear_error_underline(e.word_doc_start, e.word_doc_end);
                    e.underlined = false;
                    log!("  Pre-cleared underline {}..{}", e.word_doc_start, e.word_doc_end);
                    break;
                }
            }
            let ok = if context.is_empty() {
                let r = self.manager.find_and_replace(&find, &replace);
                log!("  find_and_replace result: {}", r);
                r
            } else {
                let r = self.manager.find_and_replace_in_context_at(&find, &replace, &context, doc_offset);
                log!("  find_and_replace_in_context result: {}", r);
                r
            };
            if ok {
                // Update cached text by applying replacement locally
                // (don't re-read COM — Word returns stale text immediately after replace)
                if let Some(pos) = self.last_doc_text.to_lowercase().find(&find.to_lowercase()) {
                    let end = pos + find.len();
                    let mut new_text = String::with_capacity(self.last_doc_text.len());
                    new_text.push_str(&self.last_doc_text[..pos]);
                    new_text.push_str(&replace);
                    new_text.push_str(&self.last_doc_text[end..]);
                    log!("ACTION replace: '{}' → '{}' | updated cached text", find, replace);
                    self.last_doc_text = new_text;
                    self.last_replaced_word = Some(find.to_lowercase());
                }
                // Remove the fixed error from the error list
                let find_lower = find.to_lowercase();
                self.writing_errors.retain(|e| {
                    !(e.word.to_lowercase() == find_lower && e.doc_offset == doc_offset)
                });
                log!("  Removed error '{}' from list", find);
                // Document changed — reset doc hash so next poll re-scans
                self.last_doc_hash = 0;
                // Clear clean hashes so ALL sentences get re-checked after fix
                self.processed_sentence_hashes.clear();
                // Clear grammar queue — document changed, stale sentences
                self.grammar_queue.clear();
                self.grammar_scanning = false;
                // Mark replacement as prolog-checked (skip sentence splitting)
                // but NOT as clean — allow grammar/spelling rescan to catch new errors
                let mark_prolog = |text: &str, prolog: &mut std::collections::HashSet<u64>| {
                    let h = hash_str(text);
                    prolog.insert(h);
                    let stripped = text.trim_end_matches(|c: char| c == '.' || c == '!' || c == '?').trim();
                    if !stripped.is_empty() && stripped != text {
                        prolog.insert(hash_str(stripped));
                    }
                };
                mark_prolog(&replace, &mut self.prolog_checked_hashes);
                for sent in replace.split_inclusive(|c: char| c == '.' || c == '!' || c == '?') {
                    let trimmed = sent.trim();
                    if !trimmed.is_empty() {
                        mark_prolog(trimmed, &mut self.prolog_checked_hashes);
                    }
                }
                // Remove the fixed error and adjust offsets of remaining errors
                let find_lower = find.to_lowercase();
                let len_delta = replace.chars().count() as isize - find.chars().count() as isize;
                self.writing_errors.retain(|e| {
                    !(e.word.to_lowercase() == find_lower && e.doc_offset == doc_offset)
                });
                // Shift doc_offset of errors after the fix point
                for e in &mut self.writing_errors {
                    if e.doc_offset > doc_offset {
                        e.doc_offset = (e.doc_offset as isize + len_delta).max(0) as usize;
                    }
                }
                // Force rescan so Step 0 re-maps remaining errors to new offsets
                self.last_doc_hash = 0;
                self.processed_sentence_hashes.remove(&hash_str(&context));
                let stripped_ctx = context.trim_end_matches(|c: char| c == '.' || c == '!' || c == '?').trim();
                if !stripped_ctx.is_empty() && stripped_ctx != context {
                    self.processed_sentence_hashes.remove(&hash_str(stripped_ctx));
                }
                self.grammar_queue.clear();
                self.grammar_scanning = false;
                log!("Fix applied: '{}' removed, {} remaining errors offset-adjusted by {}", find, self.writing_errors.len(), len_delta);
            }
        }

        // OCR: poll clipboard for new screenshots
        if let Some(ocr) = &mut self.ocr {
            let was_pending = ocr.has_pending_image();
            ocr.poll();
            // Grab focus when a new screenshot is detected
            if !was_pending && ocr.has_pending_image() {
                ctx.send_viewport_cmd(egui::ViewportCommand::Focus);
            }
        }

        // OCR: check if background OCR finished
        if let Some(rx) = &self.ocr_receiver {
            if let Ok(result) = rx.try_recv() {
                match result {
                    Ok(text) => {
                        eprintln!("OCR complete: {} chars", text.len());
                        if !text.is_empty() {
                            if self.ocr_copy_mode {
                                self.platform.copy_to_clipboard(&text);
                            } else {
                                tts::speak_word(&text);
                            }
                        }
                        self.ocr_text = Some(text);
                    }
                    Err(e) => {
                        eprintln!("OCR error: {}", e);
                        self.ocr_text = None;
                    }
                }
                self.ocr_receiver = None;
            }
        }

        // Startup: poll background loading threads
        if let Some(rx) = &self.startup_rx {
            while let Ok(item) = rx.try_recv() {
                match item {
                    StartupItem::Completer { model, prefix_index, baselines, wordfreq, embedding_store, errors } => {
                        // Store data FIRST so worker gets the right values
                        self.prefix_index = prefix_index;
                        self.baselines = baselines;
                        self.wordfreq = wordfreq;
                        self.embedding_store = embedding_store;
                        if let Some(m) = model {
                            self.bert_worker = Some(bert_worker::spawn_bert_worker(
                                m,
                                ctx.clone(),
                                build_bpe_completions,
                                build_mtag_completions,
                                build_right_completions,
                                Arc::new(self.prefix_index.clone().unwrap_or_default()),
                                self.baselines.clone(),
                                self.wordfreq.as_ref().cloned(),
                                self.embedding_store.clone(),
                                self.analyzer.clone(),
                            ));
                            self.bert_ready = true;
                        }
                        self.load_errors.extend(errors);
                        self.startup_done.push("NorBERT4".into());
                        // Force rescan — spelling was skipped while BERT was loading
                        self.last_doc_hash = 0;
                        self.processed_sentence_hashes.clear();
                        log!("Startup: NorBERT4 completer ready (bert_worker spawned)");
                    }
                }
            }
            if self.startup_done.len() >= self.startup_total {
                self.startup_rx = None;
            }
        }

        // Whisper lazy-load: check if models finished loading
        if let Some(rx) = &self.whisper_load_rx {
            let mut done_count = 0;
            let expected = if self.whisper_mode == 0 { 1 } else { 2 }; // tiny=1, beste=2
            while let Ok(item) = rx.try_recv() {
                match item {
                    WhisperLoadItem::Final(Ok(engine)) => {
                        log!("Whisper: final model loaded");
                        self.whisper_engine = Some(Arc::new(Mutex::new(engine)));
                        if self.whisper_mode == 1 {
                            self.whisper_load_status = "Stor modell lastet (medium-q5)".into();
                        }
                    }
                    WhisperLoadItem::Final(Err(e)) => {
                        log!("Whisper final model failed: {}", e);
                        self.load_errors.push(format!("Whisper: {}", e));
                    }
                    WhisperLoadItem::Streaming(Ok(engine)) => {
                        log!("Whisper: streaming model loaded");
                        self.whisper_streaming = Some(Arc::new(Mutex::new(engine)));
                        self.whisper_load_status = "Hurtigmodell lastet (base), venter på stor modell...".into();
                    }
                    WhisperLoadItem::Streaming(Err(e)) => {
                        log!("Whisper streaming model failed: {}", e);
                        self.load_errors.push(format!("Whisper-streaming: {}", e));
                    }
                }
                done_count += 1;
            }
            // Check if all expected models are loaded (or failed)
            let loaded = self.whisper_engine.is_some() as usize
                + self.whisper_streaming.is_some() as usize
                + self.load_errors.iter().filter(|e| e.starts_with("Whisper")).count();
            if loaded >= expected || done_count >= expected {
                self.whisper_load_rx = None;
                self.whisper_loading = false;
                log!("Whisper: all models ready");
                // Auto-start recording if user pressed mic while loading
                if self.whisper_pending_record {
                    self.whisper_pending_record = false;
                    if self.whisper_engine.is_some() {
                        let final_eng = self.whisper_engine.as_ref().unwrap().clone();
                        let stream_eng = self.whisper_streaming.as_ref().unwrap_or(&final_eng).clone();
                        match stt::start_recording(final_eng, stream_eng) {
                            Ok(handle) => {
                                log!("Microphone recording auto-started after load");
                                self.mic_handle = Some(handle);
                                self.mic_result_text = None;
                            }
                            Err(e) => log!("Microphone error: {}", e),
                        }
                    }
                }
            }
        }

        // Microphone: check for whisper transcription results (partial or final)
        if let Some(handle) = &self.mic_handle {
            // Drain all available results, keep the latest
            while let Ok(result) = handle.result_rx.try_recv() {
                if result.partial {
                    log!("Whisper partial: '{}'", trunc(&result.text, 60));
                    self.mic_result_text = Some(result.text);
                } else {
                    log!("Whisper final: '{}'", trunc(&result.text, 60));
                    self.mic_result_text = Some(result.text);
                    self.mic_handle = None;
                    self.mic_transcribing = false;
                    ctx.request_repaint(); // repaint immediately to show final result
                    break;
                }
            }
        }
        // Keep repainting while waiting for whisper results
        if self.mic_handle.is_some() || self.mic_transcribing {
            ctx.request_repaint_after(Duration::from_millis(100));
        }

        // Poll for new context
        if self.last_poll.elapsed() >= self.poll_interval {
            self.last_poll = Instant::now();

            let ctx_result = self.manager.read_context();
            if ctx_result.is_none() {
                // Only log once per second to avoid spam
                static LAST_NONE_LOG: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
                let now_ms = self.last_poll.elapsed().as_millis() as u64;
                let last = LAST_NONE_LOG.load(std::sync::atomic::Ordering::Relaxed);
                if now_ms.wrapping_sub(last) > 1000 || last == 0 {
                    LAST_NONE_LOG.store(now_ms, std::sync::atomic::Ordering::Relaxed);
                    log!("read_context() returned None (bridge='{}')", self.manager.active_bridge_name());
                }
            }
            if let Some(new_ctx) = ctx_result {
                log!("Context: word='{}' sentence='{}' masked={} offset={:?}",
                    trunc(&new_ctx.word, 20), trunc(&new_ctx.sentence, 40),
                    new_ctx.masked_sentence.is_some(), new_ctx.cursor_doc_offset);
                // Clear ALL stale state when switching between bridges
                if self.manager.bridge_switched {
                    self.manager.bridge_switched = false;
                    log!("Bridge switched — clearing {} errors, {} spelling queue, {} pending BERT, {} grammar queue",
                        self.writing_errors.len(), self.spelling_queue.len(),
                        self.pending_spelling_bert.len(), self.grammar_queue.len());
                    self.writing_errors.clear();
                    self.spelling_queue.clear();
                    self.pending_spelling_bert.clear();
                    self.pending_grammar_bert.clear();
                    self.pending_consonant_bert.clear();
                    self.grammar_queue.clear();
                    self.grammar_queue_total = 0;
                    self.processed_sentence_hashes.clear();
                    self.last_doc_hash = 0;
                    // Do NOT reset last_sentence_count to 0 — that causes a
                    // false "major doc change" on the very next read, which
                    // clears the BERT queue before results arrive.
                }
                // Update full doc text now — Word is confirmed foreground
                if let Some(doc) = self.manager.read_full_document() {
                    self.try_update_doc_text(doc);
                }
                let fg = self.platform.foreground_app();
                // Only update caret position if valid (not 0,0 which means unknown)
                if let Some((x, y)) = new_ctx.caret_pos {
                    if x != 0 || y != 0 {
                        self.last_caret_pos = Some((x, y));
                    }
                }
                // Only update context if we got something useful — don't overwrite
                // good context with empty when our own window is focused
                if !new_ctx.word.is_empty() || !new_ctx.sentence.is_empty() || new_ctx.masked_sentence.is_some() {
                    // Save the foreground HWND only when we got useful context
                    // (prevents saving Slack/terminal HWND when just switching windows)
                    if fg.handle != 0 {
                        let our_title = "NorskTale";
                        if !fg.title.contains(our_title)
                            && !fg.title.starts_with("Forslag")
                            && !fg.title.starts_with("Regelinfo")
                        {
                            self.app_handle = Some(fg.handle);
                            self.manager.set_target_hwnd(fg.handle);
                        }
                    }
                    // Cursor moved or word changed — clear stale suggestions and queues
                    if new_ctx.masked_sentence != self.context.masked_sentence
                        || new_ctx.word != self.context.word
                    {
                        self.completions.clear();
                        self.open_completions.clear();
                        if !self.grammar_queue.is_empty() {
                            eprintln!("Cursor moved — clearing {} stale grammar queue items", self.grammar_queue.len());
                            self.grammar_queue.clear();
                            self.grammar_scanning = false;
                        }
                    }
                    // Detect paste/cut/move: large jump in text length triggers full doc scan
                    if let Some(ref masked) = new_ctx.masked_sentence {
                        let approx_len = masked.len() - "<mask>".len() + new_ctx.word.len();
                        let big_change = self.last_doc_approx_len == 0
                            || (approx_len as isize - self.last_doc_approx_len as isize).unsigned_abs() > 20;
                        self.last_doc_approx_len = approx_len;
                        // For browser: read clean text from extension file
                        // (masked_sentence glues <mask> to prefix — not valid doc text)
                        if self.manager.last_user_was_browser {
                            if let Some(doc) = self.manager.read_full_document() {
                                if doc.len() > self.last_doc_text.len() / 2 {
                                    self.try_update_doc_text(doc);
                                }
                            }
                        }
                        if big_change {
                            // Paste/cut/move detected — trigger grammar rescan
                            self.update_grammar_errors();
                            self.sync_error_underlines();
                        }
                    }
                    // Check if cursor is on an error word → activate in Grammatikk tab
                    let cursor_word = new_ctx.word.to_lowercase();
                    {
                        let hit = if !cursor_word.is_empty() {
                            // Match by word text (works with add-in where doc offsets aren't available)
                            self.writing_errors.iter().enumerate().find(|(_, e)| {
                                !e.ignored && e.word.to_lowercase() == cursor_word
                            })
                        } else {
                            None
                        };
                        if let Some((idx, e)) = hit {
                            if self.focused_error_idx != Some(idx) {
                                log!("Click hit: word='{}' → error idx={} '{}' rule={}",
                                    cursor_word, idx, trunc(&e.explanation, 40), e.rule_name);
                            }
                            // Switch to Grammatikk tab when clicking on an error word
                            self.selected_tab = 1;
                            if self.focused_error_idx != Some(idx) {
                                self.focused_error_scroll_done = false;
                            }
                            self.focused_error_idx = Some(idx);
                            // Clear old pins, pin the clicked error
                            for e in &mut self.writing_errors { e.pinned = false; }
                            self.writing_errors[idx].pinned = true;
                            self.focused_error_set_time = Instant::now();
                        } else {
                            if self.focused_error_idx.is_some() {
                                if self.focused_error_set_time.elapsed() > Duration::from_millis(500) {
                                    self.focused_error_idx = None;
                                }
                            }
                        }
                    }

                    self.context = new_ctx;
                }
            }

            // Always try to scan for errors — doc hash check makes this cheap when unchanged
            let errors_before = self.writing_errors.len();
            self.update_grammar_errors();
            self.prune_resolved_errors();
            // Upgrade spelling suggestions when BERT becomes available
            self.upgrade_spelling_suggestions();
            if self.writing_errors.len() != errors_before {
                self.sync_error_underlines();
            }

            // Sync document sentences for topic-aware completion
            self.sync_embeddings();

            let mid = is_mid_word(&self.context.word);
            if mid {
                // Mid-word: mark prefix change for debouncing
                // Only trigger run_completion for legacy path (no masked sentence)
                // Fill-in-the-blank is handled by background completion thread
                let prefix = extract_prefix(&self.context.word);
                if self.context.masked_sentence.is_none() && prefix != self.last_completed_prefix {
                    self.last_prefix_change = Instant::now();
                    self.pending_completion = true;
                }
                if !self.selection_mode {
                    self.selected_completion = None;
                }
            } else if self.context.masked_sentence.is_some() {
                // No prefix but have context (e.g. after space): next-word handled by background thread
                if !self.selection_mode {
                    self.selected_completion = None;
                }
                // Word boundary: check spelling of the word just typed (before cursor)
                let sentence = self.context.sentence.clone();
                let cursor_off = self.context.cursor_doc_offset.unwrap_or(0);
                // Get the word just before the cursor (last word in sentence before cursor position)
                let spell_word = sentence.split_whitespace().rev()
                    .find(|w| {
                        let clean = w.trim_matches(|c: char| c.is_ascii_punctuation() || c == '«' || c == '»');
                        !clean.is_empty() && clean.len() >= 2
                    })
                    .map(|w| w.trim_matches(|c: char| c.is_ascii_punctuation() || c == '«' || c == '»').to_string());
                if let Some(ref w) = spell_word {
                    if *w != self.last_spell_checked_word {
                        let errors_before = self.writing_errors.len();
                        log!("Word boundary spell check: '{}' in '{}' (cursor_off={})", w, trunc(&sentence, 50), cursor_off);
                        {
                            let para_id = self.context.paragraph_id.clone();
                            self.check_spelling(w, &sentence, &para_id, 0);
                        }
                        self.last_spell_checked_word = w.clone();
                        // If a new error was found, highlight it (don't auto-switch tab)
                        if self.writing_errors.len() > errors_before {
                            let new_idx = self.writing_errors.len() - 1;
                            self.focused_error_scroll_done = false;
                            self.focused_error_idx = Some(new_idx);
                            for e in &mut self.writing_errors { e.pinned = false; }
                            self.writing_errors[new_idx].pinned = true;
                            self.focused_error_set_time = Instant::now();
                        }
                    }
                }
                self.validate_consonant_checks();
                // Sentence boundary: run grammar check
                self.run_grammar_check();
                // Trigger full doc scan to pick up new errors from typing
                // Only when no grammar queue is already being processed
                if self.grammar_queue.is_empty() {
                    self.update_grammar_errors();
                }
                // Word boundary work: prune, upgrade, drain grammar queue
                // Only process grammar queue when no background forward in flight (avoids model mutex contention)
                self.prune_resolved_errors();
                self.upgrade_spelling_suggestions();
                if !self.spelling_queue.is_empty() {
                    self.process_spelling_queue();
                }
                if !self.grammar_queue.is_empty() && true /* no contention — bert worker owns model */ {
                    self.process_grammar_queue();
                }
                self.sync_error_underlines();
            } else {
                // No word, no context: clear and run grammar
                self.completions.clear();
                self.open_completions.clear();
                self.last_completed_prefix.clear();
                // Check spelling + grammar on the last word/sentence
                let sentence = self.context.sentence.clone();
                let cursor_off = self.context.cursor_doc_offset.unwrap_or(0);
                let spell_word = sentence.split_whitespace().last()
                    .map(|w| w.trim_matches(|c: char| c.is_ascii_punctuation() || c == '«' || c == '»').to_string());
                if let Some(ref w) = spell_word {
                    if !w.is_empty() {
                        {
                            let para_id = self.context.paragraph_id.clone();
                            self.check_spelling(w, &sentence, &para_id, 0);
                        }
                    }
                }
                self.validate_consonant_checks();
                self.run_grammar_check();
                if self.grammar_queue.is_empty() {
                    self.update_grammar_errors();
                }
                // Word boundary work: prune, upgrade, drain grammar queue
                self.prune_resolved_errors();
                self.upgrade_spelling_suggestions();
                if !self.spelling_queue.is_empty() {
                    self.process_spelling_queue();
                }
                if !self.grammar_queue.is_empty() && true /* no contention — bert worker owns model */ {
                    self.process_grammar_queue();
                }
                self.sync_error_underlines();
            }

            // Background completion: debounce + dispatch via BERT worker
            if let Some(masked) = &self.context.masked_sentence.clone() {
                let prefix = extract_prefix(&self.context.word);
                let prefix_lower = prefix.to_lowercase();
                let cache_key = format!("{}|{}", masked, prefix);
                let needs_completion = cache_key != self.last_completed_prefix;

                if needs_completion && cache_key != self.dispatched_key {
                    // Context or prefix changed — reset debounce timer, cancel in-flight
                    self.last_context_change = Instant::now();
                    self.dispatched_key = cache_key.clone();
                    self.completion_cancel.store(true, std::sync::atomic::Ordering::Release);
                    // Immediately filter existing completions by new prefix
                    if !prefix_lower.is_empty() {
                        self.completions.retain(|c| c.word.to_lowercase().starts_with(&prefix_lower));
                    }
                }

                // Dispatch after 300ms idle
                if needs_completion
                    && self.last_context_change.elapsed() >= Duration::from_millis(300)
                {
                    if let Some(worker) = &mut self.bert_worker {
                        // Pre-fetch on main thread (fast, uses analyzer)
                        let matches: Vec<(u32, String)> = self.prefix_index.as_ref()
                            .and_then(|pi| pi.get(&prefix_lower))
                            .cloned()
                            .unwrap_or_default();
                        let mtag_candidates: Vec<String> = if matches.is_empty() && !prefix.is_empty() {
                            self.analyzer.as_ref().map_or(vec![], |a| a.prefix_lookup(&prefix_lower, 50))
                        } else {
                            vec![]
                        };
                        let nearby_words: std::collections::HashSet<String> = {
                            let before_mask = masked.split("<mask>").next().unwrap_or("");
                            let sent_start = before_mask.rfind(|c: char| ".!?".contains(c))
                                .map(|i| i + 1).unwrap_or(0);
                            let current_sent = &before_mask[sent_start..];
                            current_sent.split_whitespace()
                                .map(|w| w.trim_matches(|c: char| !c.is_alphanumeric()).to_lowercase())
                                .filter(|w| w.len() > 1)
                                .collect()
                        };
                        // Trim masked text to ~3 sentences around <mask>
                        let masked_trimmed = {
                            let parts: Vec<&str> = masked.splitn(2, "<mask>").collect();
                            let before = parts[0];
                            let after = parts.get(1).map(|s| *s).unwrap_or("");
                            let trimmed_before = {
                                let bytes = before.as_bytes();
                                let mut cuts = 0;
                                let mut start = 0;
                                for i in (0..bytes.len()).rev() {
                                    if bytes[i] == b'.' || bytes[i] == b'!' || bytes[i] == b'?' {
                                        cuts += 1;
                                        if cuts >= 3 { start = i + 1; break; }
                                    }
                                }
                                before[start..].trim_start()
                            };
                            let trimmed_after = {
                                if let Some(pos) = after.find(|c: char| ".!?".contains(c)) {
                                    &after[..=pos]
                                } else {
                                    after
                                }
                            };
                            format!("{}<mask>{}", trimmed_before, trimmed_after)
                        };
                        let capitalize = prefix.chars().next().map_or(false, |c| c.is_uppercase());
                        let cancel = self.completion_cancel.clone();
                        cancel.store(false, std::sync::atomic::Ordering::Release);

                        let wf = self.wordfreq.clone();
                        let key_clone = cache_key.clone();
                        // Pre-fetch mtag valid words for this prefix (fast lookup on main thread)
                        let mtag_valid: std::collections::HashSet<String> = self.analyzer.as_ref()
                            .map(|a| a.prefix_lookup(&prefix_lower, 100)
                                .into_iter().map(|w| w.to_lowercase()).collect())
                            .unwrap_or_default();
                        let context_for_cw = {
                            let sentence = &self.context.sentence;
                            sentence.strip_suffix(prefix).unwrap_or(sentence).trim_end().to_string()
                        };
                        let (top_n, max_steps) = match self.quality {
                            0 => (15, 0),
                            1 => (15, 1),
                            _ => (15, 3),
                        };
                        log!("Sending CompleteWord: ctx='{}' prefix='{}'", &context_for_cw[context_for_cw.len().saturating_sub(30)..], prefix);
                        worker.send(|id| bert_worker::BertRequest::CompleteWord {
                            id,
                            context: context_for_cw,
                            prefix: prefix.to_string(),
                            capitalize,
                            top_n,
                            max_steps,
                            cache_key: key_clone,
                            masked_text: masked_trimmed,
                        });
                        ctx.request_repaint_after(Duration::from_millis(50));
                    }
                } else if needs_completion {
                    ctx.request_repaint_after(Duration::from_millis(50));
                }
            }
        }

        // Poll ALL BERT worker responses (completions + sentence scoring + MLM)
        self.poll_bert_responses(&ctx);
        // Validate any consonant checks that arrived from BERT
        self.validate_consonant_checks();

        // Idle spelling + grammar queue processing
        if !self.spelling_queue.is_empty() {
            self.process_spelling_queue();
        }
        if !self.grammar_queue.is_empty() {
            self.process_grammar_queue();
        }

        // Phase 1: Ctrl+Space while Word has focus → enter selection mode
        {
            let (ctrl_down, space_down) = self.platform.check_hotkey_state();
            let both_held = ctrl_down && space_down;

            if both_held && !self.ctrl_space_held && !self.selection_mode
                && (!self.completions.is_empty() || !self.open_completions.is_empty())
            {
                self.ctrl_space_held = true;
                // Save app's window handle before stealing focus
                let fg = self.platform.foreground_app();
                self.app_handle = Some(fg.handle);
                self.selected_completion = Some(0);
                self.selection_mode = true;
                self.selected_column = 0;
                // Steal focus to our window
                if let Some(viewport_id) = ctx.input(|i| i.viewport().native_pixels_per_point.map(|_| ())) {
                    let _ = viewport_id;
                }
                ctx.send_viewport_cmd(egui::ViewportCommand::Focus);
            }
            if !both_held {
                self.ctrl_space_held = false;
            }
        }

        // Phase 2: Our window has focus → egui key events for navigation
        if self.selection_mode {
            let mut accept = false;
            let mut cancel = false;
            ctx.input(|i| {
                for event in &i.events {
                    match event {
                        egui::Event::Key { key: egui::Key::ArrowDown, pressed: true, .. } => {
                            let active = if self.selected_column == 1 && !self.open_completions.is_empty() {
                                &self.open_completions
                            } else if self.completions.is_empty() { &self.open_completions } else { &self.completions };
                            let max = active.len();
                            self.selected_completion = Some(match self.selected_completion {
                                None => 0,
                                Some(idx) => (idx + 1).min(max.saturating_sub(1)),
                            });
                        }
                        egui::Event::Key { key: egui::Key::ArrowUp, pressed: true, .. } => {
                            self.selected_completion = Some(match self.selected_completion {
                                None | Some(0) => 0,
                                Some(idx) => idx - 1,
                            });
                        }
                        egui::Event::Key { key: egui::Key::ArrowRight, pressed: true, .. } => {
                            if !self.open_completions.is_empty() && self.selected_column == 0 {
                                self.selected_column = 1;
                                let max = self.open_completions.len();
                                if let Some(idx) = self.selected_completion {
                                    if idx >= max { self.selected_completion = Some(max.saturating_sub(1)); }
                                }
                            }
                        }
                        egui::Event::Key { key: egui::Key::ArrowLeft, pressed: true, .. } => {
                            if !self.completions.is_empty() && self.selected_column == 1 {
                                self.selected_column = 0;
                                let max = self.completions.len();
                                if let Some(idx) = self.selected_completion {
                                    if idx >= max { self.selected_completion = Some(max.saturating_sub(1)); }
                                }
                            }
                        }
                        egui::Event::Key { key: egui::Key::Enter, pressed: true, .. }
                        | egui::Event::Key { key: egui::Key::Space, pressed: true, .. } => {
                            accept = true;
                        }
                        egui::Event::Key { key: egui::Key::Escape, pressed: true, .. } => {
                            cancel = true;
                        }
                        egui::Event::Key { key: egui::Key::P, pressed: true, .. } => {
                            if let Some(idx) = self.selected_completion {
                                let active = if self.selected_column == 1 && !self.open_completions.is_empty() {
                                    &self.open_completions
                                } else if self.completions.is_empty() { &self.open_completions } else { &self.completions };
                                if let Some(comp) = active.get(idx) {
                                    tts::speak_word(&comp.word);
                                }
                            }
                        }
                        egui::Event::Key { key: egui::Key::S, pressed: true, .. } => {
                            let before_cursor = self.manager.read_document_context().unwrap_or_default();
                            let before_text = before_cursor.replace('\r', " ").replace('\n', " ");
                            let sentence_start = before_text.rfind(|c: char| c == '.' || c == '!' || c == '?')
                                .map(|i| i + 1)
                                .unwrap_or(0);
                            let mut sentence = before_text[sentence_start..].trim().to_string();
                            if let Some(idx) = self.selected_completion {
                                let active = if self.selected_column == 1 && !self.open_completions.is_empty() {
                                    &self.open_completions
                                } else if self.completions.is_empty() { &self.open_completions } else { &self.completions };
                                if let Some(comp) = active.get(idx) {
                                    if !sentence.is_empty() {
                                        sentence.push(' ');
                                    }
                                    sentence.push_str(&comp.word);
                                }
                            }
                            if !sentence.is_empty() {
                                tts::speak_word(&sentence);
                            }
                        }
                        _ => {}
                    }
                }
            });

            if accept {
                if let Some(idx) = self.selected_completion {
                    let active = if self.selected_column == 1 && !self.open_completions.is_empty() {
                        &self.open_completions
                    } else if self.completions.is_empty() { &self.open_completions } else { &self.completions };
                    if let Some(comp) = active.get(idx) {
                        let word = comp.word.clone();
                        self.return_focus_to_app();
                        self.manager.replace_word(&word);
                        self.completions.clear();
                        self.open_completions.clear();
                        self.last_completed_prefix.clear();
                        self.last_replace_time = Instant::now();
                        // Force immediate context refresh after replace
                        self.last_poll = Instant::now() - self.poll_interval;
                    }
                }
                self.selection_mode = false;
                self.selected_completion = None;
            }
            if cancel {
                self.return_focus_to_app();
                self.selection_mode = false;
                self.selected_completion = None;
            }
        }

        // Reset selection when both completion lists are empty
        if self.completions.is_empty() && self.open_completions.is_empty() {
            self.selected_completion = None;
            self.selection_mode = false;
        }

        // Debounce: run completion after user stops typing
        if self.pending_completion {
            if self.last_prefix_change.elapsed() >= Duration::from_millis(self.debounce_ms) {
                self.pending_completion = false;
                self.run_completion();
            } else {
                ctx.request_repaint_after(Duration::from_millis(self.debounce_ms));
            }
        }

        // Window sizing (scaled)
        let s = self.ui_scale;
        let has_content = !self.grammar_errors.is_empty() || !self.completions.is_empty() || !self.open_completions.is_empty();
        let recently_replaced = self.last_replace_time.elapsed() < Duration::from_secs(1);
        let win_h = s * if self.selected_tab >= 1 {
            250.0
        } else {
            150.0
        };
        let win_w = s * if self.selected_tab == 0 {
            300.0
        } else {
            420.0
        };

        ctx.send_viewport_cmd(egui::ViewportCommand::Decorations(false));

        // Apply UI scale
        ctx.set_zoom_factor(self.ui_scale);

        // Check if goto freeze has expired
        if let Some(until) = self.goto_freeze_until {
            if Instant::now() >= until {
                self.goto_freeze_until = None;
            }
        }

        if self.follow_cursor && self.goto_freeze_until.is_none() {
            ctx.send_viewport_cmd(egui::ViewportCommand::InnerSize(egui::vec2(win_w, win_h)));
            if let Some((x, y)) = self.last_caret_pos {
                let (screen_w, screen_h) = get_screen_size(&*self.platform);
                let pos_y = if (y as f32 + win_h) > screen_h {
                    y as f32 - win_h - 30.0
                } else {
                    y as f32
                };
                let pos_x = (x as f32).min(screen_w - win_w).max(0.0);

                ctx.send_viewport_cmd(egui::ViewportCommand::OuterPosition(
                    egui::pos2(pos_x, pos_y),
                ));
            }
        }

        // Always keep repainting — we need to poll context from background
        // threads and render BERT results even when Word has focus.
        if self.grammar_scanning {
            ctx.request_repaint();
        } else {
            ctx.request_repaint_after(Duration::from_millis(200));
        }

        // Style
        // Clear the default background so transparency works
        // Determine tab indicators
        let has_completions = !self.completions.is_empty() || !self.open_completions.is_empty();
        let has_grammar = !self.grammar_errors.is_empty()
            || self.writing_errors.iter().any(|e| !e.ignored);

        let panel_frame = egui::Frame::new()
            .fill(egui::Color32::from_rgb(255, 255, 235))
            .stroke(egui::Stroke::new(1.0, egui::Color32::from_rgb(180, 170, 140)))
            .inner_margin(8.0);

        // Startup loading status bar
        if self.startup_rx.is_some() {
            egui::TopBottomPanel::bottom("startup_status").frame(
                egui::Frame::new()
                    .fill(egui::Color32::from_rgb(245, 245, 225))
                    .inner_margin(egui::Margin::symmetric(8, 3))
            ).show(ctx, |ui| {
                ui.horizontal(|ui| {
                    ui.spinner();
                    let progress = self.startup_done.len() as f32 / self.startup_total as f32;
                    let loading: Vec<&str> = ["NorBERT4"]
                        .iter()
                        .filter(|s| !self.startup_done.iter().any(|d| d.starts_with(*s)))
                        .copied()
                        .collect();
                    let label = if loading.is_empty() {
                        "Klar!".to_string()
                    } else {
                        format!("Laster {}...", loading.join(", "))
                    };
                    ui.add(egui::ProgressBar::new(progress)
                        .text(label)
                        .desired_width(ui.available_width())
                        .desired_height(14.0));
                });
            });
            ctx.request_repaint_after(Duration::from_millis(100));
        }

        egui::CentralPanel::default().frame(panel_frame).show(ctx, |ui| {
            // Tab bar with painted dot indicators
            let tts_speaking = tts::is_speaking();
            let ocr_is_busy = self.ocr_receiver.is_some();
            ui.horizontal(|ui| {
                let sep = egui::Color32::from_rgb(180, 170, 140);
                let active = egui::Color32::from_rgb(0, 70, 160);
                let inactive = egui::Color32::from_rgb(100, 100, 100);

                // --- Left side: 💡 ●✏ | 🎤 ▶ ---

                // 💡 Forslag (suggestions tab)
                let innhold_color = if self.selected_tab == 0 { active } else { inactive };
                if ui.add(egui::Label::new(
                    egui::RichText::new("\u{1F4A1}").size(16.0).color(innhold_color)
                ).sense(egui::Sense::click())).on_hover_text("Forslag").clicked() {
                    self.selected_tab = 0;
                }

                // ● ✏ Grammatikk (dot + pen)
                let dot_color = if has_grammar { egui::Color32::from_rgb(220, 50, 50) }
                    else { egui::Color32::from_rgb(0, 180, 60) };
                let (dot_rect, _) = ui.allocate_exact_size(egui::vec2(10.0, 14.0), egui::Sense::hover());
                let center = egui::pos2(dot_rect.min.x + 5.0, dot_rect.center().y);
                ui.painter().circle_filled(center, 4.0, dot_color);
                let gram_color = if self.selected_tab == 1 { active } else { inactive };
                if ui.add(egui::Label::new(
                    egui::RichText::new("\u{270F}").size(16.0).color(gram_color)
                ).sense(egui::Sense::click())).on_hover_text("Grammatikk").clicked() {
                    self.selected_tab = 1;
                }

                ui.add_space(2.0);
                ui.label(egui::RichText::new("|").size(12.0).color(sep));
                ui.add_space(2.0);

                // 🎤 Microphone slot: 🎤 idle, ■ recording, ⏳ transcribing
                let mic_recording = stt::is_recording() || self.mic_transcribing;
                if mic_recording {
                    if self.mic_transcribing {
                        ui.add(egui::Label::new(
                            egui::RichText::new("⏳").size(13.0)
                        )).on_hover_text("Transkriberer...");
                    } else {
                        if ui.add(egui::Button::new(
                            egui::RichText::new("■").size(12.0).color(egui::Color32::WHITE)
                        ).fill(egui::Color32::from_rgb(200, 40, 40))
                         .min_size(egui::vec2(22.0, 16.0))
                        ).on_hover_text("Stopp opptak").clicked() {
                            if let Some(handle) = &self.mic_handle {
                                handle.stop();
                                self.mic_transcribing = true;
                            }
                        }
                    }
                } else if self.whisper_loading {
                    ui.add(egui::Label::new(
                        egui::RichText::new("⏳").size(13.0)
                    )).on_hover_text("Laster talemodell...");
                    ctx.request_repaint_after(Duration::from_millis(100));
                } else {
                    let mic_color = inactive;
                    let whisper_ready = self.whisper_engine.is_some();
                    if ui.add(egui::Label::new(
                        egui::RichText::new("\u{1F3A4}").size(13.0).color(mic_color)
                    ).sense(egui::Sense::click())).on_hover_text("Talegjenkjenning").clicked() {
                            if whisper_ready {
                                // Models already loaded — start recording immediately
                                let final_eng = self.whisper_engine.as_ref().unwrap().clone();
                                let stream_eng = self.whisper_streaming.as_ref().unwrap_or(&final_eng).clone();
                                match stt::start_recording(final_eng, stream_eng) {
                                    Ok(handle) => {
                                        log!("Microphone recording started");
                                        self.mic_handle = Some(handle);
                                        self.mic_result_text = None;
                                    }
                                    Err(e) => log!("Microphone error: {}", e),
                                }
                            } else {
                                #[cfg(target_os = "macos")]
                                {
                                    // macOS: live streaming with Apple SFSpeechRecognizer
                                    match stt::start_recording_live() {
                                        Ok(handle) => {
                                            log!("Microphone recording started (Apple STT)");
                                            self.mic_handle = Some(handle);
                                            self.mic_result_text = None;
                                        }
                                        Err(e) => log!("Microphone error: {}", e),
                                    }
                                }
                                #[cfg(target_os = "windows")]
                                {
                                    // Windows: lazy-load Whisper models, then auto-start recording
                                    self.whisper_loading = true;
                                    self.whisper_pending_record = true;
                                    self.whisper_load_status = if self.whisper_mode == 0 {
                                        "Laster talemodell (tiny, 75 MB)...".into()
                                    } else {
                                        "Laster talemodeller (base + medium-q5, 690 MB)...".into()
                                    };
                                    let (tx, rx) = std::sync::mpsc::channel();
                                    self.whisper_load_rx = Some(rx);
                                    let mode = self.whisper_mode;
                                    let dll_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                                        .join("../../whisper-build/bin/Release")
                                        .to_string_lossy().to_string();
                                    if mode == 0 {
                                        let dll = dll_dir.clone();
                                        std::thread::spawn(move || {
                                            let model_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                                                .join("../contexter-repo/training-data/ggml-nb-whisper-tiny.bin")
                                                .to_string_lossy().to_string();
                                            let _ = tx.send(WhisperLoadItem::Final(
                                                stt::WhisperEngine::load(&dll, &model_path).map(|e| Box::new(e) as Box<dyn stt::SttEngine>)
                                            ));
                                        });
                                    } else {
                                        let tx2 = tx.clone();
                                        let dll2 = dll_dir.clone();
                                        std::thread::spawn(move || {
                                            let model_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                                                .join("../contexter-repo/training-data/ggml-nb-whisper-base.bin")
                                                .to_string_lossy().to_string();
                                            let _ = tx2.send(WhisperLoadItem::Streaming(
                                                stt::WhisperEngine::load(&dll2, &model_path).map(|e| Box::new(e) as Box<dyn stt::SttEngine>)
                                            ));
                                        });
                                        std::thread::spawn(move || {
                                            let model_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                                                .join("../contexter-repo/training-data/ggml-nb-whisper-medium-q5.bin")
                                                .to_string_lossy().to_string();
                                            let _ = tx.send(WhisperLoadItem::Final(
                                                stt::WhisperEngine::load(&dll_dir, &model_path).map(|e| Box::new(e) as Box<dyn stt::SttEngine>)
                                            ));
                                        });
                                    }
                                    log!("Whisper: lazy-loading models (mode={})", mode);
                                }
                            }
                        }
                    }

                // ▶ Speak selection (same group as 🎤)
                if tts_speaking || ocr_is_busy {
                    if ui.add(egui::Button::new(
                        egui::RichText::new("■").size(12.0).color(egui::Color32::WHITE)
                    ).fill(egui::Color32::from_rgb(200, 40, 40))
                     .min_size(egui::vec2(22.0, 16.0))
                    ).on_hover_text("Stopp opplesing").clicked() {
                        tts::stop_speaking();
                        self.ocr_text = None;
                    }
                } else {
                    if ui.add(egui::Label::new(
                        egui::RichText::new("▶").size(14.0).color(inactive)
                    ).sense(egui::Sense::click())).on_hover_text("Les opp markert tekst").clicked() {
                        log!("Speak button clicked!");
                        match self.platform.read_selected_text() {
                            Some(text) => {
                                let trimmed = text.trim();
                                log!("Selected text: '{}'", &trimmed[..trimmed.len().min(80)]);
                                if !trimmed.is_empty() {
                                    tts::speak_word(trimmed);
                                }
                            }
                            None => {
                                log!("No selected text found");
                            }
                        }
                    }
                }


                // --- Error count (on Grammatikk tab) ---
                if self.selected_tab == 1 {
                    let err_count = self.writing_errors.iter().filter(|e| !e.ignored).count();
                    if err_count > 0 {
                        ui.add_space(12.0);
                        ui.label(egui::RichText::new("Tips:").size(9.0).color(egui::Color32::from_rgb(120, 120, 120)));
                        ui.label(egui::RichText::new(format!("{}", err_count)).size(12.0).strong().color(egui::Color32::from_rgb(180, 60, 60)));
                    }
                }

                // --- Right side: drag area, 📌 ⚙ ✕ ---
                let remaining = ui.available_rect_before_wrap();
                let right_w = 80.0; // ▁ + 📌 + ⚙ + ✕
                let drag_rect = egui::Rect::from_min_max(
                    remaining.min,
                    egui::pos2(remaining.max.x - right_w, remaining.max.y),
                );
                let drag_resp = ui.allocate_rect(drag_rect, egui::Sense::drag());
                if drag_resp.drag_started() && !self.follow_cursor {
                    ctx.send_viewport_cmd(egui::ViewportCommand::StartDrag);
                }

                // 📌 Follow cursor toggle
                let pin_color = if self.follow_cursor {
                    egui::Color32::from_rgb(0, 120, 60)
                } else {
                    egui::Color32::from_rgb(160, 160, 160)
                };
                let pin_tooltip = if self.follow_cursor { "Følg markør (på)" } else { "Følg markør (av)" };
                if ui.add(egui::Label::new(
                    egui::RichText::new("\u{1F4CC}").size(14.0).color(pin_color)
                ).sense(egui::Sense::click())).on_hover_text(pin_tooltip).clicked() {
                    self.follow_cursor = !self.follow_cursor;
                }

                // ⚙ Settings
                let settings_color = if self.selected_tab == 2 { active } else { inactive };
                if ui.add(egui::Label::new(
                    egui::RichText::new("\u{2699}").size(16.0).color(settings_color)
                ).sense(egui::Sense::click())).on_hover_text("Innstillinger").clicked() {
                    self.selected_tab = 2;
                }

                // ▁ Minimize
                if ui.add(egui::Label::new(
                    egui::RichText::new("–").size(14.0).color(inactive)
                ).sense(egui::Sense::click())).on_hover_text("Minimer").clicked() {
                    ctx.send_viewport_cmd(egui::ViewportCommand::Minimized(true));
                }

                // ✕ Close button
                let close_resp = ui.allocate_rect(
                    egui::Rect::from_min_size(ui.cursor().min, egui::vec2(18.0, 18.0)),
                    egui::Sense::click() | egui::Sense::hover(),
                );
                let color = if close_resp.hovered() {
                    egui::Color32::from_rgb(220, 50, 50)
                } else {
                    egui::Color32::from_rgb(120, 120, 120)
                };
                let center = close_resp.rect.center();
                let s = 4.5;
                let stroke = egui::Stroke::new(1.5, color);
                ui.painter().line_segment([center + egui::vec2(-s, -s), center + egui::vec2(s, s)], stroke);
                ui.painter().line_segment([center + egui::vec2(s, -s), center + egui::vec2(-s, s)], stroke);
                if close_resp.clicked() {
                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                }
            });

            ui.separator();

            // === Whisper transcription result — shown in separate centered window ===
            // (rendering happens below via show_viewport_immediate)

            // === Tab: Innhold (0) ===
            if self.selected_tab == 0 {
                if !self.completions.is_empty() || !self.open_completions.is_empty() {

                    let sel = self.selected_completion;
                    let mut clicked_word: Option<String> = None;
                    let has_dual = !self.open_completions.is_empty() && !self.completions.is_empty();
                    let has_right_only = !self.open_completions.is_empty() && self.completions.is_empty();

                    let has_tts = tts::tts_available();
                    let icon_w: f32 = if has_tts { 16.0 } else { 0.0 };
                    let render_row = |ui: &mut egui::Ui, comp: &Completion, _idx: usize, is_selected: bool, is_top: bool, col_width: f32| -> (bool, bool) {
                        let marker = if is_selected { "▸ " } else { "  " };
                        let text = format!("{}{}", marker, comp.word);
                        let row_h = if is_top || is_selected { 18.0 } else { 16.0 };

                        // Single allocation for the whole row
                        let (rect, resp) = ui.allocate_exact_size(
                            egui::vec2(col_width, row_h),
                            egui::Sense::click() | egui::Sense::hover(),
                        );
                        let hovered = resp.hovered();
                        if is_selected {
                            ui.painter().rect_filled(rect, 2.0, egui::Color32::from_rgb(0, 100, 180));
                        } else if hovered {
                            ui.painter().rect_filled(rect, 2.0, egui::Color32::from_rgb(220, 235, 250));
                        }

                        // Speaker icon at the left edge
                        let mut spoke = false;
                        if has_tts {
                            let icon_fg = if is_selected { egui::Color32::from_rgba_premultiplied(200, 200, 200, 255) }
                                else { egui::Color32::from_rgb(150, 150, 150) };
                            let icon_center = egui::pos2(rect.min.x + icon_w * 0.5, rect.center().y);
                            ui.painter().text(icon_center, egui::Align2::CENTER_CENTER, "🔊", egui::FontId::proportional(9.0), icon_fg);

                            // Check if click was in the icon area
                            if resp.clicked() {
                                if let Some(pos) = resp.interact_pointer_pos() {
                                    if pos.x < rect.min.x + icon_w {
                                        tts::speak_word(&comp.word);
                                        spoke = true;
                                    }
                                }
                            }
                        }

                        let fg = if is_selected { egui::Color32::WHITE }
                            else if hovered { egui::Color32::from_rgb(0, 80, 140) }
                            else if is_top { egui::Color32::from_rgb(0, 120, 60) }
                            else { egui::Color32::from_rgb(60, 60, 60) };
                        let font_size = if is_top || is_selected || hovered { 13.0 } else { 12.0 };
                        let text_x = rect.min.x + icon_w;
                        ui.painter().text(egui::pos2(text_x, rect.min.y + 1.0), egui::Align2::LEFT_TOP, text, egui::FontId::proportional(font_size), fg);
                        if hovered { ui.ctx().set_cursor_icon(egui::CursorIcon::PointingHand); }
                        (resp.clicked() && !spoke, spoke)
                    };

                    if has_dual {
                        let avail_w = ui.available_width();
                        let col_w = (avail_w - 4.0) / 2.0;
                        let max_rows = self.completions.len().max(self.open_completions.len());
                        for row in 0..max_rows {
                            ui.horizontal(|ui| {
                                if row < self.completions.len() {
                                    let comp = &self.completions[row];
                                    let is_sel = self.selected_column == 0 && sel == Some(row);
                                    let is_top = row == 0 && sel.is_none();
                                    let (clicked, _) = render_row(ui, comp, row, is_sel, is_top, col_w);
                                    if clicked { clicked_word = Some(comp.word.clone()); }
                                } else {
                                    ui.allocate_exact_size(egui::vec2(col_w, 16.0), egui::Sense::hover());
                                }
                                ui.add_space(4.0);
                                if row < self.open_completions.len() {
                                    let comp = &self.open_completions[row];
                                    let is_sel = self.selected_column == 1 && sel == Some(row);
                                    let (clicked, _) = render_row(ui, comp, row + 100, is_sel, false, col_w);
                                    if clicked { clicked_word = Some(comp.word.clone()); }
                                }
                            });
                        }
                    } else if has_right_only {
                        // No prefix typed — show right column (next-word predictions) as single list
                        let avail_w = ui.available_width();
                        for (i, comp) in self.open_completions.iter().enumerate() {
                            let is_sel = sel == Some(i);
                            let is_top = i == 0 && sel.is_none();
                            let (clicked, _) = render_row(ui, comp, i, is_sel, is_top, avail_w);
                            if clicked { clicked_word = Some(comp.word.clone()); }
                        }
                    } else {
                        let avail_w = ui.available_width();
                        for (i, comp) in self.completions.iter().enumerate() {
                            let is_sel = sel == Some(i);
                            let is_top = i == 0 && sel.is_none();
                            let (clicked, _) = render_row(ui, comp, i, is_sel, is_top, avail_w);
                            if clicked { clicked_word = Some(comp.word.clone()); }
                        }
                    }


                    if let Some(word) = clicked_word {
                        log!("CLICKED word: '{}' bridge={}", word, self.manager.active_bridge_name());
                        self.return_focus_to_app();
                        self.manager.replace_word(&word);
                        self.completions.clear();
                        self.open_completions.clear();
                        self.selected_completion = None;
                        self.selection_mode = false;
                        self.last_completed_prefix.clear();
                        self.last_poll = Instant::now() - self.poll_interval;
                    }
                } else {
                    ui.label(
                        egui::RichText::new("Flytt cursoren for å se forslag...")
                            .italics()
                            .size(11.0)
                            .color(egui::Color32::from_rgb(150, 150, 140)),
                    );
                }
            }

            // === Tab: Grammatikk (1) ===
            if self.selected_tab == 1 {
                // Show scanning indicator while grammar queue is draining
                if self.grammar_scanning && self.grammar_queue_total > 0 {
                    let done = self.grammar_queue_total - self.grammar_queue.len();
                    let progress = done as f32 / self.grammar_queue_total as f32;
                    ui.horizontal(|ui| {
                        ui.spinner();
                        ui.add(egui::ProgressBar::new(progress)
                            .text(format!("{}/{}", done, self.grammar_queue_total))
                            .desired_width(ui.available_width() - 30.0)
                            .desired_height(14.0));
                    });
                    ui.add_space(2.0);
                }

                let mut active_errors: Vec<usize> = self.writing_errors.iter()
                    .enumerate()
                    .filter(|(_, e)| !e.ignored)
                    .map(|(i, _)| i)
                    .collect();
                // Sort: focused error first, then SentenceBoundary, Grammar, Spelling
                let focused = self.focused_error_idx;
                active_errors.sort_by_key(|&i| {
                    let is_focused = focused == Some(i);
                    let is_pinned = self.writing_errors[i].pinned;
                    // Focused/pinned → 0, rest sorted by category
                    if is_focused || is_pinned { (0, 0) }
                    else {
                        let cat = match self.writing_errors[i].category {
                            ErrorCategory::SentenceBoundary => 0,
                            ErrorCategory::Grammar => 1,
                            ErrorCategory::Spelling => 2,
                        };
                        (1, cat)
                    }
                });

                if active_errors.is_empty() {
                    ui.label(
                        egui::RichText::new("Ingen feil funnet.")
                            .size(12.0)
                            .color(egui::Color32::from_rgb(0, 140, 60)),
                    );
                } else {

                    let mut action: Option<(usize, &str)> = None;

                    // Group grammar errors by (sentence_context, doc_offset)
                    let mut shown_contexts: std::collections::HashSet<(String, usize)> = std::collections::HashSet::new();

                    egui::ScrollArea::vertical().max_height(ui.available_height() - 4.0).show(ui, |ui| {
                    for &idx in &active_errors {
                        let error = &self.writing_errors[idx];

                        // For grammar errors with position > 0, skip — they're shown as alternatives
                        // (but never skip the focused error — it must render for yellow highlight)
                        if matches!(error.category, ErrorCategory::Grammar) && error.position > 0
                            && self.focused_error_idx != Some(idx) && !error.pinned
                        {
                            if shown_contexts.contains(&(error.sentence_context.clone(), error.doc_offset)) {
                                continue;
                            }
                        }

                        ui.separator();
                        // Highlight and scroll to the focused error (cursor on underlined word)
                        let is_focused = self.focused_error_idx == Some(idx) || error.pinned;
                        let frame = if is_focused {
                            egui::Frame::NONE.fill(egui::Color32::from_rgba_premultiplied(255, 255, 180, 255))
                                .inner_margin(4.0).corner_radius(4.0)
                        } else {
                            egui::Frame::NONE
                        };
                        let frame_resp = frame.show(ui, |ui| {
                            if matches!(error.category, ErrorCategory::SentenceBoundary) {
                                // --- Sentence boundary suggestion ---
                                let err_suggestion = error.suggestion.clone();
                                ui.horizontal(|ui| {
                                    if icon_button(ui, "👍", "Sett inn punktum") {
                                        action = Some((idx, "fix"));
                                    }
                                    if icon_button(ui, "👎", "Ignorer") {
                                        action = Some((idx, "ignore"));
                                    }
                                    if icon_button(ui, "🔊", "Les opp") {
                                        tts::speak_word(&err_suggestion);
                                    }
                                    if icon_button(ui, "▶", "Vis i dokument") {
                                        action = Some((idx, "goto"));
                                    }
                                });
                                ui.label(
                                    egui::RichText::new("Mangler punktum:")
                                        .size(11.0)
                                        .strong()
                                        .color(egui::Color32::from_rgb(0, 100, 180)),
                                );
                                // Show the suggested punctuated version
                                ui.label(
                                    egui::RichText::new(&error.suggestion)
                                        .size(11.0)
                                        .color(egui::Color32::from_rgb(0, 120, 60)),
                                );
                                ui.label(
                                    egui::RichText::new(&error.explanation)
                                        .size(10.0)
                                        .color(egui::Color32::from_rgb(100, 100, 100)),
                                );
                            } else if matches!(error.category, ErrorCategory::Grammar) {
                                shown_contexts.insert((error.sentence_context.clone(), error.doc_offset));
                                // Show all alternatives for this sentence occurrence
                                let ctx = error.sentence_context.clone();
                                let ctx_offset = error.doc_offset;
                                let alternatives: Vec<usize> = active_errors.iter()
                                    .filter(|&&i| {
                                        let e = &self.writing_errors[i];
                                        matches!(e.category, ErrorCategory::Grammar)
                                            && e.sentence_context == ctx
                                            && e.doc_offset == ctx_offset
                                            && !e.suggestion.is_empty()
                                    })
                                    .copied()
                                    .collect();

                                // Buttons on top line
                                let first_alt = alternatives.first().copied();
                                let first_suggestion = first_alt.map(|i| self.writing_errors[i].suggestion.clone()).unwrap_or_default();
                                let err_rule = error.rule_name.clone();
                                let err_expl = error.explanation.clone();
                                let err_ctx = error.sentence_context.clone();
                                ui.horizontal(|ui| {
                                    if let Some(alt_idx) = first_alt {
                                        if icon_button(ui, "👍", "Rett opp") {
                                            action = Some((alt_idx, "fix"));
                                        }
                                    }
                                    if icon_button(ui, "👎", "Ignorer") {
                                        action = Some((idx, "ignore_group"));
                                    }
                                    if icon_button(ui, "🔊", "Les opp") {
                                        tts::speak_word(&first_suggestion);
                                    }
                                    if icon_button(ui, "💡", "Vis regelinfo") {
                                        let fix_idx = first_alt.unwrap_or(idx);
                                        self.rule_info_window = Some((err_rule.clone(), err_expl.clone(), err_ctx.clone(), fix_idx, first_suggestion.clone()));
                                    }
                                    if icon_button(ui, "▶", "Vis i dokument") {
                                        action = Some((idx, "goto"));
                                    }
                                });
                                // Original (red, strikethrough) then suggestion (green) — stacked
                                ui.label(
                                    egui::RichText::new(&error.word)
                                        .size(11.0)
                                        .strikethrough()
                                        .color(egui::Color32::from_rgb(180, 60, 60)),
                                );
                                for &alt_idx in &alternatives {
                                    let alt = &self.writing_errors[alt_idx];
                                    ui.label(
                                        egui::RichText::new(&alt.suggestion)
                                            .size(11.0)
                                            .strong()
                                            .color(egui::Color32::from_rgb(0, 120, 60)),
                                    );
                                }
                                // Explanation
                                if let Some(alt_idx) = first_alt {
                                    ui.label(
                                        egui::RichText::new(&self.writing_errors[alt_idx].explanation)
                                            .size(10.0)
                                            .color(egui::Color32::from_rgb(100, 100, 100)),
                                    );
                                }
                            } else {
                                // Spelling error — buttons on top, then word/suggestion stacked
                                let err_suggestion = error.suggestion.clone();
                                let err_word = error.word.clone();
                                ui.horizontal(|ui| {
                                    if !error.suggestion.is_empty() {
                                        if icon_button(ui, "👍", "Rett opp") {
                                            action = Some((idx, "fix"));
                                        }
                                    }
                                    if icon_button(ui, "👎", "Ignorer") {
                                        action = Some((idx, "ignore"));
                                    }
                                    if icon_button(ui, "🔊", "Les opp") {
                                        let speak = if !err_suggestion.is_empty() { &err_suggestion } else { &err_word };
                                        tts::speak_word(speak);
                                    }
                                    if icon_button(ui, "?", "Flere forslag") {
                                        action = Some((idx, "suggest"));
                                    }
                                    if icon_button(ui, "▶", "Vis i dokument") {
                                        action = Some((idx, "goto"));
                                    }
                                });
                                ui.label(
                                    egui::RichText::new(&error.word)
                                        .size(12.0)
                                        .strong()
                                        .color(egui::Color32::from_rgb(200, 40, 40)),
                                );
                                if !error.suggestion.is_empty() {
                                    ui.label(
                                        egui::RichText::new(&error.suggestion)
                                            .size(12.0)
                                            .strong()
                                            .color(egui::Color32::from_rgb(0, 120, 60)),
                                    );
                                }
                                ui.label(
                                    egui::RichText::new(&error.explanation)
                                        .size(10.0)
                                        .color(egui::Color32::from_rgb(80, 80, 80)),
                                );
                            }
                        });
                        if is_focused && !self.focused_error_scroll_done {
                            frame_resp.response.scroll_to_me(Some(egui::Align::Center));
                            self.focused_error_scroll_done = true;
                        }
                    }
                    }); // end ScrollArea

                    // Handle actions after rendering
                    if let Some((idx, act)) = action {
                        log!("ACTION received: act='{}' idx={}", act, idx);
                        // Clear pin when user acts on any error
                        if matches!(act, "fix" | "ignore" | "ignore_group") {
                            self.writing_errors[idx].pinned = false;
                        }
                        match act {
                            "fix" => {
                                let error = &self.writing_errors[idx];
                                let suggestion = error.suggestion.clone();
                                let word = error.word.clone();
                                let context = error.sentence_context.clone();
                                let off = error.doc_offset;
                                self.pending_fix = Some((word.clone(), suggestion.clone(), context, off));
                                log!("FIX action: idx={} bridge='{}' word='{}' suggestion='{}'",
                                    idx, self.manager.active_bridge_name(),
                                    trunc(&word, 60), trunc(&suggestion, 60));
                                // Mark all alternatives for this sentence occurrence as ignored
                                let ctx = self.writing_errors[idx].sentence_context.clone();
                                let ctx_off = self.writing_errors[idx].doc_offset;
                                for e in &mut self.writing_errors {
                                    if e.sentence_context == ctx && e.doc_offset == ctx_off
                                        && (matches!(e.category, ErrorCategory::Grammar)
                                            || matches!(e.category, ErrorCategory::SentenceBoundary))
                                    {
                                        e.ignored = true;
                                    }
                                }
                            }
                            "suggest" => {
                                let word = self.writing_errors[idx].word.clone();
                                let sentence_ctx = self.writing_errors[idx].sentence_context.clone();
                                // Use unified pipeline — all suggestions are grammar-verified
                                let suggestions = self.find_spelling_suggestions(&word, &sentence_ctx);
                                self.suggestion_window = Some((word, suggestions));
                            }
                            "ignore" => {
                                let error = &self.writing_errors[idx];
                                log!("ACTION ignore: word='{}' rule='{}'", error.word, error.rule_name);
                                if matches!(error.category, ErrorCategory::Spelling) {
                                    self.ignored_words.insert(error.word.clone());
                                }
                                self.writing_errors[idx].ignored = true;
                            }
                            "ignore_group" => {
                                let ctx = self.writing_errors[idx].sentence_context.clone();
                                let ctx_off = self.writing_errors[idx].doc_offset;
                                for e in &mut self.writing_errors {
                                    if e.sentence_context == ctx && e.doc_offset == ctx_off && matches!(e.category, ErrorCategory::Grammar) {
                                        e.ignored = true;
                                    }
                                }
                            }
                            "goto" => {
                                let error = &self.writing_errors[idx];
                                let start = error.doc_offset;
                                let end = start + error.sentence_context.chars().count();
                                log!("GOTO: selecting range {}..{} for '{}'", start, end,
                                    trunc(&error.sentence_context, 50));
                                match self.manager.select_range(start, end) {
                                    Some((x, y)) => {
                                        log!("GOTO: select_range returned ({}, {})", x, y);
                                        if x != 0 || y != 0 {
                                            self.last_caret_pos = Some((x, y));
                                            self.follow_cursor = true;
                                            // Freeze cursor-follow for 5s so window doesn't jump back
                                            self.goto_freeze_until = Some(Instant::now() + Duration::from_secs(5));
                                            // Force-move window to the selection position
                                            let (screen_w, screen_h) = get_screen_size(&*self.platform);
                                            let win_h = 300.0_f32;
                                            let win_w = 350.0_f32;
                                            let pos_y = if (y as f32 + win_h) > screen_h {
                                                y as f32 - win_h - 30.0
                                            } else {
                                                y as f32
                                            };
                                            let pos_x = (x as f32).min(screen_w - win_w).max(0.0);
                                            ctx.send_viewport_cmd(egui::ViewportCommand::OuterPosition(
                                                egui::pos2(pos_x, pos_y),
                                            ));
                                            log!("GOTO: moved window to ({}, {})", pos_x, pos_y);
                                        }
                                    }
                                    None => {
                                        log!("GOTO: select_range returned None");
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                }
            }

            // === Suggestion window — separate OS window, centered ===
            if self.suggestion_window.is_some() {
                // Check if a selection was made in a previous frame (via Arc<Mutex>)
                let prev_selection = self.suggestion_selection.lock().unwrap().take();
                let (word, candidates) = self.suggestion_window.as_ref().unwrap();
                let word_clone = word.clone();
                let candidates_clone: Vec<(String, f32)> = candidates.clone();

                if let Some(idx) = prev_selection {
                    if idx < candidates_clone.len() {
                        let replacement = candidates_clone[idx].0.clone();
                        log!("Suggestion selected: '{}' → '{}'", word_clone, replacement);
                        let (sent_ctx, sent_off) = self.writing_errors.iter()
                            .find(|e| e.word == word_clone && !e.ignored)
                            .map(|e| (e.sentence_context.clone(), e.doc_offset))
                            .unwrap_or_default();
                        self.pending_fix = Some((word_clone.clone(), replacement.clone(), sent_ctx, sent_off));
                        for e in &mut self.writing_errors {
                            if e.word == word_clone && !e.ignored {
                                e.suggestion = replacement;
                                e.ignored = true;
                                break;
                            }
                        }
                        self.suggestion_window = None;
                    }
                } else {
                    let mut do_close = false;
                    let selection = self.suggestion_selection.clone();

                    let win_w = 320.0_f32;
                    let win_h = 340.0_f32;
                    let monitor = ctx.input(|i| i.viewport().monitor_size.unwrap_or(egui::vec2(1920.0, 1080.0)));
                    let screen_center = egui::pos2(
                        (monitor.x - win_w) / 2.0,
                        (monitor.y - win_h) / 2.0,
                    );

                    ctx.show_viewport_immediate(
                        egui::ViewportId::from_hash_of("suggestion_viewport"),
                        egui::ViewportBuilder::default()
                            .with_title(format!("Forslag for «{}»", word_clone))
                            .with_inner_size([win_w, win_h])
                            .with_position(screen_center)
                            .with_always_on_top()
                            .with_decorations(true),
                        |vp_ctx, _class| {
                            vp_ctx.set_visuals(egui::Visuals::light());

                            if vp_ctx.input(|i| i.viewport().close_requested()) {
                                do_close = true;
                            }

                            egui::CentralPanel::default()
                                .frame(
                                    egui::Frame::new()
                                        .fill(egui::Color32::WHITE)
                                        .inner_margin(16.0),
                                )
                                .show(vp_ctx, |ui| {
                                    ui.visuals_mut().override_text_color = Some(egui::Color32::from_rgb(30, 30, 30));

                                    ui.label(
                                        egui::RichText::new(format!("Forslag for «{}»", word_clone))
                                            .size(16.0)
                                            .strong()
                                            .color(egui::Color32::from_rgb(30, 70, 150)),
                                    );
                                    ui.add_space(8.0);

                                    if candidates_clone.is_empty() {
                                        ui.label("Ingen forslag funnet.");
                                    } else {
                                        egui::ScrollArea::vertical().max_height(win_h - 80.0).show(ui, |ui| {
                                            for (i, (candidate, _score)) in candidates_clone.iter().enumerate() {
                                                ui.horizontal(|ui| {
                                                    if icon_button(ui, "🔊", "Les opp") {
                                                        tts::speak_word(candidate);
                                                    }
                                                    if ui.button(
                                                        egui::RichText::new(candidate).size(14.0).strong()
                                                    ).clicked() {
                                                        *selection.lock().unwrap() = Some(i);
                                                    }
                                                });
                                            }
                                        });
                                    }
                                });
                        },
                    );

                    if do_close {
                        self.suggestion_window = None;
                    }
                }
            }

            // === Rule info window — separate OS window ===
            if self.rule_info_window.is_some() {
                let mut do_fix = false;
                let mut do_ignore = false;
                let mut do_close = false;
                let (rule_name, explanation, sentence, fix_idx, suggestion) = self.rule_info_window.as_ref().unwrap();
                let rule_name = rule_name.clone();
                let explanation = explanation.clone();
                let sentence = sentence.clone();
                let fix_idx = *fix_idx;
                let suggestion = suggestion.clone();
                let error_word = self.writing_errors[fix_idx].word.clone();
                let corrected_sentence = if !suggestion.is_empty() {
                    sentence.replacen(&error_word, &suggestion, 1)
                } else {
                    String::new()
                };
                let (category, description, wrong, right) = rule_info(&rule_name);

                // Center on screen using actual monitor size
                let win_w = 560.0_f32;
                let win_h = 520.0_f32;
                let monitor = ctx.input(|i| i.viewport().monitor_size.unwrap_or(egui::vec2(1920.0, 1080.0)));
                let screen_center = egui::pos2(
                    (monitor.x - win_w) / 2.0,
                    (monitor.y - win_h) / 2.0,
                );

                ctx.show_viewport_immediate(
                    egui::ViewportId::from_hash_of("rule_info_viewport"),
                    egui::ViewportBuilder::default()
                        .with_title("Regelinfo")
                        .with_inner_size([win_w, win_h])
                        .with_position(screen_center)
                        .with_always_on_top()
                        .with_decorations(true),
                    |vp_ctx, _class| {
                        // Switch to light visuals for this viewport
                        vp_ctx.set_visuals(egui::Visuals::light());

                        egui::CentralPanel::default()
                            .frame(
                                egui::Frame::new()
                                    .fill(egui::Color32::WHITE)
                                    .inner_margin(24.0),
                            )
                            .show(vp_ctx, |ui| {
                                ui.visuals_mut().override_text_color = Some(egui::Color32::from_rgb(30, 30, 30));

                                // Wrap text for long sentences
                                let max_w = ui.available_width();
                                ui.set_max_width(max_w);

                                // Scrollable content area (everything except buttons)
                                let scroll_height = ui.available_height() - 50.0;
                                egui::ScrollArea::vertical().max_height(scroll_height).show(ui, |ui| {
                                    ui.set_max_width(max_w - 16.0);

                                    // Category header
                                    ui.label(
                                        egui::RichText::new(category)
                                            .size(22.0)
                                            .strong()
                                            .color(egui::Color32::from_rgb(30, 70, 150)),
                                    );
                                    ui.add_space(10.0);

                                    // Description
                                    ui.label(
                                        egui::RichText::new(description)
                                            .size(15.0)
                                            .color(egui::Color32::from_rgb(30, 30, 30)),
                                    );
                                    ui.add_space(14.0);

                                    // Original sentence (red) and corrected (green)
                                    egui::Frame::new()
                                        .fill(egui::Color32::from_rgb(255, 245, 245))
                                        .inner_margin(10.0)
                                        .corner_radius(6.0)
                                        .stroke(egui::Stroke::new(1.0, egui::Color32::from_rgb(220, 180, 180)))
                                        .show(ui, |ui| {
                                            ui.set_max_width(max_w - 40.0);
                                            ui.label(
                                                egui::RichText::new(&sentence)
                                                    .size(15.0)
                                                    .strikethrough()
                                                    .color(egui::Color32::from_rgb(180, 50, 50)),
                                            );
                                        });
                                    ui.add_space(4.0);
                                    if !corrected_sentence.is_empty() {
                                        egui::Frame::new()
                                            .fill(egui::Color32::from_rgb(240, 255, 245))
                                            .inner_margin(10.0)
                                            .corner_radius(6.0)
                                            .stroke(egui::Stroke::new(1.0, egui::Color32::from_rgb(180, 220, 190)))
                                            .show(ui, |ui| {
                                                ui.set_max_width(max_w - 40.0);
                                                ui.label(
                                                    egui::RichText::new(&corrected_sentence)
                                                        .size(15.0)
                                                        .color(egui::Color32::from_rgb(0, 120, 50)),
                                                );
                                            });
                                    }
                                    ui.add_space(12.0);

                                    // Explanation
                                    ui.label(
                                        egui::RichText::new("Forklaring:")
                                            .size(14.0)
                                            .strong()
                                            .color(egui::Color32::from_rgb(50, 50, 50)),
                                    );
                                    ui.add_space(4.0);
                                    ui.label(
                                        egui::RichText::new(&explanation)
                                            .size(14.0)
                                            .color(egui::Color32::from_rgb(30, 30, 30)),
                                    );

                                    // Examples
                                    if !wrong.is_empty() {
                                        ui.add_space(14.0);
                                        ui.separator();
                                        ui.add_space(8.0);
                                        ui.label(
                                            egui::RichText::new("Eksempler")
                                                .size(18.0)
                                                .strong()
                                                .color(egui::Color32::from_rgb(30, 70, 150)),
                                        );
                                        ui.add_space(8.0);

                                        for (w, r) in wrong.iter().zip(right.iter()) {
                                            ui.horizontal(|ui| {
                                                ui.label(egui::RichText::new("X").size(15.0).strong().color(egui::Color32::from_rgb(200, 40, 40)));
                                                ui.label(egui::RichText::new(*w).size(15.0).strikethrough().color(egui::Color32::from_rgb(160, 70, 70)));
                                            });
                                            ui.horizontal(|ui| {
                                                ui.label(egui::RichText::new("V").size(15.0).strong().color(egui::Color32::from_rgb(0, 140, 60)));
                                                ui.label(egui::RichText::new(*r).size(15.0).color(egui::Color32::from_rgb(0, 100, 40)));
                                            });
                                            ui.add_space(5.0);
                                        }
                                    }
                                });

                                // Action buttons — always visible at bottom
                                ui.separator();
                                ui.add_space(4.0);
                                ui.horizontal(|ui| {
                                    if !suggestion.is_empty() {
                                        if ui.button(egui::RichText::new("Rett opp").size(14.0).strong().color(egui::Color32::from_rgb(0, 120, 60))).clicked() {
                                            do_fix = true;
                                        }
                                        ui.add_space(8.0);
                                    }
                                    if ui.button(egui::RichText::new("Ignorer").size(14.0).color(egui::Color32::from_rgb(150, 60, 60))).clicked() {
                                        do_ignore = true;
                                    }
                                    ui.add_space(8.0);
                                    if ui.button(egui::RichText::new("Lukk").size(14.0).color(egui::Color32::from_rgb(80, 80, 80))).clicked() {
                                        do_close = true;
                                    }
                                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                                        ui.label(egui::RichText::new(format!("Regel: {}", rule_name)).size(11.0).color(egui::Color32::from_rgb(160, 160, 160)));
                                    });
                                });
                            });

                        // Close viewport when user clicks X on title bar
                        if vp_ctx.input(|i| i.viewport().close_requested()) {
                            do_close = true;
                        }
                    },
                );

                if do_fix {
                    let error = &self.writing_errors[fix_idx];
                    let s = error.suggestion.clone();
                    let w = error.word.clone();
                    let c = error.sentence_context.clone();
                    let o = error.doc_offset;
                    self.pending_fix = Some((w, s, c, o));
                    let sctx = self.writing_errors[fix_idx].sentence_context.clone();
                    let soff = self.writing_errors[fix_idx].doc_offset;
                    for e in &mut self.writing_errors {
                        if e.sentence_context == sctx && e.doc_offset == soff && matches!(e.category, ErrorCategory::Grammar) {
                            e.ignored = true;
                        }
                    }
                    self.rule_info_window = None;
                } else if do_ignore {
                    let sctx = self.writing_errors[fix_idx].sentence_context.clone();
                    let soff = self.writing_errors[fix_idx].doc_offset;
                    for e in &mut self.writing_errors {
                        if e.sentence_context == sctx && e.doc_offset == soff && matches!(e.category, ErrorCategory::Grammar) {
                            e.ignored = true;
                        }
                    }
                    self.rule_info_window = None;
                } else if do_close {
                    self.rule_info_window = None;
                }
            }

            // === OCR: screenshot detected — handled in separate window below ===

            // === Whisper transcription popup (centered window) ===
            let is_recording = stt::is_recording();
            let is_correcting = self.mic_transcribing && !is_recording;
            let is_streaming = is_recording || self.mic_transcribing;
            let show_whisper_popup = self.mic_result_text.is_some() || is_correcting || is_recording || self.whisper_loading;
            if show_whisper_popup {
                let mut do_close = false;
                let mut do_copy = false;
                let mut do_stop = false;
                let text_clone = self.mic_result_text.clone().unwrap_or_default();

                let win_w = 600.0_f32;
                let win_h = 400.0_f32;
                let monitor = ctx.input(|i| i.viewport().monitor_size.unwrap_or(egui::vec2(1920.0, 1080.0)));
                let screen_center = egui::pos2(
                    (monitor.x - win_w) / 2.0,
                    (monitor.y - win_h) / 2.0,
                );

                ctx.show_viewport_immediate(
                    egui::ViewportId::from_hash_of("whisper_result_viewport"),
                    egui::ViewportBuilder::default()
                        .with_title("Talegjenkjenning")
                        .with_inner_size([win_w, win_h])
                        .with_position(screen_center)
                        .with_always_on_top()
                        .with_decorations(true),
                    |vp_ctx, _class| {
                        vp_ctx.set_visuals(egui::Visuals::light());

                        egui::CentralPanel::default()
                            .frame(
                                egui::Frame::new()
                                    .fill(egui::Color32::WHITE)
                                    .inner_margin(20.0),
                            )
                            .show(vp_ctx, |ui| {
                                ui.visuals_mut().override_text_color = Some(egui::Color32::from_rgb(30, 30, 30));
                                let max_w = ui.available_width();
                                ui.set_max_width(max_w);

                                // Status bar: loading / stop button (recording) / correcting spinner
                                if self.whisper_loading && !is_recording {
                                    ui.horizontal(|ui| {
                                        ui.spinner();
                                        let msg = if self.whisper_load_status.is_empty() {
                                            "Laster talemodell...".to_string()
                                        } else {
                                            self.whisper_load_status.clone()
                                        };
                                        ui.label(egui::RichText::new(msg).size(14.0)
                                            .color(egui::Color32::from_rgb(80, 80, 140)));
                                    });
                                    ui.add_space(8.0);
                                } else if is_recording {
                                    ui.horizontal(|ui| {
                                        if ui.add(egui::Button::new(
                                            egui::RichText::new("■ Stopp").size(14.0).color(egui::Color32::WHITE)
                                        ).fill(egui::Color32::from_rgb(200, 40, 40))).clicked() {
                                            do_stop = true;
                                        }
                                    });
                                    ui.add_space(8.0);
                                } else if is_correcting {
                                    ui.horizontal(|ui| {
                                        ui.spinner();
                                        let msg = if self.whisper_mode == 1 { "Forbedrer med stor modell..." } else { "Transkriberer..." };
                                        ui.label(egui::RichText::new(msg).size(14.0)
                                            .color(egui::Color32::from_rgb(100, 80, 140)));
                                    });
                                    ui.add_space(8.0);
                                }

                                let btn_space = if is_streaming { 10.0 } else { 50.0 };
                                let scroll_height = ui.available_height() - btn_space;
                                egui::ScrollArea::vertical().max_height(scroll_height).show(ui, |ui| {
                                    ui.set_max_width(max_w - 16.0);
                                    ui.horizontal_wrapped(|ui| {
                                        ui.label(egui::RichText::new(&text_clone).size(20.0)
                                            .color(egui::Color32::from_rgb(30, 30, 30)));
                                        if is_recording {
                                            ui.label(egui::RichText::new(" ...").size(20.0)
                                                .color(egui::Color32::from_rgb(150, 150, 150)));
                                        }
                                    });
                                });

                                ui.add_space(8.0);
                                ui.horizontal(|ui| {
                                    if !is_streaming {
                                        if ui.button(egui::RichText::new("Kopier").size(14.0)).clicked() {
                                            do_copy = true;
                                        }
                                        ui.add_space(8.0);
                                        if ui.button(egui::RichText::new("\u{1F50A} Les opp").size(14.0)).clicked() {
                                            tts::speak_word(&text_clone);
                                        }
                                        ui.add_space(8.0);
                                    }
                                    if ui.button(egui::RichText::new("Lukk").size(14.0)).clicked() {
                                        do_close = true;
                                    }
                                });
                            });

                        if vp_ctx.input(|i| i.viewport().close_requested()) {
                            do_close = true;
                            // Prevent close from propagating to main app
                            vp_ctx.send_viewport_cmd(egui::ViewportCommand::CancelClose);
                        }
                        vp_ctx.request_repaint_after(Duration::from_millis(100));
                    },
                );

                if do_stop {
                    if let Some(handle) = &self.mic_handle {
                        handle.stop();
                        self.mic_transcribing = true;
                    }
                }
                if do_copy {
                    self.platform.copy_to_clipboard(&text_clone);
                }
                if do_close {
                    // Stop recording if active
                    if let Some(handle) = self.mic_handle.take() {
                        handle.stop();
                    }
                    // Force-clear recording flag so popup closes immediately
                    // (background thread may still be transcribing — it will finish silently)
                    stt::force_stop();
                    self.mic_transcribing = false;
                    self.mic_result_text = None;
                    self.whisper_loading = false;
                    self.whisper_pending_record = false;
                }
            }

            // === Tab: Innstillinger (2) ===
            if self.selected_tab == 2 {
                ui.label(
                    egui::RichText::new(format!("Bro: {}", self.manager.active_bridge_name()))
                        .size(12.0)
                        .color(egui::Color32::from_rgb(100, 100, 100)),
                );
                ui.label(
                    egui::RichText::new(format!("Kvalitet: {} (0=rask, 1=balansert, 2=full)", self.quality))
                        .size(12.0)
                        .color(egui::Color32::from_rgb(100, 100, 100)),
                );
                if self.grammar_completion {
                    ui.label(
                        egui::RichText::new("Grammatikkfilter: PÅ")
                            .size(12.0)
                            .color(egui::Color32::from_rgb(0, 120, 60)),
                    );
                }
                ui.add_space(6.0);
                ui.label(egui::RichText::new("Talegjenkjenning:").size(12.0).color(egui::Color32::from_rgb(80, 80, 80)));
                ui.horizontal(|ui| {
                    let rask_color = if self.whisper_mode == 0 {
                        egui::Color32::from_rgb(0, 70, 160)
                    } else {
                        egui::Color32::from_rgb(100, 100, 100)
                    };
                    let beste_color = if self.whisper_mode == 1 {
                        egui::Color32::from_rgb(0, 70, 160)
                    } else {
                        egui::Color32::from_rgb(100, 100, 100)
                    };
                    if ui.add(egui::Label::new(
                        egui::RichText::new("Rask (75 MB)").size(12.0).color(rask_color)
                    ).sense(egui::Sense::click())).clicked() {
                        if self.whisper_mode != 0 {
                            self.whisper_mode = 0;
                            // Unload existing models to free memory
                            self.whisper_engine = None;
                            self.whisper_streaming = None;
                            log!("Whisper mode: Rask (tiny)");
                        }
                    }
                    ui.label(egui::RichText::new(" | ").size(12.0).color(egui::Color32::from_rgb(160, 160, 160)));
                    if ui.add(egui::Label::new(
                        egui::RichText::new("Beste (650 MB)").size(12.0).color(beste_color)
                    ).sense(egui::Sense::click())).clicked() {
                        if self.whisper_mode != 1 {
                            self.whisper_mode = 1;
                            // Unload existing models to free memory
                            self.whisper_engine = None;
                            self.whisper_streaming = None;
                            log!("Whisper mode: Beste (base+medium-q5)");
                        }
                    }
                });
                // Load errors
                for err in &self.load_errors {
                    ui.label(
                        egui::RichText::new(err)
                            .size(10.0)
                            .color(egui::Color32::from_rgb(200, 50, 50)),
                    );
                }

                // Voice selection
                ui.add_space(6.0);
                ui.horizontal(|ui| {
                    ui.label(egui::RichText::new("Stemme:").size(12.0).color(egui::Color32::from_rgb(80, 80, 80)));
                    let current = tts::current_voice();
                    ui.label(egui::RichText::new(&current).size(12.0).color(egui::Color32::from_rgb(0, 70, 160)));
                    if ui.add(egui::Button::new(
                        egui::RichText::new("Velg...").size(11.0)
                    ).small()).clicked() {
                        self.voice_list = tts::available_voices();
                        self.show_voice_window = true;
                    }
                });

                // UI scale
                ui.add_space(6.0);
                ui.horizontal(|ui| {
                    ui.label(egui::RichText::new("Størrelse:").size(12.0).color(egui::Color32::from_rgb(80, 80, 80)));
                    if ui.small_button("−").clicked() {
                        self.ui_scale = (self.ui_scale - 0.1).max(0.5);
                    }
                    ui.label(egui::RichText::new(format!("{:.0}%", self.ui_scale * 100.0)).size(12.0)
                        .color(egui::Color32::from_rgb(0, 70, 160)));
                    if ui.small_button("+").clicked() {
                        self.ui_scale = (self.ui_scale + 0.1).min(2.5);
                    }
                });

            }

            // === Tab: Debug (3) ===
            if self.selected_tab == 3 {
                let grey = egui::Color32::from_rgb(100, 100, 100);
                let dark = egui::Color32::from_rgb(50, 50, 50);
                ui.horizontal(|ui| {
                    ui.label(egui::RichText::new("Bro:").size(11.0).strong().color(grey));
                    ui.label(egui::RichText::new(self.manager.active_bridge_name()).size(11.0).color(dark));
                    ui.add_space(12.0);
                    ui.label(egui::RichText::new("Ord:").size(11.0).strong().color(grey));
                    ui.label(
                        egui::RichText::new(if self.context.word.is_empty() { "(tomt)" } else { &self.context.word })
                            .size(13.0)
                            .color(egui::Color32::from_rgb(0, 70, 160)),
                    );
                });
                ui.add_space(2.0);
                ui.label(egui::RichText::new("Setning:").size(11.0).strong().color(grey));
                ui.label(
                    egui::RichText::new(if self.context.sentence.is_empty() { "(tom)" } else { &self.context.sentence })
                        .size(11.0)
                        .color(dark),
                );
                ui.add_space(2.0);
                ui.label(egui::RichText::new("Maskert:").size(11.0).strong().color(grey));
                let masked_text = self.context.masked_sentence.clone().unwrap_or_else(|| "(ingen)".to_string());
                egui::ScrollArea::vertical().max_height(80.0).show(ui, |ui| {
                    ui.label(
                        egui::RichText::new(&masked_text)
                            .size(10.0)
                            .color(egui::Color32::from_rgb(80, 80, 80)),
                    );
                });
                ui.add_space(4.0);
                if ui.small_button("Kopier til utklippstavle").clicked() {
                    let mut text = format!("Bro: {}\nOrd: {}\nSetning: {}", self.manager.active_bridge_name(), self.context.word, self.context.sentence);
                    if let Some(masked) = &self.context.masked_sentence {
                        text.push_str(&format!("\nMaskert: {}", masked));
                    }
                    ctx.copy_text(text);
                }
            }
        });

        // Voice selection window (separate from main panel)
        if self.show_voice_window {
            let mut open = self.show_voice_window;
            egui::Window::new("Velg stemme")
                .open(&mut open)
                .resizable(true)
                .default_width(300.0)
                .show(ctx, |ui| {
                    let current = tts::current_voice();
                    if self.voice_list.is_empty() {
                        ui.label(egui::RichText::new("Ingen norske stemmer funnet.").size(12.0));
                        ui.label(egui::RichText::new("Last ned stemmer i Systeminnstillinger > Tilgjengelighet > Opplest innhold > Systemstemme").size(11.0)
                            .color(egui::Color32::from_rgb(100, 100, 100)));
                    }
                    for voice in &self.voice_list {
                        ui.horizontal(|ui| {
                            let is_selected = voice.name == current;
                            let color = if is_selected {
                                egui::Color32::from_rgb(0, 70, 160)
                            } else {
                                egui::Color32::from_rgb(60, 60, 60)
                            };
                            let label = if is_selected {
                                format!("{} (valgt)", &voice.name)
                            } else {
                                voice.name.clone()
                            };
                            if ui.add(egui::Label::new(
                                egui::RichText::new(&label).size(13.0).color(color)
                            ).sense(egui::Sense::click())).clicked() {
                                tts::set_voice(&voice.name);
                                tts::speak_word(&voice.sample_text);
                            }
                            ui.label(egui::RichText::new(&voice.language).size(10.0)
                                .color(egui::Color32::from_rgb(140, 140, 140)));
                        });
                    }
                });
            self.show_voice_window = open;
        }

        // OCR: screenshot detected prompt (separate OS window)
        let ocr_has_pending = self.ocr.as_ref().map_or(false, |o| o.has_pending_image());
        let ocr_is_busy = self.ocr_receiver.is_some();
        if ocr_has_pending && !ocr_is_busy {
            let mut do_read = false;
            let mut do_copy = false;
            let mut do_dismiss = false;

            let monitor = ctx.input(|i| i.viewport().monitor_size.unwrap_or(egui::vec2(1920.0, 1080.0)));
            let win_w: f32 = 320.0;
            let win_h: f32 = 100.0;
            let screen_center = egui::pos2(
                (monitor.x - win_w) / 2.0,
                (monitor.y - win_h) / 2.0,
            );

            ctx.show_viewport_immediate(
                egui::ViewportId::from_hash_of("ocr_prompt"),
                egui::ViewportBuilder::default()
                    .with_title("Skjermbilde oppdaget")
                    .with_inner_size([win_w, win_h])
                    .with_position(screen_center)
                    .with_always_on_top()
                    .with_decorations(true),
                |vp_ctx, _class| {
                    vp_ctx.set_visuals(egui::Visuals::light());

                    if vp_ctx.input(|i| i.viewport().close_requested()) {
                        do_dismiss = true;
                    }

                    egui::CentralPanel::default()
                        .frame(egui::Frame::new().fill(egui::Color32::WHITE).inner_margin(16.0))
                        .show(vp_ctx, |ui| {
                            ui.label(
                                egui::RichText::new("Tekst funnet i skjermbildet")
                                    .size(14.0)
                                    .color(egui::Color32::from_rgb(30, 30, 30))
                            );
                            ui.add_space(8.0);
                            ui.horizontal(|ui| {
                                if ui.button(egui::RichText::new("Les tekst").size(13.0)).clicked() {
                                    do_read = true;
                                }
                                if ui.button(egui::RichText::new("Kopier tekst").size(13.0)).clicked() {
                                    do_copy = true;
                                }
                                if ui.button(egui::RichText::new("Avbryt").size(13.0)).clicked() {
                                    do_dismiss = true;
                                }
                            });
                        });
                },
            );

            if do_read || do_copy {
                self.ocr_copy_mode = do_copy;
                if let Some(ocr) = &mut self.ocr {
                    if let Some(rx) = ocr.start_ocr() {
                        self.ocr_receiver = Some(rx);
                    }
                }
            }
            if do_dismiss {
                if let Some(ocr) = &mut self.ocr {
                    ocr.dismiss();
                }
            }
        }
    }
}

fn main() -> eframe::Result {
    let setup_platform = platform::create_platform();

    // Set ORT_DYLIB_PATH if not already set
    if std::env::var("ORT_DYLIB_PATH").is_err() {
        let candidates = setup_platform.ort_dylib_candidates();
        for path in &candidates {
            if std::path::Path::new(path).exists() {
                unsafe { std::env::set_var("ORT_DYLIB_PATH", path); }
                break;
            }
        }
    }

    /// Test helper: block until all pending BERT worker responses are received.
    fn drain_bert_responses(app: &mut ContextApp) {
        // Wait up to 30s for all pending responses
        let deadline = Instant::now() + Duration::from_secs(30);
        while (!app.pending_spelling_bert.is_empty()
            || !app.pending_grammar_bert.is_empty()
            || !app.pending_consonant_bert.is_empty())
            && Instant::now() < deadline
        {
            // Collect one response at a time to avoid borrow issues
            let resp = app.bert_worker.as_mut().and_then(|w| w.try_recv());
            match resp {
                Some(bert_worker::BertResponse::SpellingScore { id, scored_candidates }) => {
                    if let Some(idx) = app.pending_spelling_bert.iter().position(|p| p.request_id == id) {
                        let pending = app.pending_spelling_bert.remove(idx);
                        app.handle_spelling_bert_response(pending, &scored_candidates);
                    } else if let Some(idx) = app.pending_grammar_bert.iter().position(|p| p.request_id == id) {
                        let pending = app.pending_grammar_bert.remove(idx);
                        app.handle_grammar_bert_response(pending, &scored_candidates);
                    } else if let Some(idx) = app.pending_consonant_bert.iter().position(|p| p.request_id == id) {
                        let pending = app.pending_consonant_bert.remove(idx);
                        app.handle_consonant_bert_response(pending, &scored_candidates);
                    }
                }
                Some(bert_worker::BertResponse::MlmForward { .. }) => {}
                Some(bert_worker::BertResponse::Completion { .. }) => {}
                None => {
                    std::thread::sleep(Duration::from_millis(10));
                }
            }
        }
    }

    // Console spelling test mode — exercises exact same code as GUI
    if std::env::args().any(|a| a == "--test-spelling") {
        eprintln!("=== Spelling test mode ===");
        let mut app = ContextApp::new(true, true, 2, false);
        // Wait for startup (BERT model loading) to complete
        if let Some(rx) = app.startup_rx.take() {
            eprintln!("Waiting for BERT model to load...");
            while let Ok(item) = rx.recv() {
                match item {
                    StartupItem::Completer { model, prefix_index, baselines, wordfreq, embedding_store, errors } => {
                        if let Some(m) = model {
                            app.bert_worker = Some(bert_worker::spawn_bert_worker(
                                m, egui::Context::default(), build_bpe_completions, build_mtag_completions, build_right_completions,
                                Arc::new(prefix_index.clone().unwrap_or_default()),
                                baselines.clone(),
                                wordfreq.as_ref().cloned(),
                                embedding_store.clone(),
                                app.analyzer.clone(),
                            ));
                            app.bert_ready = true;
                            eprintln!("BERT worker spawned");
                        }
                        app.prefix_index = prefix_index;
                        app.baselines = baselines;
                        app.wordfreq = wordfreq;
                        app.embedding_store = embedding_store;
                        app.load_errors.extend(errors);
                    }
                }
            }
        }
        if !app.bert_ready {
            eprintln!("ERROR: BERT model not loaded!");
            std::process::exit(1);
        }
        let mut pass = 0;
        let mut fail = 0;

        // --- Spelling suggestion tests (word NOT in dictionary) ---
        let spelling_tests: Vec<(&str, &str, &str)> = vec![
            ("fotbal", "jeg spiller og fotbal", "fotball"),
            ("blåsjell", "vi spiser blåsjell", "blåskjell"),
            ("spitlt", "Jeg hadde spitlt fotball.", "spilt"),
            ("skriverfeil", "Det er en skriverfeil.", "skrivefeil"),
        ];
        for (word, sentence, expected) in &spelling_tests {
            app.last_spell_checked_word.clear();
            app.writing_errors.clear();
            app.pending_consonant_checks.clear();
            app.pending_consonant_bert.clear();
            app.pending_spelling_bert.clear();
            app.check_spelling(word, sentence, "", 0);
            // Drain BERT worker responses (async consonant + spelling re-ranking)
            drain_bert_responses(&mut app);
            app.validate_consonant_checks();
            app.upgrade_spelling_suggestions();
            // Drain again for any spelling upgrade requests
            drain_bert_responses(&mut app);
            let suggestion = app.writing_errors.first()
                .map(|e| e.suggestion.as_str()).unwrap_or("(none)");
            let ok = suggestion == *expected;
            if ok { pass += 1; } else { fail += 1; }
            println!("{} '{}' → '{}' (expected '{}')",
                if ok { "✓" } else { "✗" }, word, suggestion, expected);
        }

        // --- Consonant confusion tests (word IS in dictionary, sibling should win) ---
        eprintln!("\n=== Consonant confusion tests ===");
        let consonant_tests: Vec<(&str, &str, &str)> = vec![
            // (word, sentence, expected_variant_word)
            ("spil", "Det er et morsomt spil.", "spill"),
            ("spiler", "Jeg spiler og fotball.", "spiller"),
        ];
        for (word, sentence, expected_variant) in &consonant_tests {
            app.last_spell_checked_word.clear();
            app.writing_errors.clear();
            app.pending_consonant_checks.clear();
            app.pending_consonant_bert.clear();
            app.check_spelling(word, sentence, "", 0);
            // Drain BERT worker responses for consonant scoring
            drain_bert_responses(&mut app);
            eprintln!("  '{}': pending={}", word, app.pending_consonant_checks.len());
            app.validate_consonant_checks();
            let got = app.writing_errors.first()
                .map(|e| e.suggestion.as_str()).unwrap_or("(none)");
            let ok = got.contains(expected_variant);
            if ok { pass += 1; } else { fail += 1; }
            println!("{} consonant '{}' → '{}' (expected '{}')",
                if ok { "✓" } else { "✗" }, word, got, expected_variant);
        }

        println!("\n{}/{} passed", pass, pass + fail);
        std::process::exit(if fail == 0 { 0 } else { 1 });
    }

    let grammar_completion = !std::env::args().any(|a| a == "--no-grammar");
    let use_swipl = !std::env::args().any(|a| a == "--no-swipl");
    let show_debug_tab = std::env::args().any(|a| a == "--debug");
    let quality: u8 = {
        let args: Vec<String> = std::env::args().collect();
        args.iter()
            .position(|a| a == "--quality")
            .and_then(|i| args.get(i + 1))
            .and_then(|v| v.parse().ok())
            .unwrap_or(2)
    };
    if grammar_completion {
        eprintln!("Grammar completion: ON");
    }
    if use_swipl {
        eprintln!("SWI-Prolog engine: ON");
    }
    eprintln!("Quality: {} (0=fast ~200ms, 1=balanced ~800ms, 2=full ~2s)", quality);

    // Initialize TTS engine (platform-specific)
    setup_platform.init_tts();

    fn make_pen_icon(size: u32) -> egui::IconData {
        let mut rgba = vec![0u8; (size * size * 4) as usize];
        let s = size as f32;

        for y in 0..size {
            for x in 0..size {
                let fx = x as f32 / s;
                let fy = y as f32 / s;
                let idx = ((y * size + x) * 4) as usize;

                // Pen body: rotated rectangle from top-right to bottom-left
                // Line from (0.75, 0.1) to (0.2, 0.85) with thickness
                let px = fx - 0.475;
                let py = fy - 0.475;
                // Rotate 45 degrees
                let cos = 0.7071;
                let sin = 0.7071;
                let rx = px * cos + py * sin;
                let ry = -px * sin + py * cos;

                // Pen body (elongated rectangle — wide and long)
                let in_body = rx.abs() < 0.14 && ry > -0.48 && ry < 0.28;
                // Pen tip (triangle narrowing to point)
                let tip_width = 0.14 * (1.0 - (ry - 0.28) / 0.20).max(0.0);
                let in_tip = ry >= 0.28 && ry < 0.48 && rx.abs() < tip_width;
                // Pen top (slightly wider grip area)
                let in_grip = rx.abs() < 0.17 && ry > -0.48 && ry < -0.35;

                if in_tip {
                    // Gold/brass nib
                    rgba[idx] = 200;
                    rgba[idx + 1] = 160;
                    rgba[idx + 2] = 40;
                    rgba[idx + 3] = 255;
                } else if in_grip {
                    // Dark grip
                    rgba[idx] = 60;
                    rgba[idx + 1] = 60;
                    rgba[idx + 2] = 80;
                    rgba[idx + 3] = 255;
                } else if in_body {
                    // Blue pen body (NorskTale blue)
                    rgba[idx] = 0;
                    rgba[idx + 1] = 70;
                    rgba[idx + 2] = 160;
                    rgba[idx + 3] = 255;
                }
            }
        }
        egui::IconData { rgba, width: size, height: size }
    }

    let pen_icon = make_pen_icon(64);
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([420.0, 250.0])
            .with_always_on_top()
            .with_decorations(false)
            .with_title("NorskTale")
            .with_close_button(false)  // prevent Alt+F4 and system close
            .with_icon(std::sync::Arc::new(pen_icon)),
        ..Default::default()
    };

    eframe::run_native(
        "NorskTale",
        options,
        Box::new({
            let emoji_font = setup_platform.emoji_font_path().map(|s| s.to_owned());
            move |cc| {
                // Load Open Sans for dyslexia-friendly UI (recommended by British Dyslexia Association)
                let font_path = concat!(env!("CARGO_MANIFEST_DIR"), "/fonts/OpenSans-Regular.ttf");
                if let Ok(font_data) = std::fs::read(font_path) {
                    let mut fonts = egui::FontDefinitions::default();
                    fonts.font_data.insert(
                        "OpenSans".to_owned(),
                        egui::FontData::from_owned(font_data).into(),
                    );
                    fonts.families.get_mut(&egui::FontFamily::Proportional).unwrap()
                        .insert(0, "OpenSans".to_owned());
                    // Add emoji font as fallback for emoji glyphs
                    if let Some(ref emoji_path) = emoji_font {
                        if let Ok(emoji_data) = std::fs::read(emoji_path) {
                            fonts.font_data.insert(
                                "EmojiFont".to_owned(),
                                egui::FontData::from_owned(emoji_data).into(),
                            );
                            fonts.families.get_mut(&egui::FontFamily::Proportional).unwrap()
                                .push("EmojiFont".to_owned());
                        }
                    }
                    cc.egui_ctx.set_fonts(fonts);
                    eprintln!("Loaded Open Sans font");
                } else {
                    eprintln!("Warning: Open Sans font not found at {}", font_path);
                }
                Ok(Box::new(ContextApp::new(grammar_completion, use_swipl, quality, show_debug_tab)))
            }
        }),
    )
}
