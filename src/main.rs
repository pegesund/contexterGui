#[macro_use]
pub mod logging;

mod bert_worker;
mod bridge;
pub mod downloader;
mod grammar_actor;
mod latext_no;
mod math_ocr;
mod ocr;
mod platform;
mod stt;
mod tts;
pub mod user_dict;
pub mod spelling_scorer;
pub mod llm_actor;
pub mod compound_walker;

use bridge::{CursorContext, TextBridge};
use std::collections::HashMap;
use std::hash::{Hash, Hasher};
use std::io::Write;
use std::path::PathBuf;
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

// ---------------------------------------------------------------------------
// Persistent user settings (saved to ~/Library/Application Support/NorskTale/)
// ---------------------------------------------------------------------------

#[derive(serde::Serialize, serde::Deserialize)]
struct UserSettings {
    quality: u8,
    whisper_mode: u8,
    speak_on_space: bool,
    ui_scale: f32,
    voice: String,
    #[serde(default = "default_language")]
    language: String,
}

fn default_language() -> String { "nb".into() }

impl Default for UserSettings {
    fn default() -> Self {
        UserSettings {
            quality: 1,
            whisper_mode: 1,
            speak_on_space: true,
            ui_scale: 1.0,
            voice: String::new(),
            language: "nb".into(),
        }
    }
}

fn settings_path() -> PathBuf {
    let dir = if cfg!(target_os = "macos") {
        dirs::home_dir()
            .map(|h| h.join("Library/Application Support/NorskTale"))
            .unwrap_or_else(|| PathBuf::from("/tmp"))
    } else {
        dirs::config_dir()
            .map(|c| c.join("NorskTale"))
            .unwrap_or_else(|| PathBuf::from("."))
    };
    let _ = std::fs::create_dir_all(&dir);
    dir.join("settings.json")
}

fn load_settings() -> UserSettings {
    let path = settings_path();
    match std::fs::read_to_string(&path) {
        Ok(json) => serde_json::from_str(&json).unwrap_or_default(),
        Err(_) => UserSettings::default(),
    }
}

fn save_settings(s: &UserSettings) {
    let path = settings_path();
    if let Ok(json) = serde_json::to_string_pretty(s) {
        let _ = std::fs::write(&path, json);
    }
}

/// Truncate a string to at most `max` bytes, backing up to the nearest char boundary.
/// Compute boost multiplier for a candidate word based on document frequency and user dictionary.
/// Returns 1.0 (no boost) for common function words or words not in doc/user_dict.
pub fn compute_boost(
    word: &str,
    doc_word_counts: &HashMap<String, u16>,
    user_dict: Option<&user_dict::UserDict>,
    wordfreq: Option<&HashMap<String, u64>>,
    lang: &dyn language::LanguageBundle,
) -> f32 {
    let lower = word.to_lowercase();
    let common_threshold = lang.wordfreq_common_threshold();
    // Never boost common function words (top ~31 in Norwegian: og, er, for, til, av, det, ...)
    if wordfreq.and_then(|wf| wf.get(&lower)).map_or(false, |&f| f >= common_threshold) {
        return 1.0;
    }
    let in_doc = doc_word_counts.get(&lower).copied().unwrap_or(0) >= 2;
    let in_user = user_dict.map_or(false, |ud| ud.has_word(&lower));
    match (in_doc, in_user) {
        (true, true)   => 1.6,  // strongest: in both document and user dict
        (false, true)  => 1.3,  // moderate: user's validated vocabulary
        (true, false)  => 1.25, // mild: topic word in current document
        (false, false) => 1.0,
    }
}

fn levenshtein_distance(a: &str, b: &str) -> u32 {
    let (a, b): (Vec<char>, Vec<char>) = (a.chars().collect(), b.chars().collect());
    let (m, n) = (a.len(), b.len());
    let mut dp = vec![vec![0u32; n + 1]; m + 1];
    for i in 0..=m { dp[i][0] = i as u32; }
    for j in 0..=n { dp[0][j] = j as u32; }
    for i in 1..=m {
        for j in 1..=n {
            let cost = if a[i-1] == b[j-1] { 0 } else { 1 };
            dp[i][j] = (dp[i-1][j] + 1).min(dp[i][j-1] + 1).min(dp[i-1][j-1] + cost);
        }
    }
    dp[m][n]
}

fn trunc(s: &str, max: usize) -> &str {
    if s.len() <= max { return s; }
    let mut end = max;
    while end > 0 && !s.is_char_boundary(end) { end -= 1; }
    &s[..end]
}

use nostos_cognio::baseline::{compute_baseline, Baselines};
use nostos_cognio::complete::{complete_word, grammar_filter, GrammarCheckResult, Completion};
use nostos_cognio::embeddings::EmbeddingStore;
use nostos_cognio::grammar::swipl_checker::SwiGrammarChecker;
use nostos_cognio::grammar::types::GrammarError;
use nostos_cognio::model::Model;
use nostos_cognio::prefix_index::{self, PrefixIndex};
use nostos_cognio::wordfreq;

// --- Grammar checker abstraction ---

pub(crate) enum AnyChecker {
    Swi(SwiGrammarChecker),
}

// SAFETY: AnyChecker is only ever accessed from one thread at a time.
// SWI-Prolog's raw pointers (PredicateT) are !Send, but the grammar actor
// ensures single-threaded access via mpsc channel serialization.
unsafe impl Send for AnyChecker {}

impl AnyChecker {
    fn has_word(&self, word: &str) -> bool {
        match self {
            AnyChecker::Swi(c) => c.has_word(word),
        }
    }

    fn prefix_lookup(&self, prefix: &str, limit: usize) -> Vec<String> {
        match self {
            AnyChecker::Swi(c) => c.prefix_lookup(prefix, limit),
        }
    }

    fn check_sentence(&mut self, text: &str) -> Vec<GrammarError> {
        match self {
            AnyChecker::Swi(c) => c.check_sentence(text),
        }
    }

    fn has_error(&mut self, text: &str) -> bool {
        match self {
            AnyChecker::Swi(c) => c.has_error(text),
        }
    }

    fn check_sentence_full(&mut self, text: &str) -> nostos_cognio::grammar::types::CheckResult {
        match self {
            AnyChecker::Swi(c) => c.check_sentence_full(text),
        }
    }

    fn check_sentence_full_with_doc(&mut self, text: &str, doc_text: &str) -> nostos_cognio::grammar::types::CheckResult {
        match self {
            AnyChecker::Swi(c) => c.check_sentence_full_with_doc(text, doc_text),
        }
    }

    fn fuzzy_lookup(&self, word: &str, max_distance: u32) -> Vec<(String, u32)> {
        match self {
            AnyChecker::Swi(c) => c.fuzzy_lookup(word, max_distance),
        }
    }

    fn suggest_compound(&self, word: &str) -> Option<String> {
        match self {
            AnyChecker::Swi(c) => c.suggest_compound(word),
        }
    }

    /// Get the set of POS tags for a word from the dictionary
    fn pos_set(&self, word: &str) -> std::collections::HashSet<String> {
        let analyzer = match self {
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
    /// The specific trigger word for grammar errors (what gets underlined)
    pub(crate) error_word: String,
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
    // Try both Mac layout (../contexter-repo) and Windows layout (../../contexter-repo)
    let mac = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../contexter-repo/training-data");
    if mac.exists() { return mac; }
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../contexter-repo/training-data")
}

/// Resolve a data file path: use S3-downloaded cache if available, otherwise
/// fall back to the language trait path (local dev layout).
fn cached_or_trait(cached: &std::path::Path, trait_path: PathBuf) -> PathBuf {
    if cached.exists() { cached.to_path_buf() } else { trait_path }
}

/// Resolve all language data paths, preferring S3-cached files.
struct ResolvedPaths {
    mtag_fst: PathBuf,
    onnx: PathBuf,
    tokenizer: PathBuf,
    wordfreq: PathBuf,
    prolog_rules: PathBuf,
    /// Directory containing compound_data.pl + sentence_split.pl
    prolog_dir: PathBuf,
}

fn resolve_paths(lang: &dyn language::LanguageBundle) -> ResolvedPaths {
    let cache = downloader::data_dir();
    let code = lang.code(); // "nb" or "nn"
    let lang_dir = cache.join(format!("lang/{}", code));
    let bert_dir = cache.join("models/bert");

    // Per-language file names in cache
    let fst_name = if code == "nn" { "fullform_nn.mfst" } else { "fullform_bm.mfst" };
    let wf_name = if code == "nn" { "wordfreq_nn.tsv" } else { "wordfreq_bm.tsv" };

    let mtag_fst = cached_or_trait(&lang_dir.join(fst_name), lang.mtag_fst_path());
    let onnx = cached_or_trait(&bert_dir.join("norbert4_base_int8.onnx"), lang.onnx_path());
    let tokenizer = cached_or_trait(&bert_dir.join("tokenizer.json"), lang.tokenizer_path());
    let wordfreq = cached_or_trait(&lang_dir.join(wf_name), lang.wordfreq_path());
    let prolog_rules = cached_or_trait(&lang_dir.join("grammar_rules.pl"), lang.prolog_rules_path());

    // Prolog dir: parent of grammar_rules.pl (compound_data.pl lives there)
    let prolog_dir = prolog_rules.parent().unwrap_or(std::path::Path::new(".")).to_path_buf();

    eprintln!("Resolved paths for '{}':", code);
    eprintln!("  FST:     {}", mtag_fst.display());
    eprintln!("  ONNX:    {}", onnx.display());
    eprintln!("  Wordfreq:{}", wordfreq.display());
    eprintln!("  Prolog:  {}", prolog_rules.display());

    ResolvedPaths { mtag_fst, onnx, tokenizer, wordfreq, prolog_rules, prolog_dir }
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
    /// Windows Word COM language ID for the active language (from LanguageVoice trait)
    lang_word_id: i32,
}

impl BridgeManager {
    fn new(platform: Box<dyn platform::PlatformServices>, lang_word_id: i32) -> Self {
        let mut bridges: Vec<Box<dyn TextBridge>> = bridge::create_bridges(lang_word_id);
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
            lang_word_id,
        }
    }

    fn read_context(&mut self) -> Option<CursorContext> {
        if self.last_check.elapsed() > Duration::from_secs(5) {
            self.last_check = Instant::now();
            let has_word = self.bridges.iter().any(|b| b.name().contains("Word"));
            if !has_word {
                for new_bridge in bridge::try_connect_word_bridge(self.lang_word_id) {
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
        // Check it first — but ONLY when Word is actually foreground (or our own
        // window is). If the user has switched to a browser or other app, skip
        // the Add-in cache so the correct bridge can win immediately.
        if !is_browser {
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
        }

        // Our window is foreground — return cached context.
        // NEVER try COM calls here — causes tight loop freeze.
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

    /// Clear cached context so stale Word data is not shown after an app switch.
    fn clear_context(&mut self) {
        self.last_context = None;
        self.last_user_was_browser = false;
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

    /// Read selected text from any available bridge.
    fn read_selected_text(&self) -> Option<String> {
        for b in &self.bridges {
            if let Some(text) = b.read_selected_text() {
                return Some(text);
            }
        }
        None
    }

    fn read_paragraph_at(&self, cursor_offset: usize) -> Option<(String, String, usize)> {
        self.effective_bridge().and_then(|b| b.read_paragraph_at(cursor_offset))
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

    fn select_word_in_paragraph(&self, word: &str, paragraph_id: &str) -> bool {
        for bridge in &self.bridges {
            if bridge.select_word_in_paragraph(word, paragraph_id) {
                return true;
            }
        }
        false
    }

    fn set_target_hwnd(&self, hwnd: isize) {
        for bridge in &self.bridges {
            bridge.set_target_hwnd(hwnd);
        }
    }

    fn mark_error_underline(&self, char_start: usize, char_end: usize, color: bridge::ErrorUnderlineColor) -> bool {
        self.effective_bridge().map(|b| b.mark_error_underline(char_start, char_end, color)).unwrap_or(false)
    }

    fn clear_error_underline(&self, char_start: usize, char_end: usize) -> bool {
        self.effective_bridge().map(|b| b.clear_error_underline(char_start, char_end)).unwrap_or(false)
    }

    fn clear_all_error_underlines(&self) -> bool {
        self.effective_bridge().map(|b| b.clear_all_error_underlines()).unwrap_or(false)
    }

    fn underline_word(&self, word: &str, paragraph_id: &str, color: &str) -> bool {
        let mut any = false;
        for b in &self.bridges { any |= b.underline_word(word, paragraph_id, color); }
        any
    }

    fn clear_underline_word(&self, word: &str, paragraph_id: &str) -> bool {
        let mut any = false;
        for b in &self.bridges { any |= b.clear_underline_word(word, paragraph_id); }
        any
    }

    fn clear_paragraph_underlines(&self, paragraph_id: &str) {
        for b in &self.bridges { b.clear_paragraph_underlines(paragraph_id); }
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
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '\\' => out.push_str("\\\\"),
            '"' => out.push_str("\\\""),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out
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
    /// Phase 14: runtime-selected language. Resolved once at startup from
    /// the --language CLI flag and shared by every part of the app that
    /// needs language-specific data (FST path, Prolog rules, BERT model,
    /// UI strings, voice/STT/OCR codes, …).
    language: std::sync::Arc<dyn language::LanguageBundle>,
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
    /// Shared errors JSON for HTTP /errors endpoint (test verification)
    shared_errors_json: std::sync::Arc<std::sync::Mutex<String>>,
    // Word completer — BERT model lives in dedicated worker thread (no lock contention)
    bert_worker: Option<bert_worker::BertWorkerHandle>,
    bert_ready: bool,
    completion_cancel: Arc<std::sync::atomic::AtomicBool>,
    /// Last time context changed — for debouncing completion dispatch
    last_context_change: Instant,
    /// The cache key we last dispatched (avoid re-dispatching same)
    dispatched_key: String,
    last_dispatched_sentence: String,
    pending_incomplete_sentence: Option<(String, String, Instant)>, // (sentence, para_id, timestamp)
    grammar_inflight: std::collections::HashSet<u64>, // hashes of sentences sent to grammar actor, not yet responded
    paragraph_texts: std::collections::HashMap<String, String>, // paragraph_id → latest text, for building doc text
    last_grammar_ctx_key: String,
    last_known_cursor_offset: Option<usize>,
    prefix_index: Option<PrefixIndex>,
    baselines: Option<Arc<Baselines>>,
    wordfreq: Option<Arc<HashMap<String, u64>>>,
    /// Raw FST for compound word decomposition (Source 13)
    compound_fst: Option<Arc<fst::raw::Fst<Vec<u8>>>>,
    embedding_store: Option<Arc<EmbeddingStore>>,
    completions: Vec<Completion>,
    /// Open suggestions (any word) for fill-in-the-blank mode
    open_completions: Vec<Completion>,
    last_completed_prefix: String,
    /// Timestamp when last CompleteWord was dispatched (for round-trip measurement)
    last_completion_dispatch: Instant,
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
    speak_on_space: bool,
    last_space_speak: Instant,
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
    /// Resolved data paths (S3 cache or local dev)
    resolved_paths: ResolvedPaths,
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
    /// User dictionary — words added by the user (persisted in redb)
    user_dict: Option<user_dict::UserDict>,
    /// Show user dictionary editor window
    show_settings_window: bool,
    show_userdict_window: bool,
    /// Text input for new word in user dict editor
    userdict_new_word: String,
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
    /// Document word frequency map — rebuilt when doc hash changes.
    /// Maps lowercase word → count in current document.
    doc_word_counts: HashMap<String, u16>,
    /// Hash when doc_word_counts was last built
    doc_word_counts_hash: u64,
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
    /// LLM grammar correction
    llm_actor: Option<llm_actor::LlmActorHandle>,
    llm_checked_hashes: std::collections::HashSet<u64>,
    llm_sent_count: Vec<Instant>,           // rate limiting (rolling hour)
    llm_waiting: bool,                      // spinner: waiting for LLM response
    llm_waiting_since: Instant,             // when we started waiting
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
    /// LLM changes for the currently open rule_info_window (from, to, why)
    rule_info_llm_changes: Vec<(String, String, String)>,
    // OCR clipboard monitoring
    ocr: Option<ocr::OcrClipboard>,
    ocr_receiver: Option<std::sync::mpsc::Receiver<Result<String, String>>>,
    ocr_text: Option<String>,
    ocr_copy_mode: bool, // true = copy to clipboard, false = speak
    // Math OCR (lazy-loaded)
    math_receiver: Option<std::sync::mpsc::Receiver<Result<String, String>>>,
    // Microphone / Whisper
    whisper_engine: Option<Arc<Mutex<Box<dyn stt::SttEngine>>>>,       // final model (medium-q5 or tiny)
    whisper_streaming: Option<Arc<Mutex<Box<dyn stt::SttEngine>>>>,    // streaming model (base; None in tiny mode)
    mic_handle: Option<stt::MicHandle>,
    mic_transcribing: bool,
    mic_result_text: Option<String>,
    /// Receiver for "Forbedre" re-transcription result (Windows only)
    improve_rx: Option<std::sync::mpsc::Receiver<String>>,
    improve_running: bool,
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
    /// Tracks whether the foreground app was a browser in the previous poll,
    /// so we can detect the Word→Browser transition and clear stale errors.
    prev_fg_was_browser: bool,
    /// Window title of the last foreground Word window, used to detect
    /// document switches (Document1 → Document2) and clear stale errors.
    prev_word_title: String,
    /// True when the foreground app is a browser this frame. Set at the
    /// top of every update() so grammar/BERT pollers can gate on it.
    suppress_errors: bool,
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
    baselines: Option<&Baselines>,
    analyzer: Option<&mtag::Analyzer>,
) -> Vec<Completion> {
    let is_valid = |w: &str| -> bool {
        let key = w.to_lowercase();
        if nearby_words.contains(&key) { return false; }
        if !wordfreq.map_or(true, |wf| wf.contains_key(&key)) { return false; }
        // mtag filter: only Norwegian words (removes Danish/English junk)
        if let Some(az) = analyzer {
            if !az.has_word(&key) { return false; }
        }
        true
    };

    // PMI: subtract baseline to demote generically common words,
    // boost contextually relevant ones. "is" is common everywhere (demoted),
    // "idrett" is specific to sports context (boosted).
    let pmi_logits: Vec<f32> = if let Some(bl) = baselines {
        logits.iter().enumerate().map(|(i, &raw)| {
            let base = if i < bl.sentence.len() { bl.sentence[i] } else { 0.0 };
            raw + 1.0 * (raw - base)
        }).collect()
    } else {
        logits.to_vec()
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
            Some((decoded, pmi_logits[i]))
        })
        .collect();
    all_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
    all_scored.into_iter()
        .take(10)
        .map(|(w, s)| Completion { word: w, score: s, elapsed_ms: 0.0 })
        .collect()
}

impl ContextApp {
    /// Reset all error/spelling/grammar state when the user switches to a
    /// different app or document, so stale results don't bleed across contexts.
    fn clear_for_app_switch(&mut self) {
        self.writing_errors.clear();
        self.context = Default::default();
        self.spelling_queue.clear();
        self.pending_spelling_bert.clear();
        self.pending_grammar_bert.clear();
        self.pending_consonant_bert.clear();
        self.grammar_queue.clear();
        self.grammar_queue_total = 0;
        self.processed_sentence_hashes.clear();
        self.last_doc_hash = 0;
        // Clear paragraph tracking so in-flight grammar actor results are
        // treated as stale and discarded (the guard checks this map).
        self.paragraph_sentence_hashes.clear();
        self.grammar_inflight.clear();
        self.manager.clear_context();
    }

    fn new(
        language: std::sync::Arc<dyn language::LanguageBundle>,
        grammar_completion: bool,
        quality: u8,
        show_debug_tab: bool,
        saved_settings: UserSettings,
        paths: ResolvedPaths,
    ) -> Self {
        let platform = platform::create_platform();
        platform.init_runtime();

        let mut load_errors = Vec::new();

        // Load dictionary from resolved path (S3 cache or local dev).
        let analyzer: Option<std::sync::Arc<mtag::Analyzer>> = match mtag::Analyzer::new(&paths.mtag_fst) {
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
        let compound_fst: Option<Arc<fst::raw::Fst<Vec<u8>>>> =
            compound_walker::load_fst_from_mfst(paths.mtag_fst.to_str().unwrap())
                .ok().map(|f| Arc::new(f));

        // Spawn heavy model loading on background threads
        let (startup_tx, startup_rx) = std::sync::mpsc::channel();

        // Thread 1: NorBERT4 + completer (using resolved paths)
        let tx2 = startup_tx.clone();
        let onnx_path = paths.onnx.clone();
        let tokenizer_path = paths.tokenizer.clone();
        let wordfreq_path = paths.wordfreq.clone();
        std::thread::spawn(move || {
            let data = data_dir();
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
            language: language.clone(),
            manager: BridgeManager::new(platform::create_platform(), language.word_language_id()),
            context: CursorContext::default(),
            last_poll: Instant::now(),
            poll_interval: Duration::from_millis(100),
            follow_cursor: true,
            goto_freeze_until: None,
            last_caret_pos: None,
            checker: None,
            analyzer,
            compound_fst,
            grammar_actor: None,
            grammar_errors: Vec::new(),
            last_checked_sentence: String::new(),
            shared_errors_json: std::sync::Arc::new(std::sync::Mutex::new("[]".to_string())),
            bert_worker: None,
            bert_ready: false,
            completion_cancel: Arc::new(std::sync::atomic::AtomicBool::new(false)),
            last_context_change: Instant::now(),
            dispatched_key: String::new(),
            last_dispatched_sentence: String::new(),
            pending_incomplete_sentence: None,
            grammar_inflight: std::collections::HashSet::new(),
            paragraph_texts: std::collections::HashMap::new(),
            last_grammar_ctx_key: String::new(),
            last_known_cursor_offset: None,
            prefix_index: None,
            baselines: None,
            wordfreq: None,
            embedding_store: None,
            completions: Vec::new(),
            open_completions: Vec::new(),
            last_completed_prefix: String::new(),
            last_completion_dispatch: Instant::now(),
            last_replace_time: Instant::now() - Duration::from_secs(10),
            cached_forward: None,
            cached_right_column: None,
            cached_mtag_supplement: None,
            last_embedding_sync: Instant::now(),
            embedding_sync_interval: Duration::from_secs(3),
            grammar_completion,
            speak_on_space: saved_settings.speak_on_space,
            last_space_speak: Instant::now(),
            quality,
            last_prefix_change: Instant::now(),
            debounce_ms: if quality == 0 { 100 } else { 150 },
            pending_completion: false,
            selected_completion: None,
            selection_mode: false,
            app_handle: None,
            platform,
            resolved_paths: paths,
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
            user_dict: {
                let dir = dirs::home_dir().unwrap_or_default().join(".norsktale");
                let _ = std::fs::create_dir_all(&dir);
                match user_dict::UserDict::open(dir.join("user_words.db")) {
                    Ok(ud) => {
                        let count = ud.list_words().len();
                        if count > 0 { eprintln!("User dictionary: {} words loaded", count); }
                        Some(ud)
                    }
                    Err(e) => { eprintln!("User dictionary unavailable: {}", e); None }
                }
            },
            show_settings_window: false,
            show_userdict_window: false,
            userdict_new_word: String::new(),
            last_spell_checked_word: String::new(),
            last_doc_text: String::new(),
            last_doc_approx_len: 0,
            last_replaced_word: None,
            last_doc_hash: 0,
            doc_word_counts: HashMap::new(),
            doc_word_counts_hash: 0,
            last_sentence_count: 0,
            prolog_checked_hashes: std::collections::HashSet::new(),
            processed_sentence_hashes: std::collections::HashSet::new(),
            paragraph_sentence_hashes: HashMap::new(),
            spelling_queue: Vec::new(),
            grammar_queue: Vec::new(),
            grammar_queue_total: 0,
            grammar_scanning: false,
            llm_actor: None,
            llm_checked_hashes: std::collections::HashSet::new(),
            llm_sent_count: Vec::new(),
            llm_waiting: false,
            llm_waiting_since: Instant::now(),
            pending_fix: None,
            pending_consonant_checks: Vec::new(),
            pending_spelling_bert: Vec::new(),
            pending_grammar_bert: Vec::new(),
            pending_consonant_bert: Vec::new(),
            suggestion_window: None,
            suggestion_selection: std::sync::Arc::new(std::sync::Mutex::new(None)),
            rule_info_window: None,
            rule_info_llm_changes: Vec::new(),
            ocr: match ocr::OcrClipboard::new(&*language) {
                Ok(o) => { eprintln!("OCR clipboard monitor ready"); Some(o) }
                Err(e) => { eprintln!("OCR not available: {}", e); None }
            },
            ocr_receiver: None,
            ocr_text: None,
            ocr_copy_mode: false,
            math_receiver: None,
            whisper_engine: None,
            whisper_streaming: None,
            mic_handle: None,
            mic_transcribing: false,
            mic_result_text: None,
            improve_rx: None,
            improve_running: false,
            whisper_mode: saved_settings.whisper_mode,
            whisper_load_rx: None,
            whisper_loading: false,
            whisper_load_status: String::new(),
            whisper_pending_record: false,
            show_voice_window: false,
            ui_scale: saved_settings.ui_scale,
            voice_list: Vec::new(),
            startup_rx: Some(startup_rx),
            startup_done: Vec::new(),
            startup_total: 1, // completer only
            prev_fg_was_browser: false,
            prev_word_title: String::new(),
            suppress_errors: false,
        }
    }

    fn load_swipl_checker(swipl_path: &str, fst_path: &str, prolog_rules: &str, prolog_dir: &str) -> Result<SwiGrammarChecker, Box<dyn std::error::Error>> {
        SwiGrammarChecker::new(
            swipl_path,
            fst_path,
            prolog_rules,
            prolog_dir,
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

        // Skip if word is in user dictionary
        if self.user_dict.as_ref().map_or(false, |ud| ud.has_word(&clean)) {
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
        // Skip email addresses and words that are parts of email addresses
        // e.g. "havard.rye@tekna.no" may be split into "havard", "rye", "tekna", "no"
        if clean.contains('@') {
            return;
        }
        if sentence_ctx.contains('@') {
            // Check if this word is part of an email in the sentence
            let word_lower = clean.to_lowercase();
            for token in sentence_ctx.split_whitespace() {
                let t = token.trim_matches(|c: char| c == '(' || c == ')' || c == ',' || c == ';');
                if t.contains('@') {
                    let email_parts: Vec<&str> = t.split(|c: char| c == '@' || c == '.' || c == '-' || c == '_').collect();
                    if email_parts.iter().any(|p| p.to_lowercase() == word_lower) {
                        return;
                    }
                }
            }
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

        // Phase 14: modal verb misspellings come from the runtime-selected
        // language. BERT can't distinguish forms like "vil" vs "ville"
        // in context, so each language carries its own pair list.
        let modal_fixes = self.language.modal_confusion_pairs();
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
                    word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: paragraph_id.to_string(), error_word: String::new(),
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
            // Check abbreviation: "osv" without period → try "osv."
            if !found && analyzer.has_word(&format!("{}.", clean)) {
                found = true;
            }
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
                // Also accept words found in wordfreq (freq ≥ 10 in Norwegian corpus)
                if !found {
                    if let Some(wf) = &self.wordfreq {
                        if wf.contains_key(&clean) {
                            found = true;
                            log!("spell: '{}' accepted via wordfreq", clean);
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
            explanation: self.language.ui_word_not_in_dict(&clean),
            rule_name: rule.to_string(),
            sentence_context: sentence_ctx.to_string(),
            doc_offset,
            position: 0,
            ignored: false,
            word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: paragraph_id.to_string(), error_word: String::new(),
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
        // If the user is at a word boundary (pressed space or end-of-sentence),
        // context_after is empty/whitespace — which breaks score_and_rerank's
        // word-position extraction (find("") returns Some(0) → word_lower = "").
        // Append a sentinel "." to BOTH context_after AND sentence so the word
        // extraction inside the scorer hits the else branch and correctly strips
        // the trailing punctuation to recover the misspelled word.
        let (context_after, sentence) = if context_after.trim().is_empty() {
            (".".to_string(), format!("{}.", sentence_ctx))
        } else {
            (context_after, sentence_ctx.to_string())
        };
        let request_id = worker.send(|id| bert_worker::BertRequest::SpellingScore { id, context_before, context_after, candidates, sentence });
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
        // Drain and discard all BERT results while browser is foreground.
        if self.suppress_errors {
            if let Some(worker) = &mut self.bert_worker {
                while worker.try_recv().is_some() {}
            }
            return;
        }
        // Collect all available responses first (avoids borrow conflicts)
        let mut responses: Vec<bert_worker::BertResponse> = Vec::new();
        if let Some(worker) = &mut self.bert_worker {
            while let Some(resp) = worker.try_recv() {
                responses.push(resp);
            }
        }

        // Only keep the LAST Completion response — discard older ones
        let mut last_completion: Option<(String, Vec<Completion>, Vec<Completion>)> = None;
        let mut other_responses = Vec::new();
        for resp in responses {
            match resp {
                bert_worker::BertResponse::Completion { id: _, cache_key, left, right } => {
                    last_completion = Some((cache_key, left, right)); // overwrite — keep latest
                }
                other => other_responses.push(other),
            }
        }
        // Process the one completion result
        if let Some((cache_key, left, right)) = last_completion {
            {
                    log!("BERT completion received: {} left [{}] | {} right [{}] (round-trip: {}ms)",
                        left.len(), left.iter().take(10).map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "),
                        right.len(), right.iter().take(10).map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "),
                        self.last_completion_dispatch.elapsed().as_millis());
                    // Completions arrive already dictionary + grammar filtered from worker thread
                    // Apply document-frequency and user-dict boosting, then re-sort
                    {
                        self.rebuild_doc_word_counts();
                        // Capitalize all completions if after period or user typed uppercase
                        let capitalize = self.context.sentence.trim().is_empty()
                            || self.context.word.chars().next().map_or(false, |c| c.is_uppercase());
                        let mut left_boosted: Vec<_> = left;
                        for c in &mut left_boosted {
                            c.score *= compute_boost(&c.word, &self.doc_word_counts,
                                self.user_dict.as_ref(), self.wordfreq.as_deref(), &*self.language);
                            if capitalize && c.word.chars().next().map_or(false, |ch| ch.is_lowercase()) {
                                let mut chars = c.word.chars();
                                c.word = chars.next().unwrap().to_uppercase().to_string() + chars.as_str();
                            }
                        }
                        left_boosted.sort_by(|a, b| b.score.partial_cmp(&a.score).unwrap_or(std::cmp::Ordering::Equal));
                        self.completions = left_boosted.into_iter().take(5).collect();

                        let mut right_boosted: Vec<_> = right;
                        for c in &mut right_boosted {
                            c.score *= compute_boost(&c.word, &self.doc_word_counts,
                                self.user_dict.as_ref(), self.wordfreq.as_deref(), &*self.language);
                        }
                        right_boosted.sort_by(|a, b| b.score.partial_cmp(&a.score).unwrap_or(std::cmp::Ordering::Equal));
                        self.open_completions = right_boosted.into_iter().take(5).collect();

                        // Inject document words and user dict words matching the prefix
                        let prefix = extract_prefix(&self.context.word).to_lowercase();
                        if prefix.len() >= 1 {
                            let existing: std::collections::HashSet<String> = self.completions.iter()
                                .map(|c| c.word.to_lowercase()).collect();

                            // Inject document words (sorted by count, highest first)
                            log!("Doc inject: prefix='{}' doc_word_counts={} entries, matches: {:?}",
                                prefix, self.doc_word_counts.len(),
                                self.doc_word_counts.iter()
                                    .filter(|(w, _)| w.starts_with(&prefix))
                                    .take(5).collect::<Vec<_>>());
                            let mut doc_matches: Vec<(&String, &u16)> = self.doc_word_counts.iter()
                                .filter(|(w, count)| **count >= 1 && w.starts_with(&prefix) && w.len() > prefix.len() && !existing.contains(w.as_str()))
                                .collect();
                            doc_matches.sort_by(|a, b| b.1.cmp(a.1));
                            let after_period = self.context.sentence.trim().is_empty()
                                || self.context.word.chars().next().map_or(false, |c| c.is_uppercase());
                            for (dw, count) in doc_matches.into_iter().take(3) {
                                // Restore original casing from paragraph_texts
                                let mut word = self.paragraph_texts.values()
                                    .flat_map(|t| t.split(|c: char| !c.is_alphanumeric() && c != '-'))
                                    .find(|w| w.to_lowercase() == *dw)
                                    .unwrap_or(dw.as_str())
                                    .to_string();
                                // After period or if user typed uppercase: capitalize
                                if after_period && word.chars().next().map_or(false, |c| c.is_lowercase()) {
                                    let mut chars = word.chars();
                                    word = chars.next().unwrap().to_uppercase().to_string() + chars.as_str();
                                }
                                self.completions.insert(0, nostos_cognio::complete::Completion {
                                    word, score: 50.0 + *count as f32, elapsed_ms: 0.0,
                                });
                            }

                            // Inject user dict words
                            if let Some(ud) = &self.user_dict {
                                let existing: std::collections::HashSet<String> = self.completions.iter()
                                    .map(|c| c.word.to_lowercase()).collect();
                                for uw in ud.list_words() {
                                    if uw.starts_with(&prefix) && uw.len() > prefix.len() && !existing.contains(&uw) {
                                        self.completions.insert(0, nostos_cognio::complete::Completion {
                                            word: uw, score: 100.0, elapsed_ms: 0.0,
                                        });
                                    }
                                }
                            }
                            self.completions.truncate(5);
                        }
                    }
                self.last_completed_prefix = cache_key;
            }
        }
        // Process non-completion responses
        for resp in other_responses {
            match resp {
                bert_worker::BertResponse::Completion { .. } => unreachable!(),
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
                let boost = compute_boost(candidate, &self.doc_word_counts,
                    self.user_dict.as_ref(), self.wordfreq.as_deref(), &*self.language);
                let final_score = bert_score * ortho_sim.sqrt() * boost;
                log!("  spelling BERT: '{}' bert={:.3} × sqrt(ortho {:.2}) × boost {:.2} = {:.3}", candidate, bert_score, ortho_sim, boost, final_score);
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
                        // Capitalize if at sentence start or originally capitalized
                        let mut suggestion = best.trim_matches(|c: char| c.is_whitespace() || c.is_control()).to_string();
                        let word_lower = e.word.to_lowercase();
                        let at_sentence_start = e.sentence_context.to_lowercase().starts_with(&word_lower);
                        let is_upper = e.sentence_context.to_lowercase().find(&word_lower)
                            .and_then(|pos| e.sentence_context[pos..].chars().next())
                            .map_or(false, |c| c.is_uppercase());
                        if at_sentence_start || is_upper {
                            let mut chars = suggestion.chars();
                            suggestion = chars.next().unwrap().to_uppercase().to_string() + chars.as_str();
                        }
                        log!("spelling BERT upgrade: '{}' → '{}' (was '{}')", e.word, suggestion, e.suggestion);
                        e.suggestion = suggestion;
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
                    word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: String::new(), error_word: String::new(),
                });
            }
        }
    }

    /// Re-run unified suggestion pipeline for spelling errors that were created before
    /// BERT was available. Once BERT loads, this upgrades ortho-only suggestions to BERT-ranked ones.
    fn upgrade_spelling_suggestions(&mut self) {
        if !self.bert_ready { return; }

        // Skip upgrade — find_spelling_suggestions is too slow for main thread.
        // Grammar actor already provides suggestions; BERT re-ranking happens async.
        return;
        #[allow(unreachable_code)]
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
            // For long words with an existing suggestion that's close (edit distance ≤ 2),
            // keep it — BERT can't score multi-token compound words properly
            let original = &self.writing_errors[idx].suggestion;
            if word.len() >= 10 && !original.is_empty() && original.len() >= 8 {
                let dist = levenshtein_distance(&word.to_lowercase(), &original.to_lowercase());
                if dist <= 2 {
                    self.writing_errors[idx].rule_name = "stavefeil_bert".to_string();
                    return;
                }
            }
            let suggestions = self.find_spelling_suggestions(&word, &sentence_ctx);
            if let Some((best, score)) = suggestions.first() {
                if !best.is_empty() {
                    let mut suggestion = best.trim_matches(|c: char| c.is_whitespace() || c.is_control()).to_string();
                    let word_lower = word.to_lowercase();
                    let at_start = sentence_ctx.to_lowercase().starts_with(&word_lower);
                    let is_upper = sentence_ctx.to_lowercase().find(&word_lower)
                        .and_then(|pos| sentence_ctx[pos..].chars().next())
                        .map_or(false, |c| c.is_uppercase());
                    if at_start || is_upper {
                        let mut chars = suggestion.chars();
                        suggestion = chars.next().unwrap().to_uppercase().to_string() + chars.as_str();
                    }
                    log!("Spelling upgrade: '{}' → '{}' score={:.2}", word, suggestion, score);
                    self.writing_errors[idx].suggestion = suggestion;
                }
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

            let _t = Instant::now();
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
            if let Some(split) = try_split_function_word(
                &word_lower,
                analyzer,
                self.language.function_words(),
            ) {
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

        // Source 8: User dictionary — words within edit distance 2
        if let Some(ud) = &self.user_dict {
            for uw in ud.list_words() {
                if uw == word_lower || seen.contains(&uw) { continue; }
                let dist = levenshtein_distance(&word_lower, &uw);
                if dist <= 2 {
                    edit_distances.entry(uw.clone()).or_insert(dist);
                    if seen.insert(uw.clone()) {
                        candidates.push(uw);
                    }
                }
            }
        }

        // Source 9: Long word truncation (>= 10 chars) — strip 1-2 chars from start/end
        // Catches typos on compound words (e.g. "PKarrierekompasset" → "karrierekompasset")
        if word_lower.len() >= 10 {
            if let Some(analyzer) = &self.analyzer {
                let is_known_or_compound = |w: &str| -> bool {
                    if analyzer.has_word(w) { return true; }
                    for j in 3..w.len().saturating_sub(2) {
                        if !w.is_char_boundary(j) { continue; }
                        let left = &w[..j];
                        let right = &w[j..];
                        if right.len() >= 3 && analyzer.has_word(left) && analyzer.has_word(right) { return true; }
                        if right.starts_with('s') && right.len() > 3 && analyzer.has_word(left) && analyzer.has_word(&right[1..]) { return true; }
                    }
                    false
                };
                for strip in 1..=2usize {
                    if word_lower.is_char_boundary(strip) {
                        let trimmed = &word_lower[strip..];
                        if trimmed.len() >= 5 && is_known_or_compound(trimmed) && seen.insert(trimmed.to_string()) {
                            edit_distances.insert(trimmed.to_string(), strip as u32);
                            candidates.push(trimmed.to_string());
                        }
                    }
                    let end = word_lower.len() - strip;
                    if word_lower.is_char_boundary(end) {
                        let trimmed = &word_lower[..end];
                        if trimmed.len() >= 5 && is_known_or_compound(trimmed) && seen.insert(trimmed.to_string()) {
                            edit_distances.insert(trimmed.to_string(), strip as u32);
                            candidates.push(trimmed.to_string());
                        }
                    }
                }
            }
        }

        // Source 10: First-character swap — try replacing first char with every letter
        // Catches "sjøkken" → "kjøkken" where first char is wrong but rest is correct
        if word_lower.len() >= 3 {
            if let Some(analyzer) = &self.analyzer {
                let rest = &word_lower[word_first.len_utf8()..];
                for c in "abcdefghijklmnopqrstuvwxyzæøå".chars() {
                    if c == word_first { continue; }
                    let candidate = format!("{}{}", c, rest);
                    if analyzer.has_word(&candidate) && seen.insert(candidate.clone()) {
                        edit_distances.insert(candidate.clone(), 1);
                        candidates.push(candidate);
                    }
                }
            }
        }

        // Source 11: Inflected forms of candidates — for each candidate lemma,
        // add its inflections so BERT can pick the grammatically correct form.
        // "kjøkken" → also adds "kjøkkenet", "kjøkkenene" etc.
        {
            use mtag::types::{Pos, Tag};
            if let Some(analyzer) = &self.analyzer {
                let base_candidates: Vec<String> = candidates.clone();
                for base in &base_candidates {
                    if let Some(readings) = analyzer.dict_lookup(base) {
                        for r in &readings {
                            if !matches!(r.pos, Pos::Subst) { continue; }
                            // Add definite and plural forms
                            for tag in &[Tag::Be, Tag::Fl] {
                                let forms = analyzer.forms_for_lemma(&r.lemma, &Pos::Subst, tag);
                                for form in forms {
                                    let fl = form.to_lowercase();
                                    if fl != word_lower && fl.len() >= 2 && seen.insert(fl.clone()) {
                                        let dist = levenshtein_distance(&word_lower, &fl);
                                        if dist <= 4 {
                                            edit_distances.insert(fl.clone(), dist);
                                            candidates.push(fl);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        // Source 13: Compound walker — decompose misspelled compound words
        // Only for words ≥ 7 chars (compounds are long)
        let mut compound_candidates: std::collections::HashSet<String> = std::collections::HashSet::new();
        if word_lower.len() >= 7 {
            if let (Some(fst), Some(analyzer)) = (&self.compound_fst, &self.analyzer) {
                let word_check = |w: &str| -> bool {
                    analyzer.dict_lookup(w).map_or(false, |rs|
                        rs.iter().any(|r| r.pos != mtag::types::Pos::Prop))
                };
                let noun_check = |w: &str| -> bool {
                    analyzer.dict_lookup(w).map_or(false, |rs| {
                        let n = rs.iter().filter(|r| r.pos == mtag::types::Pos::Subst).count();
                        let a = rs.iter().filter(|r| r.pos == mtag::types::Pos::Adj).count();
                        n > 0 && n >= a
                    })
                };
                let results = compound_walker::compound_fuzzy_walk(
                    fst, &word_lower,
                    &*self.language,
                    self.wordfreq.as_deref(),
                    Some(&word_check), Some(&noun_check),
                );
                for r in results.iter().take(10) {
                    let cw = r.compound_word.to_lowercase();
                    if seen.insert(cw.clone()) {
                        edit_distances.insert(cw.clone(), r.total_edits);
                        compound_candidates.insert(cw.clone());
                        candidates.push(cw);
                    }
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
            // Boost words found in document or user dictionary (TF-IDF style)
            ortho_sim *= compute_boost(w, &self.doc_word_counts,
                self.user_dict.as_ref(), self.wordfreq.as_deref(), &*self.language);
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
                if checked >= 15 { break; }

                // Skip hyphenated candidates when misspelled word has no hyphen
                if !word_lower.contains('-') && candidate.contains('-') {
                    continue;
                }

                // Compound walker candidates are pre-validated (each part checked
                // against wordfreq + analyzer + noun). Skip dictionary filter for them.
                if !compound_candidates.contains(candidate) {
                    // Dictionary check: every word must exist in standard or user dict
                    let ud = &self.user_dict;
                    let words: Vec<&str> = candidate.split_whitespace().collect();
                    if words.iter().any(|w| {
                        !analyzer.has_word(w)
                        && !ud.as_ref().map_or(false, |u| u.has_word(w))
                    }) {
                        continue;
                    }
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
                // If the user is at a word boundary (pressed space or end-of-
                // sentence), context_after is empty/whitespace — which breaks
                // score_and_rerank's word-position extraction. Append sentinel
                // "." to BOTH context_after and sentence so the word extraction
                // inside the scorer hits the else branch and correctly strips
                // the trailing punctuation to recover the misspelled word.
                let (context_after, sentence) = if context_after.trim().is_empty() {
                    (".".to_string(), format!("{}.", sentence_ctx))
                } else {
                    (context_after, sentence_ctx.to_string())
                };
                let request_id = worker.send(|id| bert_worker::BertRequest::SpellingScore { id, context_before, context_after, candidates, sentence });
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
                            candidates.push((removed_aa.clone(), self.language.ui_removed_aa_before(&e.word), e.rule_name.clone()));
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
                    candidates.push((removed, self.language.ui_removed_word(&e.word), e.rule_name.clone()));
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
                // Grammar rerank path: we don't have per-word context_before/after
                // (the grammar check operates on the whole sentence), so pass
                // the full sentence as-is and let the scorer work with it.
                let sentence_full = sentence.to_string();
                let request_id = worker.send(|id| bert_worker::BertRequest::SpellingScore {
                    id,
                    context_before: String::new(),
                    context_after: String::new(),
                    candidates,
                    sentence: sentence_full,
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

    /// Rebuild document word frequency map (cached — skips if text unchanged).
    fn rebuild_doc_word_counts(&mut self) {
        // Get text source — prefer paragraph_texts, fall back to last_doc_text
        let text = if !self.paragraph_texts.is_empty() {
            self.paragraph_texts.values().cloned().collect::<Vec<_>>().join(" ")
        } else if !self.last_doc_text.is_empty() {
            self.last_doc_text.clone()
        } else {
            return; // no text — keep existing counts (don't clear)
        };
        // Quick hash to avoid redundant rebuilds
        let current_hash = hash_str(&text);
        if current_hash == self.doc_word_counts_hash {
            return;
        }
        self.doc_word_counts.clear();
        for word in text.split(|c: char| !c.is_alphanumeric() && c != '-') {
            if word.len() < 2 { continue; }
            let lower = word.to_lowercase();
            *self.doc_word_counts.entry(lower).or_insert(0) += 1;
        }
        self.doc_word_counts_hash = current_hash;
    }

    /// Remove errors whose word has been corrected in the document.
    fn prune_resolved_errors(&mut self) {
        // Build per-paragraph text lookup
        let para_texts_lower: std::collections::HashMap<&str, String> = self.paragraph_texts.iter()
            .map(|(k, v)| (k.as_str(), v.to_lowercase()))
            .collect();
        // Build doc text only from known paragraphs
        let doc_text = if !para_texts_lower.is_empty() {
            para_texts_lower.values().cloned().collect::<Vec<_>>().join(" ")
        } else if !self.last_doc_text.is_empty() {
            self.last_doc_text.to_lowercase()
        } else {
            return;
        };
        // Clear underlines for errors that will be removed
        for e in &mut self.writing_errors {
            let should_remove = if e.ignored {
                true
            } else if !e.paragraph_id.is_empty() && !para_texts_lower.contains_key(e.paragraph_id.as_str()) {
                // Error's paragraph no longer in cache — paragraph was deleted or document changed
                true
            } else {
                // Check within the error's own paragraph text (not full doc)
                let check_text = if !e.paragraph_id.is_empty() {
                    para_texts_lower.get(e.paragraph_id.as_str()).map(|s| s.as_str()).unwrap_or("")
                } else {
                    doc_text.as_str()
                };
                match e.category {
                    ErrorCategory::Grammar => !check_text.contains(&e.sentence_context.to_lowercase()),
                    ErrorCategory::Spelling => {
                        let word_lower = e.word.to_lowercase();
                        !check_text.split(|c: char| !c.is_alphanumeric()).any(|w| w == word_lower)
                    }
                    ErrorCategory::SentenceBoundary => !check_text.contains(&e.word.to_lowercase()),
                }
            };
            if should_remove && e.underlined {
                if !e.paragraph_id.is_empty() {
                    self.manager.clear_underline_word(&e.word, &e.paragraph_id);
                }
                if e.word_doc_start < e.word_doc_end {
                    self.manager.clear_error_underline(e.word_doc_start, e.word_doc_end);
                }
                e.underlined = false;
            }
        }
        self.writing_errors.retain(|e| {
            if e.ignored {
                log!("Pruning ignored: {:?} '{}'", e.category, trunc(&e.word, 40));
                return false;
            }
            if !e.paragraph_id.is_empty() && !para_texts_lower.contains_key(e.paragraph_id.as_str()) {
                log!("Error resolved: {:?} '{}' paragraph gone", e.category, trunc(&e.word, 40));
                return false;
            }
            let check_text = if !e.paragraph_id.is_empty() {
                para_texts_lower.get(e.paragraph_id.as_str()).map(|s| s.as_str()).unwrap_or("")
            } else {
                doc_text.as_str()
            };
            let still_present = match e.category {
                ErrorCategory::Grammar => check_text.contains(&e.sentence_context.to_lowercase()),
                ErrorCategory::Spelling => {
                    let word_lower = e.word.to_lowercase();
                    check_text.split(|c: char| !c.is_alphanumeric()).any(|w| w == word_lower)
                }
                ErrorCategory::SentenceBoundary => check_text.contains(&e.word.to_lowercase()),
            };
            if !still_present {
                log!("Error resolved: {:?} '{}' no longer in paragraph", e.category, trunc(&e.word, 40));
            }
            still_present
        });
    }

    /// Incremental paragraph-based document scanning using ParaID.
    /// Only processes paragraphs whose text changed since last scan.
    fn update_grammar_errors_incremental(&mut self, paragraphs: Vec<(String, String, usize)>) {
        let current_ids: std::collections::HashSet<String> = paragraphs.iter().map(|(id, _, _)| id.clone()).collect();

        // Detect deleted paragraphs — remove their errors and sentence hashes
        let old_ids: Vec<String> = self.paragraph_texts.keys().cloned().collect();
        for old_id in &old_ids {
            if !current_ids.contains(old_id) {
                log!("Para deleted: {}", old_id);
                self.paragraph_texts.remove(old_id);
                if let Some(hashes) = self.paragraph_sentence_hashes.remove(old_id) {
                    for h in &hashes {
                        self.processed_sentence_hashes.remove(h);
                    }
                }
                // Remove errors for this paragraph
                self.writing_errors.retain(|e| e.paragraph_id != *old_id);
            }
        }

        // Process each paragraph — only if text changed
        for (para_id, text, char_start) in &paragraphs {
            let text_trimmed = text.trim();
            if text_trimmed.is_empty() { continue; }

            // Check if paragraph text changed
            let changed = match self.paragraph_texts.get(para_id) {
                Some(old_text) => old_text != text_trimmed,
                None => true, // new paragraph
            };
            if !changed { continue; }

            log!("Para changed: id={} offset={} text='{}'", para_id, char_start, trunc(text_trimmed, 50));
            self.paragraph_texts.insert(para_id.clone(), text_trimmed.to_string());

            // Clear old sentence hashes for this paragraph
            if let Some(old_hashes) = self.paragraph_sentence_hashes.remove(para_id) {
                for h in &old_hashes {
                    self.processed_sentence_hashes.remove(h);
                }
            }

            // Remove stale errors for this paragraph
            self.writing_errors.retain(|e| {
                if e.paragraph_id == *para_id { return false; }
                // Also remove errors whose sentence is no longer in this paragraph
                if e.doc_offset >= *char_start && e.doc_offset < char_start + text.len() {
                    return false;
                }
                true
            });

            // Split into sentences and queue for grammar checking
            let sentences = split_sentences(text_trimmed);
            let mut new_hashes = Vec::new();
            for sent in &sentences {
                let sent_trimmed = sent.trim();
                if sent_trimmed.is_empty() { continue; }
                let sent_h = hash_str(&format!("{}|{}", para_id, sent_trimmed));
                new_hashes.push(sent_h);

                if self.processed_sentence_hashes.contains(&sent_h) {
                    continue; // already checked and clean
                }

                // Queue for grammar checking
                self.grammar_queue.push((sent_trimmed.to_string(), *char_start));

                // Queue words for spelling
                let mut word_pos = *char_start;
                for word in sent_trimmed.split_whitespace() {
                    let clean = word.trim_matches(|c: char| c.is_ascii_punctuation() || c == '\u{00ab}' || c == '\u{00bb}');
                    if !clean.is_empty() && clean.len() >= 2 {
                        self.spelling_queue.push(SpellingQueueItem {
                            word: clean.to_string(),
                            sentence_ctx: sent_trimmed.to_string(),
                            paragraph_id: para_id.clone(),
                        });
                    }
                    word_pos += word.len() + 1;
                }
            }
            self.paragraph_sentence_hashes.insert(para_id.clone(), new_hashes);
        }

        // Update full doc text for other consumers (completion context etc.)
        let full_text = paragraphs.iter().map(|(_, text, _)| text.as_str()).collect::<Vec<_>>().join(" ");
        if !full_text.is_empty() {
            self.last_doc_text = full_text;
        }

        if !self.grammar_queue.is_empty() {
            self.grammar_scanning = true;
            self.grammar_queue_total = self.grammar_queue.len();
            self.process_grammar_queue();
        }
    }

    /// Process a single paragraph read via COM — mirrors process_addin_changed_paragraphs for Mac.
    /// Called on each keystroke with the paragraph at cursor. Only reprocesses if text changed.
    fn process_com_changed_paragraph(&mut self, para_id: String, text: String, char_start: usize) {
        // Clean control characters
        let clean_text: String = text.chars()
            .map(|c| if c.is_control() && c != '\n' && c != '\r' && c != '\t' { ' ' } else { c })
            .collect();

        // Evict stale paragraph entries: if a NEW para_id appears at a position where
        // we had a DIFFERENT para_id cached, the old one is from a previous document state.
        // This handles document restore/reload where all paragraph IDs change.
        if !self.paragraph_texts.contains_key(&para_id) {
            // Check if any existing entry overlaps this char_start position
            // (paragraph at same position but different ID = document reloaded)
            let para_char_end = char_start + clean_text.chars().count();
            let stale: Vec<String> = self.writing_errors.iter()
                .filter(|e| e.paragraph_id != para_id && e.doc_offset >= char_start && e.doc_offset < para_char_end)
                .map(|e| e.paragraph_id.clone())
                .collect::<std::collections::HashSet<_>>()
                .into_iter().collect();
            for id in &stale {
                log!("Evicting stale paragraph {} (replaced by {} at offset {})", trunc(id, 10), trunc(&para_id, 10), char_start);
                self.paragraph_texts.remove(id);
                if let Some(hashes) = self.paragraph_sentence_hashes.remove(id) {
                    for h in hashes { self.processed_sentence_hashes.remove(&h); }
                }
                self.writing_errors.retain(|e| e.paragraph_id != *id);
            }
        }

        // Skip if text identical to cached
        if self.paragraph_texts.get(&para_id).map_or(false, |t| t == &clean_text) {
            return;
        }

        log!("COM paragraph changed: '{}' (para={} start={})", trunc(&clean_text, 50), trunc(&para_id, 10), char_start);
        self.paragraph_texts.insert(para_id.clone(), clean_text.clone());

        // Clear all underlines in this paragraph range — will be re-applied for remaining errors
        let para_char_end = char_start + clean_text.chars().count();
        self.manager.clear_error_underline(char_start, para_char_end);
        // Mark all errors in this paragraph as not underlined so they get re-applied
        for e in &mut self.writing_errors {
            if e.paragraph_id == para_id && e.underlined {
                e.underlined = false;
            }
        }

        // Split into sentences
        let sentences = split_sentences(&clean_text);
        let new_hashes: Vec<u64> = sentences.iter()
            .map(|s| hash_str(&format!("{}|{}", para_id, s))).collect();

        // Remove old sentence hashes for this paragraph
        if let Some(old_hashes) = self.paragraph_sentence_hashes.get(&para_id) {
            for old_h in old_hashes {
                if !new_hashes.contains(old_h) {
                    self.processed_sentence_hashes.remove(old_h);
                }
            }
        }

        // Clear stale errors for this paragraph — also clear their underlines
        let new_sentence_set: std::collections::HashSet<String> = sentences.iter().map(|s| s.to_lowercase()).collect();
        let para_text_lower = clean_text.to_lowercase();
        let mut to_clear: Vec<(usize, usize)> = Vec::new();
        self.writing_errors.retain(|e| {
            if e.paragraph_id != para_id { return true; }
            if new_sentence_set.contains(&e.sentence_context.to_lowercase()) { return true; }
            if matches!(e.category, ErrorCategory::Spelling) {
                if para_text_lower.contains(&e.word.to_lowercase()) { return true; }
            }
            // Removing this error — clear its underline
            if e.underlined && e.word_doc_start < e.word_doc_end {
                to_clear.push((e.word_doc_start, e.word_doc_end));
            }
            false
        });
        for (start, end) in &to_clear {
            log!("Clearing underline {}..{} (error removed from paragraph)", start, end);
            self.manager.clear_error_underline(*start, *end);
        }

        // Send new/changed sentences to grammar actor
        // Compute per-sentence offsets within the paragraph
        let para_ends_with_boundary = clean_text.ends_with(' ')
            || clean_text.ends_with('.') || clean_text.ends_with('!')
            || clean_text.ends_with('?') || clean_text.ends_with(':')
            || clean_text.ends_with('\t');
        if let Some(actor) = &self.grammar_actor {
            let clean_lower = clean_text.to_lowercase();
            let mut search_from = 0usize;
            for sentence_text in &sentences {
                // Find sentence position within paragraph text
                let sent_lower = sentence_text.to_lowercase();
                let sent_offset = clean_lower[search_from..].find(&sent_lower)
                    .map(|pos| search_from + pos)
                    .unwrap_or(search_from);
                let doc_offset = char_start + clean_text[..sent_offset.min(clean_text.len())].chars().count();
                search_from = sent_offset + sent_lower.len();

                let sent_h = hash_str(&format!("{}|{}", para_id, sentence_text));
                let is_complete = sentence_text.ends_with('.') || sentence_text.ends_with('!')
                    || sentence_text.ends_with('?') || sentence_text.ends_with(':');

                if self.processed_sentence_hashes.contains(&sent_h)
                    || self.grammar_inflight.contains(&sent_h) {
                    continue;
                }

                // Clear errors for this changed sentence
                let sentence_lower = sentence_text.to_lowercase();
                self.writing_errors.retain(|e| {
                    !(e.paragraph_id == para_id && e.sentence_context.to_lowercase() == sentence_lower)
                });

                if is_complete || para_ends_with_boundary {
                    let uw = self.user_dict.as_ref().map_or(vec![], |ud| ud.list_words());
                    actor.check_sentence_with_doc(sentence_text, doc_offset, &para_id, 0, &self.last_doc_text, &uw);
                    self.grammar_inflight.insert(sent_h);
                    log!("Grammar send (COM): '{}' (para={} doc_off={})", trunc(sentence_text, 50), trunc(&para_id, 10), doc_offset);
                }
            }
        }

        // Store new sentence hashes
        self.paragraph_sentence_hashes.insert(para_id.clone(), new_hashes);

        // COM mode: last_doc_text = current paragraph only.
        // paragraph_texts accumulates all visited paragraphs for error tracking/pruning,
        // but feeding them all into last_doc_text would cause update_grammar_errors()
        // to rescan the entire document on every keystroke instead of just the active paragraph.
        self.last_doc_text = clean_text.clone();

        // Rebuild word counts + prune
        self.rebuild_doc_word_counts();
        self.prune_resolved_errors();
    }

    /// Prepare grammar scan: read document, split sentences, compute offsets, fill queue.
    /// This is fast (no SWI/BERT calls) and runs every poll when document changes.
    /// The actual per-sentence grammar checking happens incrementally in process_grammar_queue().
    fn update_grammar_errors(&mut self) {
        // When Word Add-in is active, sentence detection and error management
        // is handled by the add-in (process_addin_changed_paragraphs).
        if self.manager.bridges.iter().any(|b| b.name() == "Word Add-in") {
            return;
        }

        // Word COM mode: grammar checking is handled paragraph-by-paragraph in
        // process_com_changed_paragraph. paragraph_texts is non-empty when COM is (or was)
        // active — skip the full-doc scan so we never reprocess stale accumulated text.
        if !self.paragraph_texts.is_empty() && !self.manager.last_user_was_browser {
            return;
        }

        // DISABLED: incremental scanning causes freeze.
        // Fall through to original full-doc path below.
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
            self.manager.clear_all_error_underlines();
            self.writing_errors.clear();
            self.spelling_queue.clear();
            self.pending_spelling_bert.clear();
            self.grammar_queue.clear();
            self.grammar_queue_total = 0;
            ; self.processed_sentence_hashes.clear();
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
                word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: String::new(), error_word: String::new(),
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
        // DISABLED: check_spelling runs find_spelling_suggestions on main thread = freeze
        // Grammar actor handles spelling detection via unknown words.
        return;
        #[allow(unreachable_code)]
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
                log!("Reset: clearing errors + underlines (keeping sentence hashes)");
                self.manager.clear_all_error_underlines();
                // Collect hashes of sentences that HAD errors — these need re-checking
                let error_hashes: std::collections::HashSet<u64> = self.writing_errors.iter()
                    .map(|e| hash_str(&format!("{}|{}", e.paragraph_id, e.sentence_context)))
                    .collect();
                self.writing_errors.clear();
                // Only remove hashes for sentences that had errors — forces re-check
                // Sentences that were clean stay "processed" (no need to re-check)
                for h in &error_hashes {
                    self.processed_sentence_hashes.remove(h);
                }
                self.paragraph_sentence_hashes.clear();
                self.paragraph_texts.clear();
                self.last_doc_text.clear();
                self.grammar_inflight.clear();
                self.llm_checked_hashes.clear();
                self.llm_waiting = false;
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
                // Skip if text is identical to last time (add-in sends duplicates)
                if self.paragraph_texts.get(&p.paragraph_id).map_or(false, |t| t == &p.text) {
                    continue;
                }

                log!("Addin changed paragraph: '{}' (para={} cursor={:?})", trunc(&p.text, 50), trunc(&p.paragraph_id, 10), p.cursor_start);
                self.paragraph_texts.insert(p.paragraph_id.clone(), p.text.clone());

                // Derive cursor context from paragraph text for fast suggestion triggering
                // (don't wait for the slow /context POST — use paragraph data directly)
                if let Some(_cursor_abs) = p.cursor_start {
                    let text = &p.text;
                    // Extract last word (what user is typing) — everything after last whitespace
                    let last_word = text.rsplit(|c: char| c.is_whitespace())
                        .next().unwrap_or("").trim_matches(|c: char| c.is_ascii_punctuation() && c != '-');
                    // Extract sentence context — text after last sentence-ending punctuation
                    let sent_start = text.rfind(|c: char| ".!?:".contains(c))
                        .map(|i| {
                            // Skip whitespace after punctuation
                            let after = i + 1;
                            text[after..].find(|c: char| !c.is_whitespace()).map(|j| after + j).unwrap_or(after)
                        })
                        .unwrap_or(0);
                    let sentence = text[sent_start..].trim();
                    // Build masked sentence for BERT: sentence with <mask> replacing the word
                    let masked = if !last_word.is_empty() && sentence.ends_with(last_word) {
                        let before = sentence[..sentence.len() - last_word.len()].trim_end();
                        if before.is_empty() { "<mask>".to_string() } else { format!("{} <mask>", before) }
                    } else {
                        format!("{} <mask>", sentence)
                    };
                    let new_word = last_word.to_string();
                    let new_sentence = sentence.to_string();
                    // Only update context if word changed (avoids redundant completion dispatches)
                    if new_word != self.context.word || new_sentence != self.context.sentence {
                        self.context.word = new_word;
                        self.context.sentence = new_sentence;
                        self.context.masked_sentence = Some(masked);
                        self.context.paragraph_id = p.paragraph_id.clone();
                        self.context.cursor_doc_offset = p.cursor_start;
                        self.last_prefix_change = Instant::now();
                        self.pending_completion = true;
                    }
                }

                // Strip control characters (vertical tab etc.) — now properly decoded by JSON parser
                let clean_text: String = p.text.chars()
                    .map(|c| if c.is_control() && c != '\n' && c != '\r' && c != '\t' { ' ' } else { c })
                    .collect();

                // Extract email parts to skip in spelling checks
                let email_skip_words: std::collections::HashSet<String> = if clean_text.contains('@') {
                    clean_text.split_whitespace()
                        .map(|t| t.trim_matches(|c: char| c == '(' || c == ')' || c == ',' || c == ';'))
                        .filter(|t| t.contains('@'))
                        .flat_map(|email| email.split(|c: char| c == '@' || c == '.' || c == '-' || c == '_')
                            .filter(|p| p.len() >= 2 && p.chars().all(|c| c.is_alphanumeric()))
                            .map(|p| p.to_lowercase()))
                        .collect()
                } else {
                    std::collections::HashSet::new()
                };
                if !email_skip_words.is_empty() {
                    log!("  Email skip words: {:?}", email_skip_words);
                }

                // Split paragraph into sentences
                let sentences = split_sentences(&clean_text);
                let new_hashes: Vec<u64> = sentences.iter()
                    .map(|s| hash_str(&format!("{}|{}", p.paragraph_id, s))).collect();

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
                let para_text_lower = sentences.iter().map(|s| s.to_lowercase()).collect::<Vec<_>>().join(" ");
                let before_count = self.writing_errors.len();
                // When a word is gone from the paragraph (user fixed it),
                // its underline disappears with the word. No explicit clear needed.
                self.writing_errors.retain(|e| {
                    if e.paragraph_id != p.paragraph_id { return true; }
                    // Exact sentence match — keep
                    let e_sent_lower = e.sentence_context.to_lowercase();
                    if new_sentence_set.contains(&e_sent_lower) { return true; }
                    // For spelling errors: keep if the misspelled word is still in the paragraph
                    if matches!(e.category, ErrorCategory::Spelling) {
                        let word_lower = e.word.to_lowercase();
                        if para_text_lower.contains(&word_lower) { return true; }
                    }
                    log!("  Removing stale error: word='{}' sentence='{}' (not in set: {:?})",
                        e.word, trunc(&e.sentence_context, 60),
                        new_sentence_set.iter().take(3).map(|s| trunc(s, 40)).collect::<Vec<_>>());
                    false
                });
                if self.writing_errors.len() < before_count {
                    log!("  Cleared {} stale errors for para={}", before_count - self.writing_errors.len(), trunc(&p.paragraph_id, 10));
                }

                // Check each sentence: skip if already processed (hash unchanged)
                for sentence_text in &sentences {
                    let sent_h = hash_str(&format!("{}|{}", p.paragraph_id, sentence_text));

                    let is_complete = sentence_text.ends_with('.') || sentence_text.ends_with('!')
                        || sentence_text.ends_with('?') || sentence_text.ends_with(':');

                    if self.processed_sentence_hashes.contains(&sent_h)
                        || self.grammar_inflight.contains(&sent_h) {
                        continue; // Already processed or in-flight, skip
                    }

                    // This sentence is new or changed — clear old errors (underlines stay if word still exists)
                    let sentence_lower = sentence_text.to_lowercase();
                    self.writing_errors.retain(|e| {
                        !(e.paragraph_id == p.paragraph_id && e.sentence_context.to_lowercase() == sentence_lower)
                    });

                    // Send to grammar actor for spelling + grammar checking.
                    let first_seen = !self.paragraph_sentence_hashes.contains_key(&p.paragraph_id);
                    if is_complete || first_seen {
                        let mut uw = self.user_dict.as_ref().map_or(vec![], |ud| ud.list_words());
                        uw.extend(email_skip_words.iter().cloned());
                        actor.check_sentence_with_doc(sentence_text, 0, &p.paragraph_id, 0, &self.last_doc_text, &uw);
                        self.grammar_inflight.insert(sent_h);
                    }

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
                for e in &self.writing_errors {
                    if e.paragraph_id == para_id && e.underlined {
                        self.manager.clear_underline_word(&e.word, &e.paragraph_id);
                    }
                }
                self.writing_errors.retain(|e| e.paragraph_id != para_id);
                if self.writing_errors.len() < before {
                    log!("Cleared {} errors for deleted para={}", before - self.writing_errors.len(), trunc(&para_id, 10));
                }
                self.paragraph_texts.remove(&para_id);
                // Remove sentence hashes for deleted paragraph
                if let Some(hashes) = self.paragraph_sentence_hashes.remove(&para_id) {
                    for h in hashes {
                        self.processed_sentence_hashes.remove(&h);
                    }
                }
            }
        }

        // Clean stale paragraph_texts entries (paragraphs that were deleted/merged)
        let active_para_ids: std::collections::HashSet<&String> = self.paragraph_sentence_hashes.keys().collect();
        self.paragraph_texts.retain(|k, _| active_para_ids.contains(k));

        // Update last_doc_text from accumulated paragraph texts (for prune_resolved_errors)
        if !self.paragraph_texts.is_empty() {
            self.last_doc_text = self.paragraph_texts.values().cloned().collect::<Vec<_>>().join(" ");
        }

        // Rebuild document word counts for suggestion boosting
        self.rebuild_doc_word_counts();

        // Prune errors whose word is no longer in the document (e.g., after cut)
        self.prune_resolved_errors();

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

        // Send up to 3 queued sentences per frame to avoid flooding the actor
        let mut sent_count = 0;
        while sent_count < 3 {
            let (trimmed, doc_offset) = match self.grammar_queue.first().cloned() {
                Some(v) => v,
                None => break,
            };
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
            let uw = self.user_dict.as_ref().map_or(vec![], |ud| ud.list_words());
            actor.check_sentence_with_doc(&trimmed, doc_offset, "", 0, &self.last_doc_text, &uw);
            sent_count += 1;
        }
        if self.grammar_queue.is_empty() {
            self.grammar_scanning = false;
        }
    }

    /// Send all current error sentences to LLM for AI correction (button-triggered).
    fn dispatch_llm_fix_all(&mut self) {
        if self.llm_actor.is_none() || self.llm_waiting { return; }

        // Collect unique sentences from current errors
        let mut seen = std::collections::HashSet::new();
        let mut batch: Vec<(String, String)> = Vec::new();
        let mut hashes: Vec<u64> = Vec::new();
        for e in &self.writing_errors {
            if e.ignored || e.sentence_context.is_empty() { continue; }
            if e.rule_name == "llm_correction" { continue; } // already LLM-corrected
            let llm_hash = hash_str(&format!("llm|{}|{}", e.paragraph_id, e.sentence_context));
            if self.llm_checked_hashes.contains(&llm_hash) { continue; }
            if !seen.insert(llm_hash) { continue; }
            batch.push((e.sentence_context.clone(), e.paragraph_id.clone()));
            hashes.push(llm_hash);
        }

        if batch.is_empty() { return; }

        // Rate limit: 300 sentences/hour
        self.llm_sent_count.retain(|t| t.elapsed() < Duration::from_secs(3600));
        let batch_size = batch.len().min(10);
        if self.llm_sent_count.len() + batch_size > 300 {
            log!("LLM rate limit: {}/300 sentences this hour", self.llm_sent_count.len());
            return;
        }
        let batch: Vec<_> = batch.into_iter().take(batch_size).collect();
        let hashes: Vec<_> = hashes.into_iter().take(batch_size).collect();
        for _ in 0..batch_size { self.llm_sent_count.push(Instant::now()); }

        log!("LLM fix-all: {} sentences (rate: {}/300)", batch_size, self.llm_sent_count.len());
        if let Some(actor) = &mut self.llm_actor {
            actor.send(batch, hashes);
            self.llm_waiting = true;
            self.llm_waiting_since = Instant::now();
        }
    }

    /// Poll LLM actor for correction results.
    fn poll_llm_responses(&mut self) {
        let actor = match &self.llm_actor { Some(a) => a, None => return };
        while let Some(resp) = actor.try_recv() {
            for c in &resp.corrections {
                if c.corrected == c.original { continue; }

                // Build summary and find first changed word
                let error_word = if let Some((from, _, _)) = c.changes.first() {
                    from.clone()
                } else {
                    find_diff_word(&c.original, &c.corrected)
                };
                let explanation = if c.changes.len() == 1 {
                    let (from, to, why) = &c.changes[0];
                    if why.is_empty() { format!("«{}» → «{}»", from, to) }
                    else { format!("«{}» → «{}»: {}", from, to, why) }
                } else if c.changes.len() > 1 {
                    format!("Flere endringer ({})", c.changes.len())
                } else {
                    format!("AI: «{}» → «{}»", error_word, c.corrected)
                };
                log!("LLM correction: '{}' changes={} para={}", explanation, c.changes.len(), trunc(&c.paragraph_id, 10));

                // Remove ALL existing local errors for this sentence (LLM replaces them)
                // Clear underlines for removed errors
                for e in &self.writing_errors {
                    if e.paragraph_id == c.paragraph_id
                        && e.sentence_context.to_lowercase() == c.original.to_lowercase()
                        && e.underlined
                    {
                        let w = if !e.error_word.is_empty() { &e.error_word } else { &e.word };
                        self.manager.clear_underline_word(w, &e.paragraph_id);
                    }
                }
                self.writing_errors.retain(|e| {
                    !(e.paragraph_id == c.paragraph_id
                        && e.sentence_context.to_lowercase() == c.original.to_lowercase())
                });

                // Encode changes as JSON in explanation for 💡 popup
                let changes_json = serde_json::to_string(&c.changes.iter()
                    .map(|(f, t, w)| serde_json::json!({"from": f, "to": t, "why": w}))
                    .collect::<Vec<_>>()).unwrap_or_default();
                let full_explanation = format!("LLM_CHANGES:{}\n{}", changes_json, explanation);

                // Add LLM correction
                self.writing_errors.push(WritingError {
                    category: ErrorCategory::Grammar,
                    word: c.original.clone(),
                    suggestion: c.corrected.clone(),
                    explanation: full_explanation,
                    rule_name: "llm_correction".to_string(),
                    sentence_context: c.original.clone(),
                    doc_offset: 0,
                    position: 0,
                    ignored: false,
                    word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false,
                    paragraph_id: c.paragraph_id.clone(),
                    error_word: error_word.clone(),
                });

                for b in &self.manager.bridges {
                    b.underline_word(&error_word, &c.paragraph_id, "#0000FF");
                }
            }
            for h in &resp.checked_hashes {
                self.llm_checked_hashes.insert(*h);
            }
            self.llm_waiting = false;
        }
    }

    /// Poll grammar actor for results and create WritingErrors.
    fn poll_grammar_responses(&mut self) {
        let actor = match &self.grammar_actor {
            Some(a) => a,
            None => return,
        };

        while let Some(resp) = actor.try_recv() {
            // Discard grammar results while browser is foreground — they belong
            // to a previous Word session and must not bleed into browser mode.
            if self.suppress_errors {
                log!("Grammar response discarded (browser foreground): para='{}'", trunc(&resp.paragraph_id, 10));
                continue;
            }
            log!("Grammar response: sentence='{}' errors={} unknown={} para='{}'",
                trunc(&resp.sentence, 40), resp.errors.len(), resp.unknown_words.len(),
                trunc(&resp.paragraph_id, 10));
            let sent_h = if resp.paragraph_id.is_empty() {
                hash_str(&resp.sentence)
            } else {
                hash_str(&format!("{}|{}", resp.paragraph_id, resp.sentence))
            };

            // Guard: discard if the paragraph is no longer tracked (app switched and
            // paragraph_sentence_hashes was cleared) OR if the sentence hash no longer
            // matches the current paragraph content (user edited the text).
            if !resp.paragraph_id.is_empty() {
                let is_stale = match self.paragraph_sentence_hashes.get(&resp.paragraph_id) {
                    Some(current_hashes) => !current_hashes.contains(&sent_h),
                    None => true, // paragraph cleared on app switch → stale
                };
                if is_stale {
                    log!("Stale grammar response discarded: sentence no longer in para={}", trunc(&resp.paragraph_id, 10));
                    continue;
                }
            }

            self.grammar_inflight.remove(&sent_h);
            self.processed_sentence_hashes.insert(sent_h); // Mark sentence as processed

            // Handle grammar errors — only for complete sentences (ends with punctuation)
            let sentence_complete = resp.sentence.ends_with('.') || resp.sentence.ends_with('!')
                || resp.sentence.ends_with('?') || resp.sentence.ends_with(':');
            if !resp.errors.is_empty() && sentence_complete {
                for ge in &resp.errors {
                    log!("  Grammar error: '{}' → '{}' ({})", ge.word, ge.suggestion, ge.rule_name);
                }

                // Sentinel that grammar rules can use as their suggestion to mean
                // "delete this word entirely" — handled below by routing through
                // remove_word_from_sentence instead of replace_word_at_position.
                // Used by Prolog rules like infinitivsmerke_presens where the
                // correction is to remove a stray «å», not to replace a word.
                const DELETE_SENTINEL: &str = "<DELETE>";

                let errors_with_suggestions: Vec<_> = resp.errors.iter()
                    .filter(|e| !e.suggestion.is_empty())
                    .collect();

                if !errors_with_suggestions.is_empty() {
                    for (i, ge) in errors_with_suggestions.iter().enumerate() {
                        let first_alt = ge.suggestion.split('|').next().unwrap_or(&ge.suggestion);
                        let mut corrected = if first_alt == DELETE_SENTINEL {
                            remove_word_from_sentence(&resp.sentence, &ge.word)
                        } else {
                            replace_word_at_position(&resp.sentence, &ge.word, first_alt)
                        };
                        // When article gender changes (en↔et), also fix adjective agreement.
                        // Dispatched through the language trait so the actual articles
                        // and rules come from the active LanguageBundle (Bokmål here;
                        // Nynorsk and other languages override the trait method).
                        if ge.rule_name.contains("kjoenn") || ge.rule_name.contains("kjønn") {
                            corrected = self.language.fix_adjective_agreement(&corrected);
                        }
                        if corrected.trim() == resp.sentence.trim() {
                            continue;
                        }
                        // For user-facing text, render <DELETE> as the empty string
                        // (or "" with quotes) so the explanation reads cleanly.
                        let display_alt = if first_alt == DELETE_SENTINEL { "" } else { first_alt };
                        log!("  Grammar fix: '{}' → '{}' [{}]", ge.word, display_alt, ge.rule_name);
                        self.writing_errors.push(WritingError {
                            category: ErrorCategory::Grammar,
                            word: resp.sentence.to_string(),
                            suggestion: corrected,
                            explanation: format!("«{}» → «{}»: {}", ge.word, display_alt, ge.explanation),
                            rule_name: ge.rule_name.clone(),
                            sentence_context: resp.sentence.to_string(),
                            doc_offset: resp.doc_offset,
                            position: i,
                            ignored: false,
                            word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: resp.paragraph_id.clone(), error_word: ge.word.clone(),
                        });
                        // Blue underline for grammar errors
                        for b in &self.manager.bridges {
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
                        word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: resp.paragraph_id.clone(), error_word: first.word.clone(),
                    });
                    // Blue underline for grammar errors without suggestions too
                    for b in &self.manager.bridges {
                        b.underline_word(&first.word, &resp.paragraph_id, "#0000FF");
                    }
                }
            }

            // Handle unknown words (spelling errors) — from check_sentence_full
            for unk in resp.unknown_words.iter()
                .filter(|u| !self.user_dict.as_ref().map_or(false, |ud| ud.has_word(&u.word)))
                .filter(|u| !self.analyzer.as_ref().map_or(false, |a| a.has_word(&u.word)))
                .filter(|u| !self.wordfreq.as_ref().map_or(false, |wf| {
                    let freq = wf.get(&u.word.to_lowercase()).copied().unwrap_or(0);
                    freq >= 1000 // Only skip high-frequency words — low-freq entries may be junk
                }))
            {
                let mut best = unk.spelling_suggestions.first().cloned().unwrap_or_default()
                    .trim_matches(|c: char| c.is_whitespace() || c.is_control()).to_string();
                // Capitalize suggestion if word is at start of sentence or originally capitalized
                if !best.is_empty() {
                    let word_lower = unk.word.to_lowercase();
                    let at_sentence_start = resp.sentence.to_lowercase().starts_with(&word_lower);
                    let is_upper_in_original = resp.sentence.to_lowercase()
                        .find(&word_lower)
                        .and_then(|pos| resp.sentence[pos..].chars().next())
                        .map_or(false, |c| c.is_uppercase());
                    if at_sentence_start || is_upper_in_original {
                        let mut chars = best.chars();
                        best = chars.next().unwrap().to_uppercase().to_string() + chars.as_str();
                    }
                }
                if best.is_empty() && unk.split_suggestions.is_empty() {
                    // No suggestions at all — still flag as unknown
                    log!("  Unknown word: '{}' (no suggestions)", unk.word);
                } else {
                    log!("  Spelling: '{}' → '{}' (from grammar checker)", unk.word, best);
                }
                // Only add if not already in writing_errors for this paragraph
                let already_exists = self.writing_errors.iter().any(|e| {
                    e.word.to_lowercase() == unk.word.to_lowercase()
                    && e.paragraph_id == resp.paragraph_id
                    && !e.ignored
                });
                if !already_exists {
                    self.writing_errors.push(WritingError {
                        category: ErrorCategory::Spelling,
                        word: unk.word.clone(),
                        suggestion: best,
                        explanation: self.language.ui_word_not_in_dict(&unk.word),
                        rule_name: "stavefeil".to_string(),
                        sentence_context: resp.sentence.to_string(),
                        doc_offset: resp.doc_offset,
                        position: unk.position,
                        ignored: false,
                        word_doc_start: 0, word_doc_end: 0, underlined: false, pinned: false, paragraph_id: resp.paragraph_id.clone(), error_word: String::new(),
                    });
                    for b in &self.manager.bridges {
                        b.underline_word(&unk.word, &resp.paragraph_id, "#FF0000");
                    }
                }
            }

            // Stale underlines from previous sessions are cleared at app startup
            // via AppleScript (set underline of font to underline none).
        }

        // BERT re-rank spelling suggestions from grammar checker
        if self.bert_ready {
            let to_rerank: Vec<(usize, String, String)> = self.writing_errors.iter().enumerate()
                .filter(|(_, e)| {
                    matches!(e.category, ErrorCategory::Spelling)
                        && !e.ignored
                        && e.rule_name == "stavefeil"
                        && !e.sentence_context.is_empty()
                })
                .map(|(i, e)| (i, e.word.clone(), e.sentence_context.clone()))
                .collect();
            for (idx, word, sentence_ctx) in to_rerank {
                // Skip long words with close suggestions (compound words)
                let orig = &self.writing_errors[idx].suggestion;
                if word.len() >= 10 && !orig.is_empty() && orig.len() >= 8 {
                    let dist = levenshtein_distance(&word.to_lowercase(), &orig.to_lowercase());
                    if dist <= 2 {
                        self.writing_errors[idx].rule_name = "stavefeil_bert".to_string();
                        continue;
                    }
                }
                let suggestions = self.find_spelling_suggestions(&word, &sentence_ctx);
                // Pick best ortho+dict candidate (grammar filtering happens async in BERT worker)
                if let Some((best, score)) = suggestions.first().cloned() {
                    if !best.is_empty() {
                        let mut suggestion = best.trim_matches(|c: char| c.is_whitespace() || c.is_control()).to_string();
                        let word_lower = word.to_lowercase();
                        let at_start = sentence_ctx.to_lowercase().starts_with(&word_lower);
                        if at_start {
                            let mut chars = suggestion.chars();
                            suggestion = chars.next().unwrap().to_uppercase().to_string() + chars.as_str();
                        }
                        log!("Grammar spell rerank: '{}' → '{}' (score={:.2}, grammar-checked)", word, suggestion, score);
                        self.writing_errors[idx].suggestion = suggestion;
                    }
                }
                self.writing_errors[idx].rule_name = "stavefeil_bert".to_string();
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
                if !e.paragraph_id.is_empty() {
                    self.manager.clear_underline_word(&e.word, &e.paragraph_id);
                } else if e.word_doc_start < e.word_doc_end {
                    self.manager.clear_error_underline(e.word_doc_start, e.word_doc_end);
                }
                e.underlined = false;
            } else if !e.ignored && !e.underlined && !e.word.is_empty() {
                let mut marked = false;
                if !e.paragraph_id.is_empty() {
                    // Try Mac Add-in path first: underline using word + paragraph ID
                    let color = match e.category {
                        ErrorCategory::Spelling => "#FF0000",
                        ErrorCategory::Grammar => "#0000FF",
                        ErrorCategory::SentenceBoundary => "#0000FF",
                    };
                    marked = self.manager.underline_word(&e.word, &e.paragraph_id, color);
                    if marked {
                        log!("Underline: word='{}' para={} rule={} color={} ok={}",
                            e.word, trunc(&e.paragraph_id, 10), e.rule_name, color, marked);
                    }
                }
                // Fallback: Windows COM path using character range
                if !marked && e.word_doc_start < e.word_doc_end {
                    let ul_color = match e.category {
                        ErrorCategory::Spelling => bridge::ErrorUnderlineColor::Red,
                        _ => bridge::ErrorUnderlineColor::Blue,
                    };
                    marked = self.manager.mark_error_underline(e.word_doc_start, e.word_doc_end, ul_color);
                    log!("Underline: range {}..{} for '{}' rule={} color={:?} ok={}",
                        e.word_doc_start, e.word_doc_end, trunc(&e.word, 30), e.rule_name,
                        match ul_color { bridge::ErrorUnderlineColor::Red => "red", _ => "blue" }, marked);
                }
                e.underlined = marked;
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
                        self.last_dispatched_sentence.clear();
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
                    0 => (5, 1),
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
/// Find the first word that differs between two sentences.
fn find_diff_word(original: &str, corrected: &str) -> String {
    let orig_words: Vec<&str> = original.split_whitespace().collect();
    let corr_words: Vec<&str> = corrected.split_whitespace().collect();
    for (o, c) in orig_words.iter().zip(corr_words.iter()) {
        if o.to_lowercase() != c.to_lowercase() {
            return o.trim_matches(|c: char| c.is_ascii_punctuation()).to_string();
        }
    }
    // Length differs — last word of shorter
    if orig_words.len() != corr_words.len() {
        if let Some(w) = orig_words.last().or(corr_words.last()) {
            return w.trim_matches(|c: char| c.is_ascii_punctuation()).to_string();
        }
    }
    original.split_whitespace().next().unwrap_or(original).to_string()
}

fn try_split_function_word(
    word: &str,
    analyzer: &mtag::Analyzer,
    function_words: &[&str],
) -> Option<String> {
    // Function words come from the active language's `LanguageSpelling`
    // trait — see `BokmalLanguage::function_words` (and future
    // `NynorskLanguage`/`EnglishLanguage` impls). The caller passes
    // `self.language.function_words()`.

    let lower = word.to_lowercase();

    // Phase 1: Try known function word prefixes (high confidence)
    for prefix in function_words {
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
        // App-switch detection runs unconditionally at the very top of every frame,
        // before any processing or drawing, so stale errors are cleared immediately.
        {
            let fg = self.platform.foreground_app();
            let kind = self.platform.classify_app(&fg);
            let now_browser = kind == platform::AppKind::Browser;
            let now_word = kind == platform::AppKind::Word;

            if now_browser && !self.suppress_errors {
                log!("Browser foreground — clearing stale errors");
                self.clear_for_app_switch();
                ctx.request_repaint();
            }

            if now_word {
                let title = fg.title.clone();
                if !self.prev_word_title.is_empty() && title != self.prev_word_title {
                    log!("Word doc switch: '{}' → '{}' — clearing", self.prev_word_title, title);
                    self.clear_for_app_switch();
                    ctx.request_repaint();
                }
                self.prev_word_title = title;
            }

            self.suppress_errors = now_browser;
        }

        // In selection mode: handle keys at the top, set skip_processing flag
        let mut skip_processing = false;
        if self.selection_mode {
            skip_processing = true;
            let down = ctx.input(|i| i.key_pressed(egui::Key::ArrowDown) || i.key_pressed(egui::Key::Tab));
            let up = ctx.input(|i| i.key_pressed(egui::Key::ArrowUp));
            let left = ctx.input(|i| i.key_pressed(egui::Key::ArrowLeft));
            let right = ctx.input(|i| i.key_pressed(egui::Key::ArrowRight));
            let select = ctx.input(|i| i.key_pressed(egui::Key::Enter) || i.key_pressed(egui::Key::Space));
            let cancel = ctx.input(|i| i.key_pressed(egui::Key::Escape));

            if down {
                let max = if self.selected_column == 0 { self.completions.len() } else { self.open_completions.len() };
                if max > 0 { self.selected_completion = Some(self.selected_completion.map_or(0, |s| (s+1) % max)); }
            } else if up {
                self.selected_completion = Some(self.selected_completion.unwrap_or(0).saturating_sub(1));
            } else if left && !self.completions.is_empty() { self.selected_column = 0;
            } else if right && !self.open_completions.is_empty() { self.selected_column = 1;
            } else if select {
                if let Some(idx) = self.selected_completion {
                    let comp = if self.selected_column == 0 { self.completions.get(idx) } else { self.open_completions.get(idx) };
                    if let Some(c) = comp {
                        let prefix = self.context.word.clone();
                        let word = c.word.clone();
                        let col = self.selected_column;
                        log!("TAB SELECT: '{}' col={} for prefix '{}'", word, col, prefix);
                        // JS paragraph rewrite via bridge
                        self.manager.replace_word(&format!("{}|{}", prefix, word));
                        self.selection_mode = false;
                        self.platform.set_tab_intercept(false);
                        self.selected_completion = None;
                        self.completions.clear();
                        self.open_completions.clear();
                    }
                }
            } else if cancel {
                self.selection_mode = false;
                self.platform.set_tab_intercept(false);
                self.selected_completion = None;
            }
        }

        // Allow copying text from labels
        ctx.style_mut(|s| s.interaction.selectable_labels = true);

      if !skip_processing {
        // Spawn grammar actor on first update — loads SWI-Prolog on its own thread.
        if self.grammar_actor.is_none() && self.analyzer.is_some() {
            let compound_data_path = self.resolved_paths.prolog_dir.join("compound_data.pl");
            self.grammar_actor = Some(grammar_actor::spawn_grammar_actor_with_loader(
                self.platform.swipl_path().to_string(),
                self.resolved_paths.mtag_fst.to_str().unwrap().to_string(),
                self.resolved_paths.prolog_rules.to_str().unwrap().to_string(),
                self.resolved_paths.prolog_dir.to_str().unwrap().to_string(),
                std::fs::read_to_string(&compound_data_path).unwrap_or_default(),
                ctx.clone(),
            ));
            log!("Grammar actor spawning (SWI-Prolog loads on actor thread)");
        }

        // Spawn LLM actor on first update
        if self.llm_actor.is_none() {
            self.llm_actor = Some(llm_actor::spawn_llm_actor(ctx.clone()));
        }

        
        // suppress_errors and app-switch detection are handled at the very top
        // of update() (above the skip_processing gate) — nothing to do here.

        self.poll_grammar_responses();
       
        if let Some(actor) = &self.grammar_actor {
        }

        
        self.process_addin_changed_paragraphs();
       

        
        self.poll_llm_responses();
       

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
            if let Ok(mut shared) = self.shared_errors_json.lock() {
                *shared = json;
            }
        }

        // Send pending incomplete sentence to grammar actor after 1s idle (user stopped typing)
        if let Some((ref sentence, ref para_id, timestamp)) = self.pending_incomplete_sentence.clone() {
            if timestamp.elapsed() >= Duration::from_millis(1000) {
                let sent_h = hash_str(&format!("{}|{}", para_id, sentence));
                if !self.processed_sentence_hashes.contains(&sent_h) {
                    if let Some(actor) = &self.grammar_actor {
                        let uw = self.user_dict.as_ref().map_or(vec![], |ud| ud.list_words());
                        actor.check_sentence_with_doc(sentence, 0, para_id, 0, &self.last_doc_text, &uw);
                    }
                }
                self.pending_incomplete_sentence = None;
            }
        }

        // Execute deferred find-and-replace
        if let Some((find, replace, context, doc_offset)) = self.pending_fix.take() {
            log!("pending_fix: bridge='{}' find='{}' replace='{}' offset={}",
                self.manager.active_bridge_name(),
                trunc(&find, 60), trunc(&replace, 60), doc_offset);
            // Clear underline BEFORE replacement
            let find_lower_pre = find.to_lowercase();
            for e in &mut self.writing_errors {
                if (e.word.to_lowercase() == find_lower_pre || e.sentence_context.to_lowercase() == find_lower_pre)
                    && e.doc_offset == doc_offset
                {
                    // Clear both position-based (COM) and word-based (add-in) underlines
                    if e.underlined {
                        self.manager.clear_underline_word(&e.word, &e.paragraph_id);
                    }
                    self.manager.clear_underline_word(&e.word, &e.paragraph_id);
                    e.underlined = false;
                    log!("  Pre-cleared underline word='{}' para='{}'", e.word, trunc(&e.paragraph_id, 10));
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
                // processed_sentence_hashes NOT cleared — only invalidate changed sentence
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

        // Math OCR: check if background recognition finished
        if let Some(rx) = &self.math_receiver {
            if let Ok(result) = rx.try_recv() {
                match result {
                    Ok(text) => {
                        log!("Math OCR: '{}'", &text);
                        if !text.is_empty() {
                            tts::speak_word(&text);
                        }
                        self.ocr_text = Some(text);
                    }
                    Err(e) => {
                        log!("Math OCR error: {}", e);
                    }
                }
                self.math_receiver = None;
                // Dismiss the pending image
                if let Some(ocr) = &mut self.ocr {
                    ocr.dismiss();
                }
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
                        // Retroactively remove spelling errors for words now found in wordfreq
                        if let Some(wf) = &self.wordfreq {
                            let before = self.writing_errors.len();
                            self.writing_errors.retain(|e| {
                                !(matches!(e.category, ErrorCategory::Spelling) && wf.contains_key(&e.word.to_lowercase()))
                            });
                            let removed = before - self.writing_errors.len();
                            if removed > 0 {
                                log!("Wordfreq: removed {} false-positive spelling errors", removed);
                            }
                        }
                        self.embedding_store = embedding_store;
                        if let Some(m) = model {
                            let gs = self.grammar_actor.as_ref().map(|a| a.sender_clone());
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
                                gs,
                            ));
                            self.bert_ready = true;
                        }
                        self.load_errors.extend(errors);
                        self.startup_done.push("NorBERT4".into());
                        // Force rescan — spelling was skipped while BERT was loading
                        self.last_doc_hash = 0;
                        // processed_sentence_hashes NOT cleared — only invalidate changed sentence
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
                            self.whisper_load_status = self.language.ui_whisper_large_loaded().into();
                        }
                    }
                    WhisperLoadItem::Final(Err(e)) => {
                        log!("Whisper final model failed: {}", e);
                        self.load_errors.push(format!("Whisper: {}", e));
                    }
                    WhisperLoadItem::Streaming(Ok(engine)) => {
                        log!("Whisper: streaming model loaded");
                        self.whisper_streaming = Some(Arc::new(Mutex::new(engine)));
                        self.whisper_load_status = self.language.ui_whisper_fast_loaded().into();
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
                        let auto_final = cfg!(target_os = "macos");
                        match stt::start_recording(final_eng, stream_eng, auto_final, self.language.ui_no_audio_captured().to_string()) {
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
        // Poll for "Forbedre" result
        if let Some(ref rx) = self.improve_rx {
            if let Ok(text) = rx.try_recv() {
                log!("Improve result: '{}'", trunc(&text, 60));
                self.mic_result_text = Some(text);
                self.improve_rx = None;
                self.improve_running = false;
                ctx.request_repaint();
            }
        }
        // Keep repainting while waiting for whisper results
        if self.mic_handle.is_some() || self.mic_transcribing || self.improve_running {
            ctx.request_repaint_after(Duration::from_millis(100));
        }

        // Poll for new context
        if self.last_poll.elapsed() >= self.poll_interval {
            self.last_poll = Instant::now();

            // Track prev_fg_was_browser for legacy code paths.
            // The actual clear_for_app_switch() now runs in the pre-poll block
            // every frame so it's not gated by poll_interval.
            self.prev_fg_was_browser = self.suppress_errors;

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
            // Update caret position — platform-specific or from bridge context
            {
                let fg = self.platform.foreground_app();
                let kind = self.platform.classify_app(&fg);
                if kind == platform::AppKind::Word || kind == platform::AppKind::Browser {
                    // Try platform API first (macOS), fall back to bridge caret_pos (Windows)
                    if let Some((x, y)) = self.platform.caret_screen_position() {
                        self.last_caret_pos = Some((x, y + 49));
                    } else if let Some(ref ctx) = ctx_result {
                        if let Some((x, y)) = ctx.caret_pos {
                            if x != 0 || y != 0 {
                                self.last_caret_pos = Some((x, y));
                            }
                        }
                    }
                }
            }

            if let Some(new_ctx) = ctx_result {
                // Update caret position from bridge context (Windows fallback)
                if let Some((x, y)) = new_ctx.caret_pos {
                    if x != 0 || y != 0 {
                        self.last_caret_pos = Some((x, y));
                    }
                }
                let ctx_changed = new_ctx.word != self.context.word
                    || new_ctx.sentence != self.context.sentence
                    || new_ctx.masked_sentence != self.context.masked_sentence;
                // Track cursor offset for paragraph scanning when our window has focus
                if let Some(off) = new_ctx.cursor_doc_offset {
                    self.last_known_cursor_offset = Some(off);
                }
                if ctx_changed {
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
                    ; self.processed_sentence_hashes.clear();
                    self.last_doc_hash = 0;
                    // Do NOT reset last_sentence_count to 0 — that causes a
                    // false "major doc change" on the very next read, which
                    // clears the BERT queue before results arrive.
                }
                // Incremental paragraph scan: read only the paragraph at cursor (not full doc)
                let is_com_bridge = !self.manager.last_user_was_browser
                    && !self.manager.bridges.iter().any(|b| b.name() == "Word Add-in");
                if is_com_bridge {
                    if let Some(off) = new_ctx.cursor_doc_offset.or(self.last_known_cursor_offset) {
                        if let Some((para_id, text, start)) = self.manager.read_paragraph_at(off) {
                            self.process_com_changed_paragraph(para_id, text, start);
                        }
                    }
                } else if self.manager.last_user_was_browser {
                    if let Some(doc) = self.manager.read_full_document() {
                        self.try_update_doc_text(doc);
                    }
                }
                let fg = self.platform.foreground_app();
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
                        if big_change && !is_com_bridge {
                            // Paste/cut/move detected — trigger grammar rescan (non-COM only)
                            self.update_grammar_errors();
                            self.sync_error_underlines();
                        }
                    }
                    // Check if cursor is on an error word → activate in Grammatikk tab
                    let cursor_word = new_ctx.word.to_lowercase();
                    let cursor_para = new_ctx.paragraph_id.clone();
                    {
                        let hit = if !cursor_word.is_empty() {
                            self.writing_errors.iter().enumerate().find(|(_, e)| {
                                if e.ignored { return false; }
                                // Spelling errors: word field is the misspelled word
                                if matches!(e.category, ErrorCategory::Spelling) {
                                    return e.word.to_lowercase() == cursor_word;
                                }
                                // Grammar errors: match cursor word against the specific error_word
                                if matches!(e.category, ErrorCategory::Grammar) {
                                    if !cursor_para.is_empty() && e.paragraph_id != cursor_para {
                                        return false;
                                    }
                                    if !e.error_word.is_empty() {
                                        return e.error_word.to_lowercase() == cursor_word;
                                    }
                                }
                                false
                            })
                        } else {
                            None
                        };
                        if let Some((idx, e)) = hit {
                            if self.focused_error_idx != Some(idx) {
                                log!("Click hit: word='{}' → error idx={} '{}' rule={}",
                                    cursor_word, idx, trunc(&e.explanation, 40), e.rule_name);
                            }
                            // Switch to Grammatikk tab when clicking a grammar error
                            if matches!(e.category, ErrorCategory::Grammar) {
                                self.selected_tab = 1;
                            }
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

            // Scan for errors — COM bridge uses incremental paragraph scan (above),
            // other bridges use full doc scan
            let is_com = !self.manager.last_user_was_browser
                && !self.manager.bridges.iter().any(|b| b.name() == "Word Add-in");
            let errors_before = self.writing_errors.len();
            if !is_com {
                self.update_grammar_errors();
            }
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
                        // DISABLED: check_spelling runs find_spelling_suggestions on main thread = freeze
                        // Grammar actor already handles spelling via unknown words.
                        // log!("Word boundary spell check: '{}' in '{}' (cursor_off={})", w, trunc(&sentence, 50), cursor_off);
                        // {
                        //     let para_id = self.context.paragraph_id.clone();
                        //     self.check_spelling(w, &sentence, &para_id, 0);
                        // }
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
                // Non-COM bridges: trigger full doc scan
                if !is_com && self.grammar_queue.is_empty() {
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
            } else {
                // No word, no context: clear and run grammar
                self.completions.clear();
                self.open_completions.clear();
                self.last_completed_prefix.clear();
                self.last_dispatched_sentence.clear();
                // Check spelling + grammar on the last word/sentence
                let sentence = self.context.sentence.clone();
                let cursor_off = self.context.cursor_doc_offset.unwrap_or(0);
                let spell_word = sentence.split_whitespace().last()
                    .map(|w| w.trim_matches(|c: char| c.is_ascii_punctuation() || c == '«' || c == '»').to_string());
                if let Some(ref w) = spell_word {
                    if !w.is_empty() {
                        // DISABLED: check_spelling freezes main thread
                        // {
                        //     let para_id = self.context.paragraph_id.clone();
                        //     self.check_spelling(w, &sentence, &para_id, 0);
                        // }
                    }
                }
                self.validate_consonant_checks();
                self.run_grammar_check();
                if !is_com && self.grammar_queue.is_empty() {
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
            } // end if ctx_changed
            } // end if let Some(new_ctx)

            // Background completion: skip in selection mode
            if !self.selection_mode {
            if let Some(masked) = &self.context.masked_sentence.clone() {
                let prefix = extract_prefix(&self.context.word);
                let prefix_lower = prefix.to_lowercase();
                let cache_key = format!("{}|{}", masked, prefix);
                // Dedup: use sentence (masked text) as the stable key — prefix bounces don't re-trigger
                let sentence_key = masked.clone();
                let needs_completion = cache_key != self.last_completed_prefix;
                static COMP_LOG: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
                let tick = COMP_LOG.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
                if tick < 5 || (needs_completion && tick % 100 == 0) {
                    log!("COMPLETION prefix='{}' needs={} bert={} masked_len={} elapsed={}ms",
                        prefix, needs_completion, self.bert_worker.is_some(),
                        masked.len(), self.last_context_change.elapsed().as_millis());
                }

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

                // Dispatch after 150ms idle
                if needs_completion
                    && self.last_context_change.elapsed() >= Duration::from_millis(150)
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
                        let context_for_cw = if self.quality >= 2 {
                            // Høyest kvalitet: use full paragraph text for better context
                            let para_text = self.paragraph_texts.get(&self.context.paragraph_id)
                                .cloned()
                                .unwrap_or_else(|| self.context.sentence.clone());
                            para_text.strip_suffix(prefix).unwrap_or(&para_text).trim_end().to_string()
                        } else {
                            let sentence = &self.context.sentence;
                            sentence.strip_suffix(prefix).unwrap_or(sentence).trim_end().to_string()
                        };
                        let (top_n, max_steps) = match self.quality {
                            0 => (15, 1),
                            1 => (15, 1),
                            _ => (15, 3),
                        };
                        let ctx_tail: String = context_for_cw.chars().rev().take(30).collect::<Vec<_>>().into_iter().rev().collect();
                        log!("Sending CompleteWord: ctx='{}' prefix='{}' [queues: spell={} pend_bert={} gram_inflight={} gram_q={}]",
                            ctx_tail, prefix,
                            self.spelling_queue.len(),
                            self.pending_spelling_bert.len() + self.pending_grammar_bert.len() + self.pending_consonant_bert.len(),
                            self.grammar_inflight.len(),
                            self.grammar_queue.len());
                        self.last_completion_dispatch = Instant::now();
                        let cancel_clone = cancel.clone();
                        let sentence_clone = self.context.sentence.clone();
                        worker.send(|id| bert_worker::BertRequest::CompleteWord {
                            id,
                            context: context_for_cw,
                            prefix: prefix.to_string(),
                            capitalize,
                            top_n,
                            max_steps,
                            cache_key: key_clone,
                            masked_text: masked_trimmed,
                            cancel: cancel_clone,
                            sentence: sentence_clone,
                        });
                        // Mark as dispatched so we don't re-send
                        self.last_completed_prefix = cache_key.clone();
                        self.last_dispatched_sentence = sentence_key.clone();
                    }
                } else if needs_completion {
                }
            }
        }
        } // end if !selection_mode (completion dispatch)

        // Poll ALL BERT worker responses (completions + sentence scoring + MLM)
        if !self.selection_mode {
            self.poll_bert_responses(&ctx);
            self.validate_consonant_checks();
            if !self.spelling_queue.is_empty() {
                self.process_spelling_queue();
            }
            if !self.grammar_queue.is_empty() {
                self.process_grammar_queue();
            }
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

            if accept && !self.completions.is_empty() {
                // Handled by selection mode code at top of update()
                self.selection_mode = false;
                self.selected_completion = None;
            }
            if cancel {
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

      } // end if !skip_processing

        // Window sizing (scaled)
        let s = self.ui_scale;
        let has_content = !self.grammar_errors.is_empty() || !self.completions.is_empty() || !(&self.open_completions).is_empty();
        let recently_replaced = self.last_replace_time.elapsed() < Duration::from_secs(1);
        let win_h = s * if self.selected_tab >= 1 {
            250.0
        } else {
            150.0
        };
        let win_w = s * if self.selected_tab == 0 {
            242.0
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
            if self.show_settings_window {
                ctx.send_viewport_cmd(egui::ViewportCommand::WindowLevel(egui::WindowLevel::Normal));
            } else {
                ctx.send_viewport_cmd(egui::ViewportCommand::WindowLevel(egui::WindowLevel::AlwaysOnTop));
            }
            if let Some((x, y)) = self.last_caret_pos {
                let (screen_w, screen_h) = get_screen_size(&*self.platform);
                // DPI scaling: Windows returns physical pixels, macOS returns logical points.
                let (lx, ly) = if self.platform.caret_is_physical_pixels() {
                    let dpi_scale = ctx.pixels_per_point();
                    (x as f32 / dpi_scale, y as f32 / dpi_scale)
                } else {
                    (x as f32, y as f32)
                };
                // Push the window 5 cm below the caret so it doesn't cover the line
                // the user is currently writing on. 5 cm at 96 DPI = ~189 logical px.
                let caret_offset = self.platform.caret_offset_below();
                let pos_y = if (ly + caret_offset + win_h) > screen_h {
                    // Not enough room below — flip above the caret with a 30 px gap.
                    ly - win_h - 30.0
                } else {
                    ly + caret_offset
                };
                let pos_y = pos_y.max(0.0).min(screen_h - win_h);
                let pos_x = (lx + self.platform.caret_offset_right()).min(screen_w - win_w).max(0.0);

                ctx.send_viewport_cmd(egui::ViewportCommand::OuterPosition(
                    egui::pos2(pos_x, pos_y),
                ));
            }
        }

        // Repaint at 200ms interval — fast enough for responsive UI, avoids burning CPU.
        // NEVER use request_repaint() (no delay) — causes tight loop with COM calls.
        ctx.request_repaint_after(Duration::from_millis(200));

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
                        self.language.ui_ready().to_string()
                    } else {
                        self.language.ui_loading(&loading.join(", "))
                    };
                    ui.add(egui::ProgressBar::new(progress)
                        .text(label)
                        .desired_width(ui.available_width())
                        .desired_height(14.0));
                });
            });
            ctx.request_repaint_after(Duration::from_millis(100));
        }

        // Toolbar at bottom
        let tts_speaking = tts::is_speaking();
        let ocr_is_busy = self.ocr_receiver.is_some();
        egui::TopBottomPanel::bottom("toolbar").frame(panel_frame).show(ctx, |ui| {
            let header_resp = ui.horizontal(|ui| {
                let sep = egui::Color32::from_rgb(180, 170, 140);
                let active = egui::Color32::from_rgb(0, 70, 160);
                let inactive = egui::Color32::from_rgb(100, 100, 100);

                // --- Left side: 💡 ●✏ | 🎤 ▶ ---

                // 💡 Forslag (suggestions tab)
                let innhold_color = if self.selected_tab == 0 { active } else { inactive };
                if ui.add(egui::Label::new(
                    egui::RichText::new("\u{1F4A1}").size(16.0).color(innhold_color)
                ).sense(egui::Sense::click())).on_hover_text(self.language.ui_suggestions()).clicked() {
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
                ).sense(egui::Sense::click())).on_hover_text(self.language.ui_grammar()).clicked() {
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
                        )).on_hover_text(self.language.ui_transcribing());
                    } else {
                        if ui.add(egui::Button::new(
                            egui::RichText::new("■").size(12.0).color(egui::Color32::WHITE)
                        ).fill(egui::Color32::from_rgb(200, 40, 40))
                         .min_size(egui::vec2(22.0, 16.0))
                        ).on_hover_text(self.language.ui_stop_recording()).clicked() {
                            if let Some(handle) = &self.mic_handle {
                                handle.stop();
                                self.mic_transcribing = true;
                            }
                        }
                    }
                } else if self.whisper_loading {
                    ui.add(egui::Label::new(
                        egui::RichText::new("⏳").size(13.0)
                    )).on_hover_text(self.language.ui_loading_speech_model());
                    ctx.request_repaint_after(Duration::from_millis(100));
                } else {
                    let mic_color = inactive;
                    let whisper_ready = self.whisper_engine.is_some();
                    if ui.add(egui::Label::new(
                        egui::RichText::new("\u{1F3A4}").size(13.0).color(mic_color)
                    ).sense(egui::Sense::click())).on_hover_text(self.language.ui_speech_recognition()).clicked() {
                            if whisper_ready {
                                // Models already loaded — start recording immediately
                                let final_eng = self.whisper_engine.as_ref().unwrap().clone();
                                let stream_eng = self.whisper_streaming.as_ref().unwrap_or(&final_eng).clone();
                                let auto_final = cfg!(target_os = "macos");
                                match stt::start_recording(final_eng, stream_eng, auto_final, self.language.ui_no_audio_captured().to_string()) {
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
                                    let (fast_model, streaming_model, final_model) = self.language.whisper_model_names();
                                    self.whisper_load_status = if self.whisper_mode == 0 {
                                        self.language.ui_loading(self.language.whisper_fast_model_label())
                                    } else {
                                        self.language.ui_loading(self.language.whisper_best_model_label())
                                    };
                                    let (tx, rx) = std::sync::mpsc::channel();
                                    self.whisper_load_rx = Some(rx);
                                    let mode = self.whisper_mode;
                                    let dll_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                                        .join("../../whisper-build/bin/Release")
                                        .to_string_lossy().to_string();
                                    if mode == 0 {
                                        let dll = dll_dir.clone();
                                        let lang0 = self.language.clone();
                                        std::thread::spawn(move || {
                                            let model_path = data_dir()
                                                .join(fast_model)
                                                .to_string_lossy().to_string();
                                            let _ = tx.send(WhisperLoadItem::Final(
                                                stt::WhisperEngine::load(&dll, &model_path, &*lang0).map(|e| Box::new(e) as Box<dyn stt::SttEngine>)
                                            ));
                                        });
                                    } else {
                                        let tx2 = tx.clone();
                                        let dll2 = dll_dir.clone();
                                        let lang1 = self.language.clone();
                                        let lang2 = self.language.clone();
                                        std::thread::spawn(move || {
                                            let model_path = data_dir()
                                                .join(streaming_model)
                                                .to_string_lossy().to_string();
                                            let _ = tx2.send(WhisperLoadItem::Streaming(
                                                stt::WhisperEngine::load(&dll2, &model_path, &*lang1).map(|e| Box::new(e) as Box<dyn stt::SttEngine>)
                                            ));
                                        });
                                        std::thread::spawn(move || {
                                            let model_path = data_dir()
                                                .join(final_model)
                                                .to_string_lossy().to_string();
                                            let _ = tx.send(WhisperLoadItem::Final(
                                                stt::WhisperEngine::load(&dll_dir, &model_path, &*lang2).map(|e| Box::new(e) as Box<dyn stt::SttEngine>)
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
                    ).on_hover_text(self.language.ui_stop_reading()).clicked() {
                        tts::stop_speaking();
                        self.ocr_text = None;
                    }
                } else {
                    if ui.add(egui::Label::new(
                        egui::RichText::new("▶").size(14.0).color(inactive)
                    ).sense(egui::Sense::click())).on_hover_text(self.language.ui_read_selected_text()).clicked() {
                        log!("Speak button clicked!");
                        match self.manager.read_selected_text().or_else(|| self.platform.read_selected_text()) {
                            Some(text) => {
                                let trimmed = text.trim();
                                log!("Selected text: '{}'", trimmed);
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
                        ui.label(egui::RichText::new(self.language.ui_tip()).size(9.0).color(egui::Color32::from_rgb(120, 120, 120)));
                        ui.label(egui::RichText::new(format!("{}", err_count)).size(12.0).strong().color(egui::Color32::from_rgb(180, 60, 60)));
                    }
                }

                // --- Right side: drag area, 📌 ⚙ – ✕ ---
                let remaining = ui.available_rect_before_wrap();
                let right_w = 80.0; // 📌 + ⚙ + – + ✕
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
                // Phase 14: UI strings come from the runtime-selected
                // language stored on ContextApp.
                let pin_tooltip = if self.follow_cursor {
                    self.language.ui_pin_cursor_on()
                } else {
                    self.language.ui_pin_cursor_off()
                };
                if ui.add(egui::Label::new(
                    egui::RichText::new("\u{1F4CC}").size(14.0).color(pin_color)
                ).sense(egui::Sense::click())).on_hover_text(pin_tooltip).clicked() {
                    self.follow_cursor = !self.follow_cursor;
                }

                // ⚙ Settings
                let settings_color = if self.show_settings_window { active } else { inactive };
                if ui.add(egui::Label::new(
                    egui::RichText::new("\u{2699}").size(16.0).color(settings_color)
                ).sense(egui::Sense::click())).on_hover_text(self.language.ui_settings()).clicked() {
                    self.show_settings_window = !self.show_settings_window;
                }

                // ▁ Minimize
                if ui.add(egui::Label::new(
                    egui::RichText::new("–").size(14.0).color(inactive)
                ).sense(egui::Sense::click())).on_hover_text(self.language.ui_minimize()).clicked() {
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
            }).response;
            // Drag the window by dragging anywhere on the toolbar (when unpinned)
            if !self.follow_cursor {
                let header_drag = header_resp.interact(egui::Sense::drag());
                if header_drag.drag_started() {
                    ctx.send_viewport_cmd(egui::ViewportCommand::StartDrag);
                }
            }
        });

        egui::CentralPanel::default().frame(panel_frame).show(ctx, |ui| {
            // === Whisper transcription result — shown in separate centered window ===
            // (rendering happens below via show_viewport_immediate)

            // === Space press: speak the word just typed (accessibility TTS) ===
            if self.platform.take_space_press() && self.speak_on_space
                && self.last_space_speak.elapsed() > Duration::from_millis(400)
            {
                let fg = self.platform.foreground_app();
                let kind = self.platform.classify_app(&fg);
                if kind != platform::AppKind::OurApp {
                    if let Some(word) = self.platform.get_word_before_cursor() {
                        if !word.is_empty() {
                            self.last_space_speak = Instant::now();
                            tts::speak_word(&word);
                        }
                    }
                }
            }

            // === Tab: Innhold (0) ===
            if self.selected_tab == 0 {
                // Tab key selection
                let has_sugg = !self.completions.is_empty() || !self.open_completions.is_empty();
                if has_sugg && !self.selection_mode { self.platform.set_tab_intercept(true); }
                else if !has_sugg { self.platform.set_tab_intercept(false); self.selection_mode = false; }

                if self.platform.take_tab_press() && has_sugg {
                    self.selection_mode = true;
                    self.selected_completion = Some(0);
                    ctx.send_viewport_cmd(egui::ViewportCommand::Focus);
                    self.selected_column = if !self.completions.is_empty() { 0 } else { 1 };
                    ctx.send_viewport_cmd(egui::ViewportCommand::Focus);
                }

                // Selection mode keyboard handling is at the top of update()
                if !self.completions.is_empty() || !self.open_completions.is_empty() {

                    let sel = self.selected_completion;
                    let mut clicked_word: Option<(String, usize)> = None; // (word, column: 0=left, 1=right)
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
                        let col_w = (avail_w - 2.0) / 2.0;
                        let max_rows = self.completions.len().max(self.open_completions.len());
                        for row in 0..max_rows {
                            ui.horizontal(|ui| {
                                if row < self.completions.len() {
                                    let comp = &self.completions[row];
                                    let is_sel = self.selected_column == 0 && sel == Some(row);
                                    let is_top = row == 0 && sel.is_none();
                                    let (clicked, _) = render_row(ui, comp, row, is_sel, is_top, col_w);
                                    if clicked { clicked_word = Some((comp.word.clone(), 0)); }
                                } else {
                                    ui.allocate_exact_size(egui::vec2(col_w, 16.0), egui::Sense::hover());
                                }
                                ui.add_space(2.0);
                                if row < self.open_completions.len() {
                                    let comp = &self.open_completions[row];
                                    let is_sel = self.selected_column == 1 && sel == Some(row);
                                    let (clicked, _) = render_row(ui, comp, row + 100, is_sel, false, col_w);
                                    if clicked { clicked_word = Some((comp.word.clone(), 1)); }
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
                            if clicked { clicked_word = Some((comp.word.clone(), 1)); }
                        }
                    } else {
                        let avail_w = ui.available_width();
                        for (i, comp) in self.completions.iter().enumerate() {
                            let is_sel = sel == Some(i);
                            let is_top = i == 0 && sel.is_none();
                            let (clicked, _) = render_row(ui, comp, i, is_sel, is_top, avail_w);
                            if clicked { clicked_word = Some((comp.word.clone(), 0)); }
                        }
                    }


                    if let Some((word, col)) = clicked_word {
                        let prefix = self.context.word.clone();
                        log!("CLICKED word: '{}' col={} replacing '{}'", word, col, prefix);
                        // JS paragraph rewrite via bridge
                        self.manager.replace_word(&format!("{}|{}", prefix, word));
                        // Return focus to Word
                        let word_pid = self.manager.last_user_pid;
                        if word_pid > 0 {
                            std::thread::spawn(move || {
                                std::thread::sleep(std::time::Duration::from_millis(100));
                                let _ = std::process::Command::new("osascript").arg("-e")
                                    .arg(format!(r#"tell application "System Events"
                                        set frontProcess to first application process whose unix id is {}
                                        set frontmost of frontProcess to true
                                    end tell"#, word_pid)).output();
                            });
                        }
                        self.completions.clear();
                        self.open_completions.clear();
                        self.selected_completion = None;
                        self.selection_mode = false;
                    }
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
                        egui::RichText::new(self.language.ui_no_errors())
                            .size(12.0)
                            .color(egui::Color32::from_rgb(0, 140, 60)),
                    );
                } else {

                    // AI fix-all button + spinner
                    ui.horizontal(|ui| {
                        if self.llm_waiting {
                            let elapsed = self.llm_waiting_since.elapsed().as_secs();
                            let phases = ["🔄", "🔃"];
                            let phase = phases[(elapsed as usize) % phases.len()];
                            let pulse = if (self.llm_waiting_since.elapsed().as_millis() / 500) % 2 == 0 {
                                egui::Color32::from_rgb(0, 120, 220)
                            } else {
                                egui::Color32::from_rgb(60, 160, 240)
                            };
                            ui.label(egui::RichText::new(phase).size(16.0));
                            ui.label(egui::RichText::new(self.language.ui_ai_correcting_seconds(elapsed))
                                .size(12.0).strong().color(pulse));
                            ctx.request_repaint_after(Duration::from_millis(500));
                        } else {
                            let err_count = self.writing_errors.iter()
                                .filter(|e| !e.ignored && e.rule_name != "llm_correction")
                                .count();
                            if err_count > 0 {
                                if ui.add(egui::Button::new(
                                    egui::RichText::new(self.language.ui_ai_fix_all()).size(11.0)
                                ).min_size(egui::vec2(0.0, 18.0))).clicked() {
                                    self.dispatch_llm_fix_all();
                                }
                            }
                        }
                    });
                    ui.add_space(2.0);

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
                                    if icon_button(ui, "👍", self.language.ui_insert_period()) {
                                        action = Some((idx, "fix"));
                                    }
                                    if icon_button(ui, "👎", self.language.ui_ignore()) {
                                        action = Some((idx, "ignore"));
                                    }
                                    if icon_button(ui, "🔊", self.language.ui_read_aloud()) {
                                        tts::speak_word(&err_suggestion);
                                    }
                                    if icon_button(ui, "▶", self.language.ui_show_in_document()) {
                                        action = Some((idx, "goto"));
                                    }
                                });
                                ui.label(
                                    egui::RichText::new(self.language.ui_missing_period())
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
                                        if icon_button(ui, "👍", self.language.ui_fix()) {
                                            action = Some((alt_idx, "fix"));
                                        }
                                    }
                                    if icon_button(ui, "👎", self.language.ui_ignore()) {
                                        action = Some((idx, "ignore_group"));
                                    }
                                    if icon_button(ui, "🔊", self.language.ui_read_aloud()) {
                                        tts::speak_word(&first_suggestion);
                                    }
                                    if icon_button(ui, "💡", self.language.ui_show_rule_info()) {
                                        let fix_idx = first_alt.unwrap_or(idx);
                                        // Extract LLM changes if present
                                        self.rule_info_llm_changes = if err_expl.starts_with("LLM_CHANGES:") {
                                            let json_end = err_expl.find('\n').unwrap_or(err_expl.len());
                                            let json_str = &err_expl[12..json_end];
                                            serde_json::from_str::<Vec<serde_json::Value>>(json_str)
                                                .unwrap_or_default()
                                                .iter()
                                                .filter_map(|v| {
                                                    let from = v["from"].as_str()?.to_string();
                                                    let to = v["to"].as_str()?.to_string();
                                                    let why = v["why"].as_str().unwrap_or("").to_string();
                                                    Some((from, to, why))
                                                })
                                                .collect()
                                        } else {
                                            Vec::new()
                                        };
                                        self.rule_info_window = Some((err_rule.clone(), err_expl.clone(), err_ctx.clone(), fix_idx, first_suggestion.clone()));
                                    }
                                    if icon_button(ui, "▶", self.language.ui_show_in_document()) {
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
                                // Explanation (skip entirely for LLM corrections)
                                if error.rule_name != "llm_correction" {
                                    if let Some(alt_idx) = first_alt {
                                        ui.label(
                                            egui::RichText::new(&self.writing_errors[alt_idx].explanation)
                                                .size(10.0)
                                                .color(egui::Color32::from_rgb(100, 100, 100)),
                                        );
                                    }
                                }
                            } else {
                                // Spelling error — buttons on top, then word/suggestion stacked
                                let err_suggestion = error.suggestion.clone();
                                let err_word = error.word.clone();
                                ui.horizontal(|ui| {
                                    if !error.suggestion.is_empty() {
                                        if icon_button(ui, "👍", self.language.ui_fix()) {
                                            action = Some((idx, "fix"));
                                        }
                                    }
                                    if icon_button(ui, "👎", self.language.ui_ignore()) {
                                        action = Some((idx, "ignore"));
                                    }
                                    if icon_button(ui, "+", self.language.ui_add_to_dictionary()) {
                                        action = Some((idx, "add_to_dict"));
                                    }
                                    if icon_button(ui, "🔊", self.language.ui_read_aloud()) {
                                        let speak = if !err_suggestion.is_empty() { &err_suggestion } else { &err_word };
                                        tts::speak_word(speak);
                                    }
                                    if icon_button(ui, "?", self.language.ui_more_suggestions()) {
                                        action = Some((idx, "suggest"));
                                    }
                                    if icon_button(ui, "▶", self.language.ui_show_in_document()) {
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
                                // Skip raw explanation for LLM corrections (shown in 💡 diff view)
                                if error.rule_name != "llm_correction" {
                                    ui.label(
                                        egui::RichText::new(&error.explanation)
                                            .size(10.0)
                                            .color(egui::Color32::from_rgb(80, 80, 80)),
                                    );
                                }
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
                                // Clear underline immediately
                                self.manager.clear_underline_word(&error.word, &error.paragraph_id);
                                if error.underlined {
                                    self.manager.clear_underline_word(&error.word, &error.paragraph_id);
                                }
                                if matches!(error.category, ErrorCategory::Spelling) {
                                    self.ignored_words.insert(error.word.clone());
                                }
                                self.writing_errors[idx].ignored = true;
                                self.writing_errors[idx].underlined = false;
                            }
                            "add_to_dict" => {
                                let word = self.writing_errors[idx].word.clone();
                                log!("ACTION add_to_dict: '{}'", word);
                                if let Some(ud) = &self.user_dict {
                                    if let Err(e) = ud.add_word(&word) {
                                        eprintln!("Failed to add '{}' to user dict: {}", word, e);
                                    }
                                }
                                // Clear underlines for all instances of this word, then remove
                                let word_lower = word.to_lowercase();
                                for e in &self.writing_errors {
                                    if matches!(e.category, ErrorCategory::Spelling) && e.word.to_lowercase() == word_lower {
                                        self.manager.clear_underline_word(&e.word, &e.paragraph_id);
                                        if e.underlined {
                                            self.manager.clear_underline_word(&e.word, &e.paragraph_id);
                                        }
                                    }
                                }
                                self.writing_errors.retain(|e| {
                                    !(matches!(e.category, ErrorCategory::Spelling) && e.word.to_lowercase() == word_lower)
                                });
                                // Force rescan so the word is no longer flagged
                                // processed_sentence_hashes NOT cleared — only invalidate changed sentence
                                self.last_doc_hash = 0;
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
                                // Find the specific error word to select
                                let goto_word = if !error.error_word.is_empty() {
                                    error.error_word.clone()
                                } else {
                                    error.word.clone()
                                };
                                let para_id = error.paragraph_id.clone();
                                log!("GOTO: select '{}' in paragraph '{}'", goto_word, para_id);
                                // Try paragraph-based selection first (add-in), fall back to range
                                if !para_id.is_empty() && self.manager.select_word_in_paragraph(&goto_word, &para_id) {
                                    log!("GOTO: select_word_in_paragraph succeeded");
                                    // No screen position from add-in — skip window move
                                }
                                // Fall back to character range selection (Windows COM)
                                else {
                                let start = error.doc_offset;
                                let end = start + error.sentence_context.chars().count();
                                log!("GOTO: fallback selecting range {}..{}", start, end);
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
                                } // end else (fallback)
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
                let lang_for_vp = self.language.clone();

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
                            .with_title(lang_for_vp.ui_suggestions_for(&word_clone))
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
                                        egui::RichText::new(lang_for_vp.ui_suggestions_for(&word_clone))
                                            .size(16.0)
                                            .strong()
                                            .color(egui::Color32::from_rgb(30, 70, 150)),
                                    );
                                    ui.add_space(8.0);

                                    if candidates_clone.is_empty() {
                                        ui.label(lang_for_vp.ui_no_suggestions());
                                    } else {
                                        egui::ScrollArea::vertical().max_height(win_h - 80.0).show(ui, |ui| {
                                            for (i, (candidate, _score)) in candidates_clone.iter().enumerate() {
                                                ui.horizontal(|ui| {
                                                    if icon_button(ui, "🔊", lang_for_vp.ui_read_aloud()) {
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
                // Strip LLM_CHANGES: prefix from explanation for display
                let explanation = if explanation.starts_with("LLM_CHANGES:") {
                    explanation.find('\n').map(|p| &explanation[p+1..]).unwrap_or("").to_string()
                } else {
                    explanation.clone()
                };
                let sentence = sentence.clone();
                let fix_idx = *fix_idx;
                let suggestion = suggestion.clone();
                let error_word = self.writing_errors[fix_idx].word.clone();
                let corrected_sentence = if rule_name == "llm_correction" && !suggestion.is_empty() {
                    suggestion.clone() // LLM suggestion is already the full corrected sentence
                } else if !suggestion.is_empty() {
                    sentence.replacen(&error_word, &suggestion, 1)
                } else {
                    String::new()
                };
                let (category, description, wrong, right) = rule_info(&rule_name);
                let lang_for_rule = self.language.clone();

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
                        .with_title(lang_for_rule.ui_rule_info())
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

                                    // Word-level diff for LLM corrections, standard for others
                                    if rule_name == "llm_correction" && !corrected_sentence.is_empty() {
                                        egui::Frame::new()
                                            .fill(egui::Color32::from_rgb(250, 250, 255))
                                            .inner_margin(12.0)
                                            .corner_radius(6.0)
                                            .stroke(egui::Stroke::new(1.0, egui::Color32::from_rgb(200, 200, 220)))
                                            .show(ui, |ui| {
                                                ui.set_max_width(max_w - 40.0);
                                                // Word-level diff: red strikethrough for changed, green for corrections
                                                let orig_words: Vec<&str> = sentence.split_whitespace().collect();
                                                let corr_words: Vec<&str> = corrected_sentence.split_whitespace().collect();
                                                let job = egui::text::LayoutJob::default();
                                                ui.horizontal_wrapped(|ui| {
                                                    let mut oi = 0;
                                                    let mut ci = 0;
                                                    while oi < orig_words.len() || ci < corr_words.len() {
                                                        if oi < orig_words.len() && ci < corr_words.len()
                                                            && orig_words[oi].to_lowercase() == corr_words[ci].to_lowercase() {
                                                            // Same word — normal
                                                            ui.label(egui::RichText::new(orig_words[oi]).size(14.0));
                                                            oi += 1;
                                                            ci += 1;
                                                        } else {
                                                            // Find how many words differ
                                                            // Simple: consume original words until we find a match with corrected
                                                            let mut skip_orig = 0;
                                                            let mut skip_corr = 0;
                                                            // Try to re-sync by looking ahead
                                                            let mut found = false;
                                                            for ahead in 1..5 {
                                                                if ci + ahead < corr_words.len() && oi < orig_words.len()
                                                                    && orig_words[oi].to_lowercase() == corr_words[ci + ahead].to_lowercase() {
                                                                    skip_corr = ahead; found = true; break;
                                                                }
                                                                if oi + ahead < orig_words.len() && ci < corr_words.len()
                                                                    && orig_words[oi + ahead].to_lowercase() == corr_words[ci].to_lowercase() {
                                                                    skip_orig = ahead; found = true; break;
                                                                }
                                                            }
                                                            if !found { skip_orig = 1; skip_corr = 1; }
                                                            // Find matching explanation from LLM changes
                                                            let removed_text: String = (0..skip_orig)
                                                                .filter_map(|k| orig_words.get(oi + k))
                                                                .cloned().collect::<Vec<_>>().join(" ");
                                                            let tooltip = self.rule_info_llm_changes.iter()
                                                                .find(|(from, _, _)| removed_text.to_lowercase().contains(&from.to_lowercase()))
                                                                .map(|(_, _, why)| why.as_str())
                                                                .unwrap_or("");
                                                            // Show removed words in red strikethrough (with tooltip)
                                                            for _ in 0..skip_orig {
                                                                if oi < orig_words.len() {
                                                                    let r = ui.label(egui::RichText::new(orig_words[oi]).size(14.0)
                                                                        .strikethrough().color(egui::Color32::from_rgb(200, 50, 50)));
                                                                    if !tooltip.is_empty() { r.on_hover_text(tooltip); }
                                                                    oi += 1;
                                                                }
                                                            }
                                                            // Show added words in green bold (with tooltip)
                                                            for _ in 0..skip_corr {
                                                                if ci < corr_words.len() {
                                                                    let r = ui.label(egui::RichText::new(corr_words[ci]).size(14.0)
                                                                        .strong().color(egui::Color32::from_rgb(0, 140, 50)));
                                                                    if !tooltip.is_empty() { r.on_hover_text(tooltip); }
                                                                    ci += 1;
                                                                }
                                                            }
                                                        }
                                                    }
                                                });
                                            });
                                    } else {
                                        // Standard: original (red) and corrected (green)
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
                                    }
                                    ui.add_space(12.0);

                                    // Explanation (skip for LLM corrections — diff tooltips explain)
                                    if rule_name != "llm_correction" {
                                    ui.label(
                                        egui::RichText::new(lang_for_rule.ui_explanation())
                                            .size(14.0)
                                            .strong()
                                            .color(egui::Color32::from_rgb(50, 50, 50)),
                                    );
                                    ui.add_space(4.0);
                                    }
                                    if rule_name != "llm_correction" {
                                    ui.label(
                                        egui::RichText::new(&explanation)
                                            .size(14.0)
                                            .color(egui::Color32::from_rgb(30, 30, 30)),
                                    );
                                    }

                                    // Examples
                                    if !wrong.is_empty() {
                                        ui.add_space(14.0);
                                        ui.separator();
                                        ui.add_space(8.0);
                                        ui.label(
                                            egui::RichText::new(lang_for_rule.ui_examples())
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
                                        if ui.button(egui::RichText::new(lang_for_rule.ui_fix()).size(14.0).strong().color(egui::Color32::from_rgb(0, 120, 60))).clicked() {
                                            do_fix = true;
                                        }
                                        ui.add_space(8.0);
                                    }
                                    if ui.button(egui::RichText::new(lang_for_rule.ui_ignore()).size(14.0).color(egui::Color32::from_rgb(150, 60, 60))).clicked() {
                                        do_ignore = true;
                                    }
                                    ui.add_space(8.0);
                                    if ui.button(egui::RichText::new(lang_for_rule.ui_close()).size(14.0).color(egui::Color32::from_rgb(80, 80, 80))).clicked() {
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
                let mut do_improve = false;
                let text_clone = self.mic_result_text.clone().unwrap_or_default();
                let lang_for_stt = self.language.clone();

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
                        .with_title(lang_for_stt.ui_speech_recognition())
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
                                            lang_for_stt.ui_loading_speech_model().to_string()
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
                                            egui::RichText::new(lang_for_stt.ui_stop()).size(14.0).color(egui::Color32::WHITE)
                                        ).fill(egui::Color32::from_rgb(200, 40, 40))).clicked() {
                                            do_stop = true;
                                        }
                                    });
                                    ui.add_space(8.0);
                                } else if is_correcting {
                                    ui.horizontal(|ui| {
                                        ui.spinner();
                                        let msg = if self.whisper_mode == 1 && cfg!(target_os = "macos") {
                                            lang_for_stt.ui_improving_with_large_model()
                                        } else {
                                            lang_for_stt.ui_transcribing()
                                        };
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

                                // Show spinner when improving
                                if self.improve_running {
                                    ui.horizontal(|ui| {
                                        ui.spinner();
                                        ui.label(egui::RichText::new(lang_for_stt.ui_improving_with_large_model()).size(14.0)
                                            .color(egui::Color32::from_rgb(100, 80, 140)));
                                    });
                                    ui.add_space(4.0);
                                }

                                ui.add_space(8.0);
                                ui.horizontal(|ui| {
                                    if !is_streaming {
                                        if ui.button(egui::RichText::new(lang_for_stt.ui_copy()).size(14.0)).clicked() {
                                            do_copy = true;
                                        }
                                        ui.add_space(8.0);
                                        if ui.button(egui::RichText::new(format!("\u{1F50A} {}", lang_for_stt.ui_read_aloud())).size(14.0)).clicked() {
                                            tts::speak_word(&text_clone);
                                        }
                                        ui.add_space(8.0);
                                        if cfg!(target_os = "windows")
                                            && self.whisper_mode == 1
                                            && self.whisper_engine.is_some()
                                            && !self.improve_running
                                        {
                                            if ui.button(egui::RichText::new(lang_for_stt.ui_improve_result()).size(14.0)).clicked() {
                                                do_improve = true;
                                            }
                                            ui.add_space(8.0);
                                        }
                                    }
                                    if ui.button(egui::RichText::new(lang_for_stt.ui_close()).size(14.0)).clicked() {
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
                if do_improve {
                    if let Some(final_eng) = &self.whisper_engine {
                        if let Some(rx) = stt::improve_with_final_model(final_eng.clone()) {
                            self.improve_rx = Some(rx);
                            self.improve_running = true;
                            log!("Improve: started re-transcription with final model");
                        }
                    }
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

            // (Settings tab removed — settings now open in a separate window via the ⚙ icon)

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

        // Settings window (separate OS window, dyslexia-friendly large fonts)
        if self.show_settings_window {
            let mut do_close = false;
            let quality = self.quality;
            let grammar_completion = self.grammar_completion;
            let whisper_mode = self.whisper_mode;
            let speak_on_space = self.speak_on_space;
            let ui_scale = self.ui_scale;
            let dict_count = self.user_dict.as_ref().map_or(0, |ud| ud.list_words().len());
            let bridge_name = self.manager.active_bridge_name().to_string();
            let load_errors: Vec<String> = self.load_errors.clone();

            let mut new_quality = quality;
            let mut new_whisper_mode = whisper_mode;
            let mut new_speak_on_space = speak_on_space;
            let mut new_ui_scale = ui_scale;
            let mut open_voice = false;
            let mut open_userdict = false;
            let mut switch_to_language: Option<String> = None;
            let current_lang_code = self.language.code().to_string();
            let lang_for_settings = self.language.clone();

            ctx.show_viewport_immediate(
                egui::ViewportId::from_hash_of("settings_window"),
                egui::ViewportBuilder::default()
                    .with_title(lang_for_settings.ui_settings())
                    .with_inner_size([500.0, 600.0])
                    .with_decorations(true),
                |vp_ctx, _class| {
                    vp_ctx.set_visuals(egui::Visuals::light());

                    if vp_ctx.input(|i| i.viewport().close_requested()) {
                        do_close = true;
                    }

                    egui::CentralPanel::default()
                        .frame(egui::Frame::new().fill(egui::Color32::WHITE).inner_margin(24.0))
                        .show(vp_ctx, |ui| {
                            egui::ScrollArea::vertical().show(ui, |ui| {
                            let heading = 22.0_f32;
                            let body = 18.0_f32;
                            let label_color = egui::Color32::from_rgb(50, 50, 50);
                            let active_color = egui::Color32::from_rgb(0, 100, 180);
                            let on_color = egui::Color32::from_rgb(0, 130, 60);
                            let off_color = egui::Color32::from_rgb(140, 140, 140);

                            // -- Quality --
                            ui.label(egui::RichText::new(lang_for_settings.ui_quality()).size(heading).strong().color(label_color));
                            ui.add_space(6.0);
                            {
                                let quality_label = match new_quality {
                                    0 => lang_for_settings.ui_faster(),
                                    1 => lang_for_settings.ui_normal_quality(),
                                    _ => lang_for_settings.ui_highest_quality(),
                                };
                                let mut selected = new_quality as usize;
                                egui::ComboBox::from_id_salt("settings_quality_combo")
                                    .selected_text(egui::RichText::new(quality_label).size(body))
                                    .width(220.0)
                                    .show_index(ui, &mut selected, 3, |i| {
                                        match i {
                                            0 => lang_for_settings.ui_faster(),
                                            1 => lang_for_settings.ui_normal_quality(),
                                            _ => lang_for_settings.ui_highest_quality(),
                                        }.to_string()
                                    });
                                new_quality = selected as u8;
                            }
                            if grammar_completion {
                                ui.add_space(4.0);
                                ui.label(egui::RichText::new(lang_for_settings.ui_grammar_filter_on()).size(body).color(on_color));
                            }

                            ui.add_space(16.0);
                            ui.separator();
                            ui.add_space(12.0);

                            // -- Speech recognition --
                            ui.label(egui::RichText::new(lang_for_settings.ui_speech_recognition()).size(heading).strong().color(label_color));
                            ui.add_space(6.0);
                            ui.horizontal(|ui| {
                                let rask_color = if new_whisper_mode == 0 { active_color } else { off_color };
                                let beste_color = if new_whisper_mode == 1 { active_color } else { off_color };
                                if ui.add(egui::Label::new(
                                    egui::RichText::new(lang_for_settings.ui_stt_fast_model()).size(body).color(rask_color)
                                ).sense(egui::Sense::click())).clicked() {
                                    new_whisper_mode = 0;
                                }
                                ui.label(egui::RichText::new("  |  ").size(body).color(off_color));
                                if ui.add(egui::Label::new(
                                    egui::RichText::new(lang_for_settings.ui_stt_best_model()).size(body).color(beste_color)
                                ).sense(egui::Sense::click())).clicked() {
                                    new_whisper_mode = 1;
                                }
                            });

                            ui.add_space(16.0);
                            ui.separator();
                            ui.add_space(12.0);

                            // -- Voice --
                            ui.label(egui::RichText::new(lang_for_settings.ui_voice()).size(heading).strong().color(label_color));
                            ui.add_space(6.0);
                            ui.horizontal(|ui| {
                                let current = tts::current_voice();
                                ui.label(egui::RichText::new(&current).size(body).color(active_color));
                                if ui.add(egui::Button::new(
                                    egui::RichText::new(lang_for_settings.ui_choose()).size(body)
                                )).clicked() {
                                    open_voice = true;
                                }
                            });

                            ui.add_space(16.0);
                            ui.separator();
                            ui.add_space(12.0);

                            // -- Speak on space --
                            ui.label(egui::RichText::new(lang_for_settings.ui_read_words_aloud()).size(heading).strong().color(label_color));
                            ui.add_space(6.0);
                            {
                                let (label, color) = if new_speak_on_space {
                                    (lang_for_settings.ui_speak_on_space_on(), on_color)
                                } else {
                                    (lang_for_settings.ui_speak_on_space_off(), off_color)
                                };
                                if ui.add(egui::Label::new(
                                    egui::RichText::new(label).size(body).color(color)
                                ).sense(egui::Sense::click())).clicked() {
                                    new_speak_on_space = !new_speak_on_space;
                                }
                            }

                            ui.add_space(16.0);
                            ui.separator();
                            ui.add_space(12.0);

                            // -- UI scale --
                            ui.label(egui::RichText::new(lang_for_settings.ui_size()).size(heading).strong().color(label_color));
                            ui.add_space(6.0);
                            ui.horizontal(|ui| {
                                if ui.add(egui::Button::new(egui::RichText::new("  −  ").size(body))).clicked() {
                                    new_ui_scale = (new_ui_scale - 0.1).max(0.5);
                                }
                                ui.label(egui::RichText::new(format!("{:.0}%", new_ui_scale * 100.0)).size(body)
                                    .color(active_color));
                                if ui.add(egui::Button::new(egui::RichText::new("  +  ").size(body))).clicked() {
                                    new_ui_scale = (new_ui_scale + 0.1).min(2.5);
                                }
                            });

                            ui.add_space(16.0);
                            ui.separator();
                            ui.add_space(12.0);

                            // -- User dictionary --
                            ui.label(egui::RichText::new(lang_for_settings.ui_user_dict()).size(heading).strong().color(label_color));
                            ui.add_space(6.0);
                            ui.horizontal(|ui| {
                                ui.label(egui::RichText::new(format!("{} ord", dict_count)).size(body).color(active_color));
                                if ui.add(egui::Button::new(
                                    egui::RichText::new(lang_for_settings.ui_edit()).size(body)
                                )).clicked() {
                                    open_userdict = true;
                                }
                            });

                            ui.add_space(16.0);
                            ui.separator();
                            ui.add_space(12.0);

                            // -- Language --
                            ui.label(egui::RichText::new("Språk").size(heading).strong().color(label_color));
                            ui.add_space(6.0);

                            for lang in AVAILABLE_LANGUAGES {
                                let is_active = lang.code == current_lang_code;
                                let is_cached = downloader::language_cached(lang.code);

                                ui.horizontal(|ui| {
                                    // Flag + name
                                    paint_lang_flag(ui, lang.code, 18.0);
                                    ui.add_space(6.0);
                                    let color = if is_active {
                                        on_color
                                    } else if is_cached {
                                        active_color
                                    } else {
                                        off_color
                                    };
                                    ui.label(egui::RichText::new(lang.name).size(body).color(color));

                                    ui.add_space(8.0);

                                    if is_active {
                                        ui.label(egui::RichText::new("(aktiv)").size(14.0).color(on_color));
                                    } else if is_cached {
                                        // Already downloaded — offer to activate
                                        if ui.add(egui::Button::new(
                                            egui::RichText::new("Aktiver").size(15.0)
                                        )).clicked() {
                                            switch_to_language = Some(lang.code.to_string());
                                        }
                                    } else {
                                        // Not downloaded — offer to download + activate
                                        if ui.add(egui::Button::new(
                                            egui::RichText::new("Last ned").size(15.0)
                                        )).clicked() {
                                            switch_to_language = Some(lang.code.to_string());
                                        }
                                    }
                                });
                                ui.add_space(4.0);
                            }

                            // Load errors (if any)
                            if !load_errors.is_empty() {
                                ui.add_space(16.0);
                                ui.separator();
                                ui.add_space(12.0);
                                for err in &load_errors {
                                    ui.label(egui::RichText::new(err).size(16.0)
                                        .color(egui::Color32::from_rgb(200, 50, 50)));
                                }
                            }

                            ui.add_space(12.0);
                            ui.label(egui::RichText::new(format!("Bro: {}", bridge_name))
                                .size(14.0).color(off_color));
                            }); // end ScrollArea
                        });
                },
            );

            // Apply changes back
            if new_quality != self.quality {
                self.quality = new_quality;
                self.debounce_ms = if self.quality == 0 { 100 } else { 150 };
                log!("Quality changed to {}", self.quality);
            }
            if new_whisper_mode != self.whisper_mode {
                self.whisper_mode = new_whisper_mode;
                self.whisper_engine = None;
                self.whisper_streaming = None;
                log!("Whisper mode changed to {}", self.whisper_mode);
            }
            if new_speak_on_space != self.speak_on_space {
                self.speak_on_space = new_speak_on_space;
                log!("Speak on space: {}", self.speak_on_space);
            }
            if (new_ui_scale - self.ui_scale).abs() > 0.01 {
                self.ui_scale = new_ui_scale;
            }
            if open_voice {
                self.voice_list = tts::available_voices();
                self.show_voice_window = true;
            }
            if open_userdict {
                self.show_userdict_window = true;
            }
            // Save to disk whenever any setting changed
            if new_quality != quality || new_whisper_mode != whisper_mode
                || new_speak_on_space != speak_on_space
                || (new_ui_scale - ui_scale).abs() > 0.01
            {
                save_settings(&UserSettings {
                    quality: self.quality,
                    whisper_mode: self.whisper_mode,
                    speak_on_space: self.speak_on_space,
                    ui_scale: self.ui_scale,
                    voice: tts::current_voice(),
                    language: self.language.code().to_string(),
                });
            }
            if let Some(new_lang) = switch_to_language {
                // Save the new language and restart — download happens on startup
                save_settings(&UserSettings {
                    quality: self.quality,
                    whisper_mode: self.whisper_mode,
                    speak_on_space: self.speak_on_space,
                    ui_scale: self.ui_scale,
                    voice: tts::current_voice(),
                    language: new_lang.clone(),
                });
                // Restart the process with the new language
                let exe = std::env::current_exe().unwrap();
                let mut cmd = std::process::Command::new(exe);
                cmd.arg("--language").arg(&new_lang);
                let _ = cmd.spawn();
                std::process::exit(0);
            }
            if do_close {
                self.show_settings_window = false;
            }
        }

        // Voice selection window (separate from main panel)
        if self.show_voice_window {
            let mut open = self.show_voice_window;
            egui::Window::new(self.language.ui_choose_voice())
                .open(&mut open)
                .resizable(true)
                .default_width(300.0)
                .show(ctx, |ui| {
                    let current = tts::current_voice();
                    if self.voice_list.is_empty() {
                        ui.label(egui::RichText::new(self.language.ui_no_voices_found()).size(12.0));
                        ui.label(egui::RichText::new(self.language.ui_voice_download_help()).size(11.0)
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
                                save_settings(&UserSettings {
                                    quality: self.quality,
                                    whisper_mode: self.whisper_mode,
                                    speak_on_space: self.speak_on_space,
                                    ui_scale: self.ui_scale,
                                    voice: voice.name.clone(),
                                    language: self.language.code().to_string(),
                                });
                                tts::speak_word(&voice.sample_text);
                            }
                            ui.label(egui::RichText::new(&voice.language).size(10.0)
                                .color(egui::Color32::from_rgb(140, 140, 140)));
                        });
                    }
                });
            self.show_voice_window = open;
        }

        // User dictionary editor window (separate OS window)
        if self.show_userdict_window {
            let mut word_to_remove: Option<String> = None;
            let mut word_to_add: Option<String> = None;
            let mut do_close = false;
            let mut words: Vec<String> = self.user_dict.as_ref().map_or(vec![], |ud| ud.list_words());
            words.sort();
            let mut new_word_buf = self.userdict_new_word.clone();
            let scale = self.ui_scale;
            let lang_for_dict = self.language.clone();

            ctx.show_viewport_immediate(
                egui::ViewportId::from_hash_of("userdict_editor"),
                egui::ViewportBuilder::default()
                    .with_title(lang_for_dict.ui_user_dict())
                    .with_inner_size([350.0 * scale, 400.0 * scale])
                    .with_decorations(true),
                |vp_ctx, _class| {
                    vp_ctx.set_visuals(egui::Visuals::light());
                    vp_ctx.set_pixels_per_point(scale);

                    if vp_ctx.input(|i| i.viewport().close_requested()) {
                        do_close = true;
                    }

                    egui::CentralPanel::default()
                        .frame(egui::Frame::new().fill(egui::Color32::WHITE).inner_margin(12.0))
                        .show(vp_ctx, |ui| {
                            ui.label(egui::RichText::new(lang_for_dict.ui_words_you_added())
                                .size(14.0).color(egui::Color32::from_rgb(30, 30, 30)));
                            ui.add_space(4.0);

                            ui.horizontal(|ui| {
                                let response = ui.add(
                                    egui::TextEdit::singleline(&mut new_word_buf)
                                        .hint_text(lang_for_dict.ui_new_word_hint())
                                        .desired_width(220.0)
                                );
                                if ui.button(lang_for_dict.ui_add()).clicked()
                                    || (response.lost_focus() && ui.input(|i| i.key_pressed(egui::Key::Enter)))
                                {
                                    let trimmed = new_word_buf.trim().to_string();
                                    if !trimmed.is_empty() {
                                        word_to_add = Some(trimmed);
                                        new_word_buf.clear();
                                    }
                                }
                            });

                            ui.add_space(8.0);
                            ui.separator();
                            ui.add_space(4.0);

                            egui::ScrollArea::vertical().show(ui, |ui| {
                                if words.is_empty() {
                                    ui.label(egui::RichText::new(lang_for_dict.ui_no_words_added()).size(11.0)
                                        .color(egui::Color32::from_rgb(140, 140, 140)));
                                }
                                for w in &words {
                                    ui.horizontal(|ui| {
                                        ui.label(egui::RichText::new(w).size(13.0));
                                        if ui.small_button(lang_for_dict.ui_remove()).clicked() {
                                            word_to_remove = Some(w.clone());
                                        }
                                    });
                                }
                            });
                        });
                },
            );

            self.userdict_new_word = new_word_buf;

            if let Some(w) = word_to_add {
                if let Some(ud) = &self.user_dict {
                    if let Err(e) = ud.add_word(&w) {
                        eprintln!("Failed to add word: {}", e);
                    } else {
                        log!("User dict: added '{}'", w);
                    }
                }
                self.userdict_new_word.clear();
                let lower = w.to_lowercase();
                self.writing_errors.retain(|e| {
                    !(matches!(e.category, ErrorCategory::Spelling) && e.word.to_lowercase() == lower)
                });
                // processed_sentence_hashes NOT cleared — only invalidate changed sentence
                self.last_doc_hash = 0;
            }
            if let Some(w) = word_to_remove {
                if let Some(ud) = &self.user_dict {
                    if let Err(e) = ud.remove_word(&w) {
                        eprintln!("Failed to remove word: {}", e);
                    } else {
                        log!("User dict: removed '{}'", w);
                    }
                }
                // processed_sentence_hashes NOT cleared — only invalidate changed sentence
                self.last_doc_hash = 0;
            }
            if do_close {
                self.show_userdict_window = false;
            }
        }

        // OCR: screenshot detected prompt (separate OS window)
        let ocr_has_pending = self.ocr.as_ref().map_or(false, |o| o.has_pending_image());
        let ocr_is_busy = self.ocr_receiver.is_some() || self.math_receiver.is_some();
        if ocr_has_pending && !ocr_is_busy {
            let mut do_read = false;
            let mut do_copy = false;
            let mut do_math = false;
            let mut do_dismiss = false;

            let monitor = ctx.input(|i| i.viewport().monitor_size.unwrap_or(egui::vec2(1920.0, 1080.0)));
            let win_w: f32 = 320.0;
            let win_h: f32 = 100.0;
            let screen_center = egui::pos2(
                (monitor.x - win_w) / 2.0,
                (monitor.y - win_h) / 2.0,
            );

            let lang_for_ocr = self.language.clone();
            ctx.show_viewport_immediate(
                egui::ViewportId::from_hash_of("ocr_prompt"),
                egui::ViewportBuilder::default()
                    .with_title(lang_for_ocr.ui_screenshot_detected())
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
                                egui::RichText::new(lang_for_ocr.ui_text_found_in_screenshot())
                                    .size(14.0)
                                    .color(egui::Color32::from_rgb(30, 30, 30))
                            );
                            ui.add_space(8.0);
                            ui.horizontal(|ui| {
                                if ui.button(egui::RichText::new(lang_for_ocr.ui_read_text()).size(13.0)).clicked() {
                                    do_read = true;
                                }
                                if ui.button(egui::RichText::new(lang_for_ocr.ui_copy_text()).size(13.0)).clicked() {
                                    do_copy = true;
                                }
                                if ui.button(egui::RichText::new(lang_for_ocr.ui_math()).size(13.0)).clicked() {
                                    do_math = true;
                                }
                                if ui.button(egui::RichText::new(lang_for_ocr.ui_cancel()).size(13.0)).clicked() {
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
            if do_math {
                // Save clipboard image to temp file, then start math OCR
                if let Some(ocr) = &mut self.ocr {
                    let tmp_path = std::env::temp_dir().join("math_ocr_input.png");
                    if ocr.save_image_to(&tmp_path) {
                        let rx = math_ocr::start_math_ocr(tmp_path.to_string_lossy().to_string());
                        self.math_receiver = Some(rx);
                    } else {
                        log!("Math OCR: failed to save clipboard image");
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

// ── Language definitions for picker UI ──

struct LangOption {
    code: &'static str,
    name: &'static str,
}

const AVAILABLE_LANGUAGES: &[LangOption] = &[
    LangOption { code: "nb", name: "Bokmål" },
    LangOption { code: "nn", name: "Nynorsk" },
];

/// Paint a small Norwegian flag at the given position.
fn paint_norwegian_flag(painter: &egui::Painter, pos: egui::Pos2, size: f32) {
    let w = size * 1.5;
    let h = size;
    let rect = egui::Rect::from_min_size(pos, egui::vec2(w, h));

    // Red background
    let red = egui::Color32::from_rgb(186, 12, 47);
    painter.rect_filled(rect, 2.0, red);

    // White cross
    let white = egui::Color32::WHITE;
    let cx = pos.x + w * 0.36; // cross center x (off-center like real flag)
    let cross_w = h * 0.22;    // white cross width
    // Vertical bar
    painter.rect_filled(
        egui::Rect::from_min_max(
            egui::pos2(cx - cross_w / 2.0, pos.y),
            egui::pos2(cx + cross_w / 2.0, pos.y + h),
        ), 0.0, white,
    );
    // Horizontal bar
    painter.rect_filled(
        egui::Rect::from_min_max(
            egui::pos2(pos.x, pos.y + h / 2.0 - cross_w / 2.0),
            egui::pos2(pos.x + w, pos.y + h / 2.0 + cross_w / 2.0),
        ), 0.0, white,
    );

    // Blue cross (inside white)
    let blue = egui::Color32::from_rgb(0, 32, 91);
    let blue_w = h * 0.12;
    // Vertical bar
    painter.rect_filled(
        egui::Rect::from_min_max(
            egui::pos2(cx - blue_w / 2.0, pos.y),
            egui::pos2(cx + blue_w / 2.0, pos.y + h),
        ), 0.0, blue,
    );
    // Horizontal bar
    painter.rect_filled(
        egui::Rect::from_min_max(
            egui::pos2(pos.x, pos.y + h / 2.0 - blue_w / 2.0),
            egui::pos2(pos.x + w, pos.y + h / 2.0 + blue_w / 2.0),
        ), 0.0, blue,
    );

    // Border
    painter.rect_stroke(rect, 2.0, egui::Stroke::new(1.0, egui::Color32::from_rgb(180, 180, 180)), egui::StrokeKind::Outside);
}

/// Paint a flag for the given language code and return the space used.
fn paint_lang_flag(ui: &mut egui::Ui, lang_code: &str, size: f32) {
    let (rect, _response) = ui.allocate_exact_size(egui::vec2(size * 1.5, size), egui::Sense::hover());
    match lang_code {
        "nb" | "nn" => paint_norwegian_flag(ui.painter(), rect.min, size),
        _ => {} // future: other flags
    }
}

// ── Language picker: shown on first run ──
// Default UI language is Bokmål for the picker itself.

fn run_language_picker() -> Option<String> {
    let chosen: Arc<Mutex<Option<String>>> = Arc::new(Mutex::new(None));

    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([400.0, 340.0])
            .with_decorations(true)
            .with_title("NorskTale — Velg språk"),
        ..Default::default()
    };

    struct PickerApp {
        chosen: Arc<Mutex<Option<String>>>,
    }

    impl eframe::App for PickerApp {
        fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
            ctx.set_visuals(egui::Visuals::light());

            egui::CentralPanel::default()
                .frame(egui::Frame::new().fill(egui::Color32::WHITE).inner_margin(32.0))
                .show(ctx, |ui| {
                    ui.vertical_centered(|ui| {
                        ui.add_space(8.0);
                        ui.label(egui::RichText::new("Velg språk")
                            .size(26.0).strong().color(egui::Color32::from_rgb(40, 40, 40)));
                        ui.add_space(8.0);
                        ui.label(egui::RichText::new("Språkdata lastes ned etter valget.")
                            .size(16.0).color(egui::Color32::from_rgb(120, 120, 120)));
                        ui.add_space(24.0);

                        for lang in AVAILABLE_LANGUAGES {
                            let response = ui.horizontal(|ui| {
                                let btn_rect = ui.allocate_exact_size(
                                    egui::vec2(300.0, 48.0), egui::Sense::click()
                                );
                                let rect = btn_rect.0;
                                let response = btn_rect.1;

                                // Button background
                                let bg = if response.hovered() {
                                    egui::Color32::from_rgb(230, 240, 250)
                                } else {
                                    egui::Color32::from_rgb(245, 245, 245)
                                };
                                ui.painter().rect_filled(rect, 8.0, bg);
                                ui.painter().rect_stroke(rect, 8.0,
                                    egui::Stroke::new(1.0, egui::Color32::from_rgb(200, 200, 200)),
                                    egui::StrokeKind::Outside);

                                // Flag
                                let flag_y = rect.min.y + (rect.height() - 22.0) / 2.0;
                                paint_norwegian_flag(ui.painter(),
                                    egui::pos2(rect.min.x + 16.0, flag_y), 22.0);

                                // Text
                                ui.painter().text(
                                    egui::pos2(rect.min.x + 56.0, rect.center().y),
                                    egui::Align2::LEFT_CENTER,
                                    lang.name,
                                    egui::FontId::proportional(22.0),
                                    egui::Color32::from_rgb(40, 40, 40),
                                );

                                response
                            });

                            if response.inner.clicked() {
                                if let Ok(mut c) = self.chosen.lock() {
                                    *c = Some(lang.code.to_string());
                                }
                                ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                            }
                            ui.add_space(10.0);
                        }
                    });
                });
        }
    }

    let chosen_clone = Arc::clone(&chosen);
    let _ = eframe::run_native(
        "NorskTale — Velg språk",
        options,
        Box::new(move |_cc| {
            Ok(Box::new(PickerApp { chosen: chosen_clone }) as Box<dyn eframe::App>)
        }),
    );

    chosen.lock().ok().and_then(|c| c.clone())
}

// ── Download window: shown when language data is missing ──

fn run_download_window(lang_code: &str) {
    let items = downloader::language_files(lang_code);
    let progress = downloader::download_missing(items);

    // If nothing to download, return immediately
    if downloader::all_done(&progress) {
        return;
    }

    let lang_info = AVAILABLE_LANGUAGES.iter().find(|l| l.code == lang_code);
    let lang_name = lang_info.map(|l| l.name).unwrap_or(lang_code);
    let dl_lang_code = lang_code.to_string();

    let (win_title, heading_text) = if lang_code == "nn" {
        (
            format!("NorskTale — Lastar ned {}", lang_name),
            format!("Lastar ned {}...", lang_name),
        )
    } else {
        (
            format!("NorskTale — Laster ned {}", lang_name),
            format!("Laster ned {}...", lang_name),
        )
    };

    let prog = std::sync::Arc::clone(&progress);
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([480.0, 340.0])
            .with_decorations(true)
            .with_title(&win_title),
        ..Default::default()
    };

    struct DownloadApp {
        progress: downloader::SharedProgress,
        done: bool,
        heading: String,
        lang_code: String,
        error_text: &'static str,
    }

    impl eframe::App for DownloadApp {
        fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
            ctx.set_visuals(egui::Visuals::light());

            egui::CentralPanel::default()
                .frame(egui::Frame::new().fill(egui::Color32::WHITE).inner_margin(24.0))
                .show(ctx, |ui| {
                    ui.horizontal(|ui| {
                        paint_lang_flag(ui, &self.lang_code, 24.0);
                        ui.add_space(8.0);
                        ui.label(egui::RichText::new(&self.heading)
                            .size(22.0).strong().color(egui::Color32::from_rgb(50, 50, 50)));
                    });
                    ui.add_space(16.0);

                    if let Ok(items) = self.progress.lock() {
                        for item in items.iter() {
                            ui.horizontal(|ui| {
                                let pct = if item.total > 0 {
                                    item.downloaded as f32 / item.total as f32
                                } else {
                                    0.0
                                };
                                let status = if item.done {
                                    if item.error.is_some() { "Feil" } else { "Ferdig" }
                                } else if item.total > 0 {
                                    ""
                                } else {
                                    "Ventar..."
                                };

                                let color = if item.error.is_some() {
                                    egui::Color32::from_rgb(200, 50, 50)
                                } else if item.done {
                                    egui::Color32::from_rgb(0, 130, 60)
                                } else {
                                    egui::Color32::from_rgb(50, 50, 50)
                                };

                                ui.label(egui::RichText::new(&item.label).size(16.0).color(color));
                                ui.add_space(8.0);

                                if !status.is_empty() {
                                    ui.label(egui::RichText::new(status).size(14.0).color(color));
                                } else {
                                    let bar = egui::ProgressBar::new(pct)
                                        .desired_width(180.0)
                                        .text(format!("{:.0}%", pct * 100.0));
                                    ui.add(bar);
                                    let mb = item.downloaded as f64 / (1024.0 * 1024.0);
                                    let total_mb = item.total as f64 / (1024.0 * 1024.0);
                                    ui.label(egui::RichText::new(
                                        format!("{:.1}/{:.1} MB", mb, total_mb)
                                    ).size(13.0).color(egui::Color32::from_rgb(120, 120, 120)));
                                }
                            });
                            ui.add_space(4.0);
                        }
                    }

                    if downloader::all_done(&self.progress) {
                        self.done = true;
                        ui.add_space(12.0);
                        if downloader::any_error(&self.progress).is_some() {
                            ui.label(egui::RichText::new(self.error_text)
                                .size(16.0).color(egui::Color32::from_rgb(200, 50, 50)));
                        } else {
                            // Auto-close after download completes
                            ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                        }
                    }
                });

            // Repaint frequently to update progress
            ctx.request_repaint_after(Duration::from_millis(100));
        }
    }

    let error_text_static: &'static str = if lang_code == "nn" {
        "Nedlasting feila. Start programmet på nytt."
    } else {
        "Nedlasting feilet. Start programmet på nytt."
    };
    let _ = eframe::run_native(
        &win_title,
        options,
        Box::new(move |_cc| {
            Ok(Box::new(DownloadApp {
                progress: prog,
                done: false,
                heading: heading_text,
                lang_code: dl_lang_code,
                error_text: error_text_static,
            }) as Box<dyn eframe::App>)
        }),
    );
}

fn main() -> eframe::Result {
    let setup_platform = platform::create_platform();

    // Clear stale underlines from previous session (saved in document formatting)
    #[cfg(target_os = "macos")]
    {
        let _ = std::process::Command::new("osascript")
            .arg("-e")
            .arg("tell application \"Microsoft Word\"\ntry\nset f to font object of text object of active document\nset underline of f to underline none\nend try\nend tell")
            .output();
    }

    // --clear-cache: remove all downloaded data and settings (for testing)
    if std::env::args().any(|a| a == "--clear-cache") {
        let data = downloader::data_dir();
        eprintln!("Clearing cache: {}", data.display());
        let _ = std::fs::remove_dir_all(&data);
        let settings = settings_path();
        eprintln!("Clearing settings: {}", settings.display());
        let _ = std::fs::remove_file(&settings);
        eprintln!("Cache cleared.");
    }

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
        let test_language: std::sync::Arc<dyn language::LanguageBundle> =
            std::sync::Arc::new(language::BokmalLanguage);
        let test_paths = resolve_paths(&*test_language);
        let mut app = ContextApp::new(test_language, true, 2, false, UserSettings::default(), test_paths);
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
                                None,
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

    let mut saved = load_settings();
    let grammar_completion = !std::env::args().any(|a| a == "--no-grammar");
    let show_debug_tab = std::env::args().any(|a| a == "--debug");
    let quality: u8 = {
        let args: Vec<String> = std::env::args().collect();
        args.iter()
            .position(|a| a == "--quality")
            .and_then(|i| args.get(i + 1))
            .and_then(|v| v.parse().ok())
            .unwrap_or(saved.quality)
    };

    // Language: CLI flag overrides saved setting
    let lang_code: String = {
        let args: Vec<String> = std::env::args().collect();
        args.iter()
            .position(|a| a == "--language")
            .and_then(|i| args.get(i + 1).cloned())
            .unwrap_or_else(|| {
                if saved.language.is_empty() { "nb".to_string() }
                else { saved.language.clone() }
            })
    };

    // ── First-run: language picker + download ──
    let lang_code = if !downloader::language_cached(&lang_code) {
        // No cached data — show language picker first (unless CLI forced a language)
        let cli_forced = std::env::args().any(|a| a == "--language");
        let picked = if cli_forced {
            lang_code.clone()
        } else {
            match run_language_picker() {
                Some(code) => code,
                None => {
                    eprintln!("No language selected — exiting.");
                    std::process::exit(0);
                }
            }
        };
        eprintln!("Downloading language data for '{}'...", picked);
        run_download_window(&picked);
        picked
    } else {
        lang_code
    };

    // Persist the chosen language
    if saved.language != lang_code {
        saved.language = lang_code.clone();
        save_settings(&saved);
    }
    let selected_language: std::sync::Arc<dyn language::LanguageBundle> =
        match language::resolve_language(&lang_code) {
            Ok(l) => l,
            Err(e) => {
                eprintln!("Language: {}", e);
                std::process::exit(2);
            }
        };

    if grammar_completion {
        eprintln!("Grammar completion: ON");
    }
    eprintln!("SWI-Prolog engine: ON");
    eprintln!(
        "Language: {} ({})",
        selected_language.display_name(),
        selected_language.code()
    );
    let quality_name = match quality { 0 => "Raskere", 1 => "Normal", _ => "Høyeste kvalitet" };
    eprintln!("Quality: {} ({})", quality, quality_name);

    // Initialize TTS engine (platform-specific)
    setup_platform.init_tts(&*selected_language);
    if !saved.voice.is_empty() {
        tts::set_voice(&saved.voice);
    }

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

    let language_for_app = std::sync::Arc::clone(&selected_language);
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
                let paths = resolve_paths(&*language_for_app);
                let app = ContextApp::new(language_for_app.clone(), grammar_completion, quality, show_debug_tab, saved, paths);
                // Start HTTP /errors endpoint for integration tests
                let errors_json = app.shared_errors_json.clone();
                std::thread::Builder::new().name("test-http".into()).spawn(move || {
                    if let Ok(listener) = std::net::TcpListener::bind("127.0.0.1:52580") {
                        log!("Test HTTP server listening on http://127.0.0.1:52580/errors");
                        for stream in listener.incoming().flatten() {
                            let mut buf = [0u8; 1024];
                            let _ = std::io::Read::read(&mut &stream, &mut buf);
                            let json = errors_json.lock().map(|j| j.clone()).unwrap_or_else(|_| "[]".into());
                            let resp = format!(
                                "HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nAccess-Control-Allow-Origin: *\r\nContent-Length: {}\r\n\r\n{}",
                                json.len(), json
                            );
                            let _ = std::io::Write::write_all(&mut &stream, resp.as_bytes());
                        }
                    }
                }).ok();
                Ok(Box::new(app))
            }
        }),
    )
}
