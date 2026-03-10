mod bridge;
mod grammar_actor;
mod microphone;
mod ocr;
mod tts;

use bridge::{CursorContext, TextBridge};
use std::collections::HashMap;
use std::hash::{Hash, Hasher};
use std::io::Write;
use std::path::PathBuf;
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

static LOG_FILE: std::sync::LazyLock<Mutex<std::fs::File>> = std::sync::LazyLock::new(|| {
    let path = std::env::temp_dir().join("acatts-rust.log");
    eprintln!("Logging to: {}", path.display());
    let f = std::fs::OpenOptions::new()
        .create(true).write(true).truncate(true)
        .open(&path).expect("failed to open log file");
    Mutex::new(f)
});

macro_rules! log {
    ($($arg:tt)*) => {{
        let msg = format!($($arg)*);
        eprintln!("{}", msg);
        if let Ok(mut f) = LOG_FILE.lock() {
            let _ = writeln!(f, "{}", msg);
            let _ = f.flush();
        }
    }};
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
}

// --- Data paths ---

fn data_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../contexter-repo/training-data")
}

fn dict_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../rustSpell/mtag-rs/data/fullform_bm.mfst")
}

fn compound_data_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../syntaxer/compound_data.pl")
}

fn grammar_rules_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../syntaxer/grammar_rules.pl")
}

fn syntaxer_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../syntaxer")
}

fn swipl_dll_path() -> &'static str {
    "C:/Program Files/swipl/bin/libswipl.dll"
}

// --- Bridge manager: picks the best available bridge ---

struct BridgeManager {
    bridges: Vec<Box<dyn TextBridge>>,
    last_check: Instant,
    /// Index of the bridge that last successfully read context
    active_idx: usize,
}

impl BridgeManager {
    fn new() -> Self {
        let mut bridges: Vec<Box<dyn TextBridge>> = Vec::new();

        #[cfg(target_os = "windows")]
        {
            if let Some(word) = bridge::word_com::WordComBridge::try_connect() {
                println!("Word COM bridge connected");
                bridges.push(Box::new(word));
            }
            bridges.push(Box::new(bridge::accessibility_win::AccessibilityBridge::new()));
        }

        BridgeManager {
            bridges,
            last_check: Instant::now(),
            active_idx: 0,
        }
    }

    fn read_context(&mut self) -> Option<CursorContext> {
        #[cfg(target_os = "windows")]
        if self.last_check.elapsed() > Duration::from_secs(5) {
            self.last_check = Instant::now();
            let has_word = self.bridges.iter().any(|b| b.name() == "Word COM");
            if !has_word {
                if let Some(word) = bridge::word_com::WordComBridge::try_connect() {
                    println!("Word COM bridge connected (late)");
                    self.bridges.insert(0, Box::new(word));
                }
            }
        }

        // Only update active_idx when a real app has focus (not our own windows)
        let our_window_focused = {
            #[cfg(target_os = "windows")]
            {
                use windows::Win32::UI::WindowsAndMessaging::{GetForegroundWindow, GetWindowTextW};
                let fg = unsafe { GetForegroundWindow() };
                let mut buf = [0u16; 64];
                let len = unsafe { GetWindowTextW(fg, &mut buf) };
                let title = String::from_utf16_lossy(&buf[..len as usize]);
                title.contains("NorskTale") || title.starts_with("Forslag") || title.starts_with("Regelinfo")
            }
            #[cfg(not(target_os = "windows"))]
            { false }
        };

        for (i, bridge) in self.bridges.iter().enumerate() {
            if bridge.is_available() {
                if let Some(ctx) = bridge.read_context() {
                    if !our_window_focused {
                        if self.active_idx != i {
                            eprintln!("Bridge switch: {} → {} ('{}')",
                                self.bridges[self.active_idx].name(), bridge.name(),
                                if !ctx.word.is_empty() { &ctx.word } else { &ctx.sentence });
                        }
                        self.active_idx = i;
                    }
                    return Some(ctx);
                }
            }
        }
        None
    }

    fn active_bridge(&self) -> Option<&dyn TextBridge> {
        self.bridges.get(self.active_idx).map(|b| b.as_ref())
    }

    fn active_bridge_name(&self) -> &str {
        self.active_bridge().map(|b| b.name()).unwrap_or("none")
    }

    #[allow(dead_code)]
    fn replace_word(&self, new_text: &str) -> bool {
        self.active_bridge().map(|b| b.replace_word(new_text)).unwrap_or(false)
    }

    fn find_and_replace(&self, find: &str, replace: &str) -> bool {
        self.active_bridge().map(|b| b.find_and_replace(find, replace)).unwrap_or(false)
    }

    fn find_and_replace_in_context(&self, find: &str, replace: &str, context: &str) -> bool {
        self.active_bridge().map(|b| b.find_and_replace_in_context(find, replace, context)).unwrap_or(false)
    }

    fn find_and_replace_in_context_at(&self, find: &str, replace: &str, context: &str, char_offset: usize) -> bool {
        self.active_bridge().map(|b| b.find_and_replace_in_context_at(find, replace, context, char_offset)).unwrap_or(false)
    }

    fn read_document_context(&self) -> Option<String> {
        self.active_bridge().and_then(|b| b.read_document_context())
    }

    fn read_full_document(&self) -> Option<String> {
        self.active_bridge().and_then(|b| b.read_full_document())
    }

    fn select_range(&self, char_start: usize, char_end: usize) -> bool {
        self.active_bridge().map(|b| b.select_range(char_start, char_end)).unwrap_or(false)
    }

    fn set_target_hwnd(&self, hwnd: isize) {
        for bridge in &self.bridges {
            bridge.set_target_hwnd(hwnd);
        }
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

// --- egui app ---

/// Items delivered by background startup threads
enum StartupItem {
    Completer {
        model: Option<Arc<Mutex<Model>>>,
        prefix_index: Option<PrefixIndex>,
        baselines: Option<Baselines>,
        wordfreq: Option<Arc<HashMap<String, u64>>>,
        embedding_store: Option<EmbeddingStore>,
        errors: Vec<String>,
    },
    Whisper(Result<microphone::WhisperEngine, String>),
}

struct ContextApp {
    manager: BridgeManager,
    context: CursorContext,
    last_poll: Instant,
    poll_interval: Duration,
    follow_cursor: bool,
    last_caret_pos: Option<(i32, i32)>,
    // Grammar checker (kept for main-thread dictionary lookups; SWI grammar ops go through actor)
    checker: Option<AnyChecker>,
    /// Direct analyzer reference for dictionary lookups (cloned from checker before actor takes it)
    analyzer: Option<std::sync::Arc<mtag::Analyzer>>,
    /// Grammar actor: runs grammar checking on background thread
    grammar_actor: Option<grammar_actor::GrammarActorHandle>,
    grammar_errors: Vec<GrammarError>,
    last_checked_sentence: String,
    // Word completer
    model: Option<Arc<Mutex<Model>>>,
    /// Background completion: receives (cache_key, completions, open_completions)
    completion_rx: Option<std::sync::mpsc::Receiver<(String, Vec<Completion>, Vec<Completion>)>>,
    completion_cancel: Arc<std::sync::atomic::AtomicBool>,
    /// Last time context changed — for debouncing completion dispatch
    last_context_change: Instant,
    /// The cache key we last dispatched (avoid re-dispatching same)
    dispatched_key: String,
    prefix_index: Option<PrefixIndex>,
    baselines: Option<Baselines>,
    wordfreq: Option<Arc<HashMap<String, u64>>>,
    embedding_store: Option<EmbeddingStore>,
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
    word_hwnd: Option<isize>,
    /// Track Ctrl+Space held to prevent repeated activation
    ctrl_space_held: bool,
    /// Which column is selected: 0=left (completions), 1=right (open_completions)
    selected_column: u8,
    // Status
    load_errors: Vec<String>,
    // Tab navigation
    selected_tab: usize, // 0=Innhold, 1=Grammatikk, 2=Innstillinger, 3=Debug
    // Error list (spelling + grammar)
    writing_errors: Vec<WritingError>,
    /// Words the user has chosen to ignore (spelling)
    ignored_words: std::collections::HashSet<String>,
    /// Last word that was spell-checked (to avoid re-checking)
    last_spell_checked_word: String,
    /// Previous document text — used to detect changes and skip re-checking unchanged sentences
    last_doc_text: String,
    /// Hash of last document text — skip entire update if doc unchanged
    last_doc_hash: u64,
    /// Number of sentences in last scan — detect paste vs fix
    last_sentence_count: usize,
    /// Hashes of sentences already checked for Prolog sub-splitting (expensive, persists across doc changes)
    prolog_checked_hashes: std::collections::HashSet<u64>,
    /// Hashes of sentences grammar-checked and found clean (no errors)
    clean_sentence_hashes: std::collections::HashSet<u64>,
    /// Pending grammar work: sentences still to check (incremental, one per frame)
    grammar_queue: Vec<(String, usize)>,
    /// Total sentences when grammar scan started (for progress bar)
    grammar_queue_total: usize,
    /// Whether a grammar scan is in progress (shows indicator in UI)
    grammar_scanning: bool,
    /// Deferred find-and-replace (word, replacement, optional sentence context, doc char offset) — executed next frame
    pending_fix: Option<(String, String, String, usize)>,
    /// Pending consonant confusion candidates — validated with grammar checker after check_spelling
    pending_consonant_checks: Vec<WritingError>,
    /// Suggestion window: (misspelled_word, candidates)
    suggestion_window: Option<(String, Vec<(String, f32)>)>,
    suggestion_selection: std::sync::Arc<std::sync::Mutex<Option<usize>>>,
    /// Rule info popup: (rule_name, explanation, sentence_context, fix_idx, suggestion)
    rule_info_window: Option<(String, String, String, usize, String)>,
    // OCR clipboard monitoring
    ocr: Option<ocr::OcrClipboard>,
    ocr_receiver: Option<std::sync::mpsc::Receiver<Result<String, String>>>,
    ocr_text: Option<String>,
    // Microphone / Whisper
    whisper_engine: Option<Arc<Mutex<microphone::WhisperEngine>>>,
    mic_handle: Option<microphone::MicHandle>,
    mic_transcribing: bool,
    mic_result_text: Option<String>,
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

    let max_steps = if prefix_lower.len() <= 3 { 1 } else { 0 };
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
    fn new(grammar_completion: bool, use_swipl: bool, quality: u8) -> Self {
        #[cfg(target_os = "windows")]
        unsafe {
            use windows::Win32::System::Com::*;
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok();
        }

        let mut load_errors = Vec::new();

        // Grammar checker must load on main thread (SWI-Prolog requires it)
        let checker: Option<AnyChecker> = if use_swipl {
            match Self::load_swipl_checker() {
                Ok(c) => {
                    eprintln!("SWI-Prolog grammar checker loaded");
                    Some(AnyChecker::Swi(c))
                }
                Err(e) => {
                    let msg = format!("SWI-Prolog: {}", e);
                    eprintln!("{}", msg);
                    load_errors.push(msg);
                    match Self::load_checker() {
                        Ok(c) => {
                            eprintln!("Fallback: neorusticus loaded ({} clauses)", c.clause_count());
                            Some(AnyChecker::Neo(c))
                        }
                        Err(e2) => {
                            load_errors.push(format!("Grammar: {}", e2));
                            None
                        }
                    }
                }
            }
        } else {
            match Self::load_checker() {
                Ok(c) => {
                    eprintln!("GrammarChecker loaded ({} clauses)", c.clause_count());
                    Some(AnyChecker::Neo(c))
                }
                Err(e) => {
                    let msg = format!("Grammar: {}", e);
                    eprintln!("{}", msg);
                    load_errors.push(msg);
                    None
                }
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
            let model = model_opt.map(|m| Arc::new(Mutex::new(m)));
            let _ = tx2.send(StartupItem::Completer {
                model, prefix_index, baselines, wordfreq: wf, embedding_store, errors,
            });
        });

        // Thread 3: Whisper engine
        let tx3 = startup_tx;
        std::thread::spawn(move || {
            let dll_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../whisper-build/bin/Release")
                .to_string_lossy().to_string();
            let model_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../contexter-repo/training-data/ggml-nb-whisper-small.bin")
                .to_string_lossy().to_string();
            let _ = tx3.send(StartupItem::Whisper(
                microphone::WhisperEngine::load(&dll_dir, &model_path)
            ));
        });

        ContextApp {
            manager: BridgeManager::new(),
            context: CursorContext::default(),
            last_poll: Instant::now(),
            poll_interval: Duration::from_millis(300),
            follow_cursor: true,
            last_caret_pos: None,
            checker,
            analyzer: None,
            grammar_actor: None,
            grammar_errors: Vec::new(),
            last_checked_sentence: String::new(),
            model: None,
            completion_rx: None,
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
            word_hwnd: None,
            ctrl_space_held: false,
            selected_column: 0,
            load_errors,
            selected_tab: 0,
            writing_errors: Vec::new(),
            ignored_words: std::collections::HashSet::new(),
            last_spell_checked_word: String::new(),
            last_doc_text: String::new(),
            last_doc_hash: 0,
            last_sentence_count: 0,
            prolog_checked_hashes: std::collections::HashSet::new(),
            clean_sentence_hashes: std::collections::HashSet::new(),
            grammar_queue: Vec::new(),
            grammar_queue_total: 0,
            grammar_scanning: false,
            pending_fix: None,
            pending_consonant_checks: Vec::new(),
            suggestion_window: None,
            suggestion_selection: std::sync::Arc::new(std::sync::Mutex::new(None)),
            rule_info_window: None,
            ocr: match ocr::OcrClipboard::new() {
                Ok(o) => { eprintln!("OCR clipboard monitor ready"); Some(o) }
                Err(e) => { eprintln!("OCR not available: {}", e); None }
            },
            ocr_receiver: None,
            ocr_text: None,
            whisper_engine: None,
            mic_handle: None,
            mic_transcribing: false,
            mic_result_text: None,
            startup_rx: Some(startup_rx),
            startup_done: Vec::new(),
            startup_total: 2, // completer, whisper
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

    fn load_swipl_checker() -> Result<SwiGrammarChecker, Box<dyn std::error::Error>> {
        SwiGrammarChecker::new(
            swipl_dll_path(),
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
        eprintln!("NorBERT4 loaded. Vocab: {}", model.vocab_size());

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

        if let Some(checker) = &mut self.checker {
            let t = Instant::now();
            self.grammar_errors = checker.check_sentence(&sentence);
            if !self.grammar_errors.is_empty() {
                eprintln!("Grammar: {} errors in {:.0}ms", self.grammar_errors.len(), t.elapsed().as_secs_f64() * 1000.0);
            }
        }
    }

    /// Check spelling of a word. `sentence_ctx` is the sentence it appears in.
    fn check_spelling(&mut self, word: &str, sentence_ctx: &str, doc_offset: usize) {
        let clean = word.trim().to_lowercase();
        if clean.is_empty() || clean.len() < 2 || clean == self.last_spell_checked_word {
            return;
        }
        self.last_spell_checked_word = clean.clone();
        eprintln!("spell-check: '{}'", clean);

        // Skip if word is in ignore list
        if self.ignored_words.contains(&clean) {
            return;
        }

        // Skip punctuation-only or numbers
        if clean.chars().all(|c| c.is_ascii_punctuation() || c.is_ascii_digit()) {
            return;
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
                });
            }
            return;
        }

        // Phase 1: Dictionary lookups (immutable borrow on checker)
        let found;
        let kt_gt_valid_alt: Option<String>;
        let consonant_alts: Vec<String>;
        let compound_suggestion: Option<String>;
        let fuzzy_candidates: Vec<(String, u32)>;
        let original_found;

        {
            let checker = match &self.checker {
                Some(c) => c,
                None => return,
            };

            found = checker.has_word(&clean);
            eprintln!("  has_word('{}') = {}", clean, found);

            // kt/gt confusion: check if alt form exists
            kt_gt_valid_alt = if found {
                let alt = if clean.ends_with("kt") {
                    Some(format!("{}gt", &clean[..clean.len()-2]))
                } else if clean.ends_with("gt") {
                    Some(format!("{}kt", &clean[..clean.len()-2]))
                } else {
                    None
                };
                alt.filter(|a| checker.has_word(a))
            } else {
                None
            };

            // Consonant confusion: find valid alternatives with shared POS
            consonant_alts = if found && clean.len() >= 4 {
                let orig_pos = checker.pos_set(&clean);
                consonant_variants(&clean).into_iter()
                    .filter(|v| {
                        if !checker.has_word(v) { return false; }
                        let v_pos = checker.pos_set(v);
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

            original_found = if !found { checker.has_word(word.trim()) } else { false };

            compound_suggestion = if !found && !original_found {
                checker.suggest_compound(&clean)
            } else {
                None
            };

            fuzzy_candidates = if !found && !original_found && compound_suggestion.is_none() {
                checker.fuzzy_lookup(&clean, 2).into_iter()
                    .filter(|(w, _)| w != &clean)
                    .take(10)
                    .collect()
            } else {
                Vec::new()
            };
        } // checker borrow dropped

        // Phase 2: BERT scoring + writing errors (mutable borrow on self)
        if found {
            // kt/gt confusion — BERT sentence scoring
            if let Some(alt) = kt_gt_valid_alt {
                let s_orig = self.bert_sentence_score(sentence_ctx);
                let alt_sentence = sentence_ctx.replacen(&clean, &alt, 1);
                let s_alt = self.bert_sentence_score(&alt_sentence);
                log!("kt/gt check: '{}' score={:.2}, '{}' score={:.2}", clean, s_orig, alt, s_alt);
                if s_alt > s_orig {
                    if !self.writing_errors.iter().any(|e| {
                        matches!(e.category, ErrorCategory::Spelling) && e.word == clean && e.doc_offset == doc_offset && !e.ignored
                    }) {
                        log!("kt/gt confusion: '{}' → '{}'", clean, alt);
                        self.writing_errors.push(WritingError {
                            category: ErrorCategory::Spelling,
                            word: clean.clone(),
                            suggestion: alt.clone(),
                            explanation: format!("«{}» → «{}» (kt/gt-forveksling)", clean, alt),
                            rule_name: "kt_gt".to_string(),
                            sentence_context: sentence_ctx.to_string(),
                            doc_offset,
                            position: 0,
                            ignored: false,
                        });
                    }
                }
            }

            // Consonant confusion — BERT sentence scoring
            if !consonant_alts.is_empty() {
                let s_orig = self.bert_sentence_score(sentence_ctx);
                let mut best_alt: Option<(String, f32)> = None;
                for alt in &consonant_alts {
                    let alt_sentence = sentence_ctx.replacen(&clean, alt, 1);
                    let s_alt = self.bert_sentence_score(&alt_sentence);
                    log!("consonant BERT sentence: '{}' orig={:.2}, '{}' variant={:.2}", clean, s_orig, alt, s_alt);
                    if best_alt.as_ref().map_or(true, |(_, bs)| s_alt > *bs) {
                        best_alt = Some((alt.clone(), s_alt));
                    }
                }
                if let Some((best, s_best)) = best_alt {
                    log!("consonant check: '{}' score={:.2}, '{}' score={:.2}", clean, s_orig, best, s_best);
                    if s_best > s_orig {
                        let corrected_sentence = sentence_ctx.replacen(&clean, &best, 1);
                        self.pending_consonant_checks.push(WritingError {
                            category: ErrorCategory::Grammar,
                            word: sentence_ctx.to_string(),
                            suggestion: corrected_sentence,
                            explanation: format!("«{}» → «{}» (enkel/dobbel konsonant)", clean, best),
                            rule_name: format!("consonant_confusion:{}:{}", clean, best),
                            sentence_context: sentence_ctx.to_string(),
                            doc_offset,
                            position: 0,
                            ignored: false,
                        });
                    }
                }
            }

            return;
        }

        if original_found {
            return;
        }

        // Try compound suggestion via Prolog
        if let Some(compound) = compound_suggestion {
            log!("  Compound suggestion: '{}' → '{}'", clean, compound);
            if self.writing_errors.iter().any(|e| {
                matches!(e.category, ErrorCategory::Spelling) && e.word == clean && e.doc_offset == doc_offset && !e.ignored
            }) {
                return;
            }
            self.writing_errors.push(WritingError {
                category: ErrorCategory::Spelling,
                word: clean.clone(),
                suggestion: compound.clone(),
                explanation: format!("«{}» → «{}» (sammensatt ord)", clean, compound),
                rule_name: "stavefeil".to_string(),
                sentence_context: sentence_ctx.to_string(),
                doc_offset,
                position: 0,
                ignored: false,
            });
            return;
        }

        // Word not found — use fuzzy suggestions
        let mut candidates = fuzzy_candidates;

        // Boost by word frequency
        if let Some(wf) = &self.wordfreq {
            candidates.sort_by(|a, b| {
                let freq_a = wf.get(&a.0).copied().unwrap_or(0);
                let freq_b = wf.get(&b.0).copied().unwrap_or(0);
                // Primary: distance ascending, secondary: frequency descending
                a.1.cmp(&b.1).then(freq_b.cmp(&freq_a))
            });
        }

        let best = candidates.first().map(|(w, _)| w.clone()).unwrap_or_default();

        // Don't add duplicate errors for the same word at the same offset
        if self.writing_errors.iter().any(|e| {
            matches!(e.category, ErrorCategory::Spelling) && e.word == clean && e.doc_offset == doc_offset && !e.ignored
        }) {
            return;
        }

        let error = WritingError {
            category: ErrorCategory::Spelling,
            word: clean.clone(),
            suggestion: best,
            explanation: format!("«{}» finnes ikke i ordboken.", clean),
            rule_name: "stavefeil".to_string(),
            sentence_context: sentence_ctx.to_string(),
            doc_offset,
            position: 0,
            ignored: false,
        };
        self.writing_errors.push(error);
        eprintln!("Spelling: '{}' not found, suggesting '{}'",
            clean, self.writing_errors.last().unwrap().suggestion);
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
        let checker = match &mut self.checker {
            Some(c) => c,
            None => return,
        };
        for mut candidate in pending {
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

            // Grammar check original sentence vs corrected sentence
            let orig_errors = checker.check_sentence(&candidate.word);
            let variant_errors = checker.check_sentence(&candidate.suggestion);
            log!("consonant grammar validate: '{}' → '{}' | orig_errors={}, variant_errors={}",
                orig_word, variant_word, orig_errors.len(), variant_errors.len());

            // Clean up rule_name for display
            candidate.rule_name = "consonant_confusion".to_string();

            if variant_errors.len() <= orig_errors.len() {
                // BERT already decided the variant is better — accept unless grammar gets worse
                log!("consonant confirmed: '{}' → '{}' (BERT preferred, grammar ok)", orig_word, variant_word);
                self.writing_errors.push(candidate);
            } else {
                log!("consonant rejected: '{}' → '{}' (grammar worse)", orig_word, variant_word);
            }
        }
    }

    /// Upgrade spelling error suggestions using BERT context.
    /// Called after update_grammar_errors() to replace fuzzy-only suggestions
    /// with contextually appropriate ones (e.g. "bossller" → "boller" not "fossiler").
    fn upgrade_spelling_suggestions(&mut self) {
        // Wait for BERT model — without it, scoring falls back to trigrams which can't distinguish candidates
        if self.model.is_none() { return; }

        // Collect indices + data for spelling errors that need upgrading
        let to_upgrade: Vec<(usize, String, String)> = self.writing_errors.iter().enumerate()
            .filter(|(_, e)| {
                matches!(e.category, ErrorCategory::Spelling)
                    && !e.ignored
                    && e.rule_name == "stavefeil" // not yet upgraded
                    && !e.sentence_context.is_empty()
            })
            .map(|(i, e)| (i, e.word.clone(), e.sentence_context.clone()))
            .collect();

        for (idx, word, sentence_ctx) in to_upgrade {
            let existing = self.writing_errors[idx].suggestion.clone();
            let mut suggestions = self.trigram_suggestions(&word, &sentence_ctx);
            // Ensure existing suggestion is in the candidate list (it's dictionary-confirmed)
            if !existing.is_empty() {
                if !suggestions.iter().any(|(w, _)| *w == existing) {
                    suggestions.push((existing.clone(), 0.0)); // will be re-scored by BERT below
                }
            }
            // Pick best suggestion that doesn't introduce grammar errors
            if !suggestions.is_empty() {
                for (i, (w, s)) in suggestions.iter().take(5).enumerate() {
                    log!("  #{}: '{}' score={:.2}", i+1, w, s);
                }
                let mut picked: Option<(String, f32)> = None;
                if let Some(checker) = &mut self.checker {
                    for (candidate, score) in suggestions.iter().take(8) {
                        let corrected = sentence_ctx.replacen(&word, candidate, 1);
                        let errors = checker.check_sentence(&corrected);
                        log!("Spelling grammar-check: '{}' in '{}' → {} errors", candidate, corrected, errors.len());
                        if errors.is_empty() {
                            picked = Some((candidate.clone(), *score));
                            break;
                        }
                    }
                }
                // Fallback to BERT-best if none pass grammar
                if picked.is_none() {
                    picked = suggestions.first().map(|(w, s)| (w.clone(), *s));
                }
                if let Some((best, score)) = &picked {
                    log!("Spelling upgrade: '{}' → '{}' score={:.2} (was '{}', {} candidates)",
                        word, best, score, existing, suggestions.len());
                    self.writing_errors[idx].suggestion = best.clone();
                }
            }
            // Mark as processed regardless so we don't retry
            self.writing_errors[idx].rule_name = "stavefeil_bert".to_string();
        }
    }

    /// Find spelling suggestion candidates and rank them with BERT sentence scoring.
    /// 1. Collect candidates from all sources (fuzzy, prefix, compound, wordfreq)
    /// 2. Score each by encoding the word and summing BERT logits at the masked position
    /// 3. Tiebreaker: prefix match wins when BERT scores are close
    fn trigram_suggestions(&mut self, word: &str, sentence_ctx: &str) -> Vec<(String, f32)> {
        let word_lower = word.to_lowercase();
        let word_trigrams = Self::trigrams(&word_lower);
        let word_first = word_lower.chars().next().unwrap_or(' ');

        // ── Collect unique candidates from all sources ──
        let mut candidates: Vec<String> = Vec::new();
        let mut seen = std::collections::HashSet::new();
        let mut edit_distances: HashMap<String, u32> = HashMap::new();
        let mut add = |w: String, seen: &mut std::collections::HashSet<String>| {
            let wl = w.to_lowercase();
            if wl != word_lower && wl.len() >= 2 && seen.insert(wl.clone()) {
                // Pre-filter: need trigram overlap or same first letter
                let w_trigrams = Self::trigrams(&wl);
                let common = word_trigrams.iter().filter(|t| w_trigrams.contains(t)).count();
                if common > 0 || wl.chars().next().unwrap_or(' ') == word_first {
                    return Some(wl);
                }
            }
            None
        };

        // Source 1: Fuzzy Levenshtein matches (distance 2)
        if let Some(checker) = &self.checker {
            let fuzzy = checker.fuzzy_lookup(&word_lower, 2);
            eprintln!("Forslag: fuzzy(2) returned {} matches for '{}'", fuzzy.len(), word_lower);
            for (w, dist) in fuzzy {
                let wl = w.to_lowercase();
                edit_distances.insert(wl, dist);
                if let Some(wl) = add(w, &mut seen) { candidates.push(wl); }
            }
        }

        // Source 2: Prefix lookup (missing-letter typos: "fotbal" → "fotball")
        let mut prefix_matches: std::collections::HashSet<String> = std::collections::HashSet::new();
        if let Some(checker) = &self.checker {
            for w in checker.prefix_lookup(&word_lower, 20) {
                let wl = w.to_lowercase();
                let extra = wl.len() as i32 - word_lower.len() as i32;
                if extra >= 1 && extra <= 3 {
                    prefix_matches.insert(wl.clone());
                    // Prefix match = insertion of extra chars, edit distance = extra chars
                    edit_distances.entry(wl.clone()).or_insert(extra as u32);
                    if let Some(wl) = add(w, &mut seen) { candidates.push(wl); }
                }
            }
        }

        // Source 3: Prefix with last char removed (typo in final position)
        if word_lower.len() >= 3 {
            let shorter = &word_lower[..word_lower.len()-1];
            if let Some(checker) = &self.checker {
                for w in checker.prefix_lookup(shorter, 20) {
                    let wl = w.to_lowercase();
                    // Approximate edit distance: removed 1 char + added extra
                    let diff = (wl.len() as i32 - word_lower.len() as i32).unsigned_abs() + 1;
                    edit_distances.entry(wl).or_insert(diff);
                    if let Some(wl) = add(w, &mut seen) { candidates.push(wl); }
                }
            }
        }

        // Source 4: Wordfreq — common words with trigram overlap
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

        eprintln!("Forslag: {} candidates for '{}'", candidates.len(), word_lower);

        // ── Score each candidate with BERT ──
        // Build masked context: replace the misspelled word with <mask>
        let sentence_lower = sentence_ctx.to_lowercase();
        let masked_context = if let Some(pos) = sentence_lower.find(&word_lower) {
            let before = &sentence_ctx[..pos];
            let after = &sentence_ctx[pos + word_lower.len()..];
            format!("{}<mask>{}", before.trim_end(), after)
        } else {
            format!("{} <mask>", sentence_ctx)
        };
        eprintln!("Forslag: masked context = '{}'", masked_context);

        let mut scored: Vec<(String, f32)> = Vec::new();

        if let Some(model_arc) = &self.model {
            let mut model = model_arc.lock().unwrap();
            if let Ok((logits, _ms)) = model.single_forward(&masked_context) {
                // Score = BERT logit score × trigram similarity
                // BERT ensures contextual fit, trigram ensures orthographic closeness
                for w in &candidates {
                    let bert_score = if let Ok(enc) = model.tokenizer.encode(w.as_str(), false) {
                        let ids = enc.get_ids();
                        if ids.is_empty() { 0.0 }
                        else {
                            ids.iter()
                                .map(|&id| logits.get(id as usize).copied().unwrap_or(0.0))
                                .sum::<f32>()
                                / ids.len() as f32
                        }
                    } else { 0.0 };

                    let w_trigrams = Self::trigrams(w);
                    let common = word_trigrams.iter().filter(|t| w_trigrams.contains(t)).count();
                    let max_t = word_trigrams.len().max(w_trigrams.len()).max(1);
                    let trigram_sim = common as f32 / max_t as f32;

                    // Orthographic similarity: combine trigram and edit distance
                    // Edit distance is more reliable for transpositions/insertions
                    // dist 1 → 0.85, dist 2 → 0.65, unknown → use trigram only
                    let edit_sim = match edit_distances.get(w) {
                        Some(1) => 0.85,
                        Some(2) => 0.65,
                        _ => 0.0,
                    };
                    // Take the best of trigram or edit-distance similarity
                    let ortho_sim = trigram_sim.max(edit_sim);

                    // Prefix match bonus: the fewer extra chars, the stronger the signal
                    let prefix_bonus = if prefix_matches.contains(w) {
                        let extra = w.len() as f32 - word_lower.len() as f32;
                        if extra <= 1.0 { 1.5 } else if extra <= 2.0 { 1.2 } else { 1.1 }
                    } else { 1.0 };

                    // Combined: BERT × orthographic similarity × prefix bonus
                    let score = bert_score.max(0.0) * ortho_sim * prefix_bonus;

                    scored.push((w.clone(), score));
                }
            }
        }

        // If no BERT model, fall back to trigram-only scoring
        if scored.is_empty() {
            for w in &candidates {
                let w_trigrams = Self::trigrams(w);
                let common = word_trigrams.iter().filter(|t| w_trigrams.contains(t)).count();
                let max_t = word_trigrams.len().max(w_trigrams.len()).max(1);
                scored.push((w.clone(), common as f32 / max_t as f32));
            }
        }

        scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

        eprintln!("Forslag: top BERT×trigram for '{}': {:?}", word_lower,
            scored.iter().take(5).collect::<Vec<_>>());

        scored.truncate(10);
        scored
    }

    /// Score a sentence using BERT pseudo-log-likelihood.
    /// For each word position, mask it and check how well BERT predicts the actual word.
    fn bert_sentence_score(&mut self, sentence: &str) -> f32 {
        let words: Vec<&str> = sentence.split_whitespace().collect();
        if words.is_empty() { return f32::NEG_INFINITY; }

        let model_arc = match &self.model {
            Some(m) => m,
            None => return 0.0,
        };
        let mut model = match model_arc.try_lock() {
            Ok(m) => m,
            Err(_) => return 0.0, // model busy (background forward) — skip scoring
        };

        let mut total_score: f32 = 0.0;
        for i in 0..words.len() {
            // Build masked sentence
            let masked: String = words.iter().enumerate()
                .map(|(j, w)| if j == i { "<mask>" } else { *w })
                .collect::<Vec<_>>()
                .join(" ");

            if let Ok((logits, _)) = model.single_forward(&masked) {
                // Look up the actual word's token and get its logit
                let word_clean = words[i].trim_matches(|c: char| c.is_ascii_punctuation());
                // Try with Ġ prefix (word-initial BPE token)
                let token_with_g = format!("Ġ{}", word_clean.to_lowercase());
                let token_id = model.tokenizer.token_to_id(&token_with_g)
                    .or_else(|| model.tokenizer.token_to_id(&word_clean.to_lowercase()));
                if let Some(tid) = token_id {
                    total_score += logits[tid as usize];
                }
            }
        }
        // Normalize by word count to avoid bias toward longer sentences
        total_score / words.len() as f32
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
                        if let Some(checker) = &self.checker {
                            if checker.has_word(&variant) {
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

        // Grammar-check each candidate: discard ones that still have errors
        // Also verify all words in the correction exist in the dictionary
        let valid_candidates: Vec<(String, String, String)> = if let Some(checker) = &mut self.checker {
            candidates.into_iter()
                .filter(|(c, _, _)| {
                    let errs = checker.check_sentence(c);
                    if !errs.is_empty() {
                        eprintln!("    REJECTED (grammar): '{}' — {} errors", c, errs.len());
                        return false;
                    }
                    // Check all words exist in dictionary (catches misspelled suggestions like "spile")
                    for word in c.split_whitespace() {
                        let clean = word.trim_matches(|ch: char| ch.is_ascii_punctuation() || ch == '\u{00ab}' || ch == '\u{00bb}').to_lowercase();
                        if clean.len() >= 2 && !clean.chars().all(|ch| ch.is_ascii_digit()) && !checker.has_word(&clean) {
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
            eprintln!("  No grammatically valid candidates found");
            return Vec::new();
        }

        // Score valid candidates with BERT
        eprintln!("  Scoring {} valid candidates with BERT...", valid_candidates.len());
        let mut scored: Vec<(String, String, String, f32)> = valid_candidates.into_iter()
            .map(|(c, e, r)| {
                let score = self.bert_sentence_score(&c);
                eprintln!("    {:.1}: '{}'", score, c);
                (c, e, r, score)
            })
            .collect();

        // Sort by BERT score — best correction wins regardless of type
        scored.sort_by(|a, b| b.3.partial_cmp(&a.3).unwrap_or(std::cmp::Ordering::Equal));
        scored.truncate(1);
        scored
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

    /// Remove errors whose word has been corrected in the document.
    fn prune_resolved_errors(&mut self) {
        // Only prune with a FRESH document read — stale cache causes false pruning
        let doc_text = match self.manager.read_full_document() {
            Some(t) => {
                self.last_doc_text = t.clone();
                t.to_lowercase()
            }
            None => return, // Can't read doc (our window focused) — skip pruning
        };
        self.writing_errors.retain(|e| {
            if e.ignored {
                log!("Pruning ignored: {:?} '{}'", e.category, &e.word[..e.word.len().min(40)]);
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
                    // Resolved when the original unpunctuated text no longer matches
                    // (user accepted the fix or changed the text)
                    doc_text.contains(&e.word.to_lowercase())
                }
            };
            if !still_present {
                log!("Error resolved: {:?} '{}' no longer in document", e.category, &e.word[..e.word.len().min(40)]);
            }
            still_present
        });
    }

    /// Prepare grammar scan: read document, split sentences, compute offsets, fill queue.
    /// This is fast (no SWI/BERT calls) and runs every poll when document changes.
    /// The actual per-sentence grammar checking happens incrementally in process_grammar_queue().
    fn update_grammar_errors(&mut self) {
        // Called on paste/cut/move only — not on every keystroke.
        // Queue processing happens at word boundaries in the main poll loop.

        // Read document text and check all complete sentences
        let doc_text = match self.manager.read_full_document() {
            Some(t) => { self.last_doc_text = t.clone(); t }
            None => {
                // Can't read (our window focused?) — use cached text
                if self.last_doc_text.is_empty() { return; }
                self.last_doc_text.clone()
            }
        };

        // Quick check: if document hasn't changed at all, skip everything
        let doc_hash = hash_str(&doc_text);
        if doc_hash == self.last_doc_hash {
            return;
        }
        self.last_doc_hash = doc_hash;

        // Count sentences to detect paste (large change) vs fix (small change)
        let new_sentence_count = split_sentences(&doc_text).len();
        let old_sentence_count = self.last_sentence_count;
        self.last_sentence_count = new_sentence_count;
        let is_major_change = (new_sentence_count as isize - old_sentence_count as isize).unsigned_abs() > 2;

        if is_major_change {
            // Paste/delete — clean hashes are kept (same sentence text = same grammar result,
            // regardless of position). No need to rescan sentences already known clean.
            log!("Major doc change: {} → {} sentences (keeping {} clean hashes)",
                old_sentence_count, new_sentence_count, self.clean_sentence_hashes.len());
        }

        let mut sentences = split_sentences(&doc_text);
        // Track which original sentences were sub-split by Prolog
        // (original_text → split_sentences) for boundary suggestions
        let mut prolog_splits: Vec<(String, Vec<String>)> = Vec::new();

        // If no punctuated sentences but text exists, try Prolog sentence splitting
        if sentences.is_empty() && nostos_cognio::punctuation::needs_punctuation_check(&doc_text) {
            let doc_h = hash_str(&doc_text);
            if !self.prolog_checked_hashes.contains(&doc_h) {
                if let Some(checker) = &mut self.checker {
                    if let Some(prolog_sentences) = checker.split_by_prolog(&doc_text) {
                        eprintln!("Grammar: Prolog split {} sentences from fully unpunctuated text", prolog_sentences.len());
                        prolog_splits.push((doc_text.clone(), prolog_sentences.clone()));
                        sentences = prolog_sentences;
                    }
                }

                // Fallback to BERT if Prolog found nothing
                if sentences.is_empty() {
                    if let Some(model_arc) = &self.model {
                        let verb_fn: Option<Box<dyn Fn(&str) -> bool>> = match &self.checker {
                            Some(AnyChecker::Swi(c)) => {
                                let analyzer = c.analyzer().clone();
                                Some(Box::new(move |word: &str| -> bool {
                                    nostos_cognio::punctuation::is_finite_verb_mtag(&analyzer, word)
                                }))
                            }
                            _ => None,
                        };

                        let verb_ref: Option<&dyn Fn(&str) -> bool> = verb_fn.as_deref();
                        let mut model = model_arc.lock().unwrap();
                        match nostos_cognio::punctuation::split_into_sentences_with_verbs(&mut *model, &doc_text, 10.0, verb_ref) {
                            Ok(predicted) => {
                                sentences = predicted;
                            }
                            Err(e) => {
                                eprintln!("Grammar: BERT punctuation prediction failed: {}", e);
                            }
                        }
                    }
                }
                self.prolog_checked_hashes.insert(doc_h);
            }
        }

        // Also check each punctuated sentence for internal boundaries
        // e.g. "Jeg spiller fotball jeg går tur." — has final period but missing internal one
        if let Some(checker) = &mut self.checker {
            let mut expanded: Vec<String> = Vec::new();
            for sent in &sentences {
                let sent_h = hash_str(sent);
                if self.prolog_checked_hashes.contains(&sent_h) {
                    // Already checked for sub-splitting — just pass through
                    expanded.push(sent.clone());
                    continue;
                }
                // Strip trailing punctuation for Prolog analysis
                let stripped = sent.trim_end_matches(|c: char| c == '.' || c == '!' || c == '?').trim();
                if stripped.split_whitespace().count() >= 4 {
                    if let Some(sub_sentences) = checker.split_by_prolog(stripped) {
                        eprintln!("Grammar: Prolog sub-split '{}' into {} sentences",
                            &stripped[..stripped.len().min(50)], sub_sentences.len());
                        prolog_splits.push((sent.clone(), sub_sentences.clone()));
                        self.prolog_checked_hashes.insert(sent_h);
                        expanded.extend(sub_sentences);
                        continue;
                    }
                }
                self.prolog_checked_hashes.insert(sent_h);
                expanded.push(sent.clone());
            }
            sentences = expanded;
        }

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
                    let char_offset = doc_text[..abs_byte].chars().count();
                    if !claimed_offsets.contains(&char_offset) {
                        found_offset = Some(char_offset);
                        claimed_offsets.push(char_offset);
                        break;
                    }
                    search_from = abs_byte + 1;
                }
                result.push((s.clone(), found_offset.unwrap_or(0)));
            }
            result
        };

        // --- Step 0: Re-map existing errors to new offsets, remove stale ones ---
        {
            let mut available_offsets: std::collections::HashMap<String, Vec<usize>> = std::collections::HashMap::new();
            for (s, off) in &new_sentences {
                available_offsets.entry(s.clone()).or_default().push(*off);
            }
            let mut claimed: std::collections::HashMap<String, Vec<usize>> = std::collections::HashMap::new();
            for e in &mut self.writing_errors {
                if e.ignored { continue; }
                let key = e.sentence_context.clone();
                if let Some(offsets) = available_offsets.get(&key) {
                    let already_claimed = claimed.entry(key.clone()).or_default();
                    if let Some(&off) = offsets.iter().find(|o| !already_claimed.contains(o)) {
                        e.doc_offset = off;
                        already_claimed.push(off);
                    } else {
                        e.ignored = true;
                        log!("Removed stale error: '{}' (no matching position)", &e.word[..e.word.len().min(40)]);
                    }
                } else {
                    e.ignored = true;
                    log!("Removed stale error: '{}' (sentence gone)", &e.word[..e.word.len().min(40)]);
                }
            }
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
                &original_text[..original_text.len().min(60)],
                &punctuated[..punctuated.len().min(60)]);

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
            });
        }

        // --- Step 2: Check each sentence — skip already-known ones, only scan new ---
        let mut queue: Vec<(String, usize)> = Vec::new();
        for (trimmed, doc_offset) in &new_sentences {
            let sent_h = hash_str(trimmed);

            // Already seen this sentence text? Position updated in Step 0, skip entirely.
            if self.clean_sentence_hashes.contains(&sent_h) {
                continue;
            }
            // Also skip if this occurrence already has errors recorded (re-mapped in Step 0)
            let has_errors = self.writing_errors.iter().any(|e| {
                e.sentence_context == *trimmed && e.doc_offset == *doc_offset && !e.ignored
            });
            if has_errors {
                continue;
            }

            // New sentence — run spelling + queue for grammar
            for word in trimmed.split_whitespace() {
                let clean = word.trim_matches(|c: char| c.is_ascii_punctuation() || c == '\u{00ab}' || c == '\u{00bb}');
                if !clean.is_empty() {
                    self.check_spelling(clean, trimmed, *doc_offset);
                }
            }
            self.validate_consonant_checks();

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

    /// Process ONE sentence from the grammar queue per call.
    /// This keeps the UI responsive — each call does ~5-50ms of work.
    fn process_grammar_queue(&mut self) {
        let (trimmed, doc_offset) = match self.grammar_queue.first() {
            Some(item) => item.clone(),
            None => {
                self.grammar_scanning = false;
                return;
            }
        };
        self.grammar_queue.remove(0);

        if self.grammar_queue.is_empty() {
            self.grammar_scanning = false;
        }

        let sent_h = hash_str(&trimmed);

        // Re-check: skip if already has errors (may have been added by spelling in preparation)
        let has_errors = self.writing_errors.iter().any(|e| {
            e.sentence_context == trimmed && e.doc_offset == doc_offset && !e.ignored
        });
        if has_errors {
            return;
        }

        log!("Grammar check: '{}' (offset={}, {} remaining)", trimmed, doc_offset, self.grammar_queue.len());

        let checker = match &mut self.checker {
            Some(c) => c,
            None => return,
        };

        let errors = checker.check_sentence(&trimmed);
        if errors.is_empty() {
            // Mark as clean so we don't re-check next poll
            self.clean_sentence_hashes.insert(sent_h);
            return;
        }

        for ge in &errors {
            log!("  Grammar error: '{}' → '{}' ({})", ge.word, ge.suggestion, ge.rule_name);
        }

        // Score candidates with BERT (only runs when Prolog found errors)
        let corrections = self.best_sentence_corrections(&trimmed, &errors);

        if corrections.is_empty() {
            // No BERT-scored correction — fall back to direct grammar suggestions
            let errors_with_suggestions: Vec<_> = errors.iter()
                .filter(|e| !e.suggestion.is_empty())
                .collect();
            if !errors_with_suggestions.is_empty() {
                for (i, ge) in errors_with_suggestions.iter().enumerate() {
                    let first_alt = ge.suggestion.split('|').next().unwrap_or(&ge.suggestion);
                    let corrected = replace_word_at_position(&trimmed, &ge.word, first_alt);
                    if corrected.trim() == trimmed.trim() {
                        continue;
                    }
                    log!("  Direct grammar fix: '{}' → '{}' [{}]", ge.word, first_alt, ge.rule_name);
                    self.writing_errors.push(WritingError {
                        category: ErrorCategory::Grammar,
                        word: trimmed.to_string(),
                        suggestion: corrected,
                        explanation: format!("«{}» → «{}»: {}", ge.word, first_alt, ge.explanation),
                        rule_name: ge.rule_name.clone(),
                        sentence_context: trimmed.to_string(),
                        doc_offset,
                        position: i,
                        ignored: false,
                    });
                }
            } else {
                let first = &errors[0];
                log!("  Flagging without correction: '{}' ({})", first.word, first.rule_name);
                self.writing_errors.push(WritingError {
                    category: ErrorCategory::Grammar,
                    word: trimmed.to_string(),
                    suggestion: String::new(),
                    explanation: first.explanation.clone(),
                    rule_name: first.rule_name.clone(),
                    sentence_context: trimmed.to_string(),
                    doc_offset,
                    position: 0,
                    ignored: false,
                });
            }
        }

        for (i, (corrected, explanation, rule_name, score)) in corrections.iter().enumerate() {
            if corrected.trim() == trimmed.trim() {
                log!("  Skipping no-op correction: '{}'", corrected);
                continue;
            }
            log!("  Correction #{}: ({:.1}) '{}' -> '{}' [{}]", i+1, score, &trimmed, corrected, rule_name);
            self.writing_errors.push(WritingError {
                category: ErrorCategory::Grammar,
                word: trimmed.to_string(),
                suggestion: corrected.clone(),
                explanation: explanation.clone(),
                rule_name: rule_name.clone(),
                sentence_context: trimmed.to_string(),
                doc_offset,
                position: i,
                ignored: false,
            });
        }
    }

    fn sync_embeddings(&mut self) {
        if self.last_embedding_sync.elapsed() < self.embedding_sync_interval {
            return;
        }
        self.last_embedding_sync = Instant::now();

        if let Some(store) = &mut self.embedding_store {
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
            let fallback_fn: Option<Box<dyn Fn(&str) -> bool + '_>> = self.checker.as_ref().map(|c| {
                Box::new(move |word: &str| c.has_word(word)) as Box<dyn Fn(&str) -> bool>
            });
            let fallback_ref: Option<&dyn Fn(&str) -> bool> = fallback_fn.as_ref().map(|b| b.as_ref());
            let prefix_fn: Option<Box<dyn Fn(&str, usize) -> Vec<String> + '_>> = self.checker.as_ref().map(|c| {
                Box::new(move |p: &str, limit: usize| c.prefix_lookup(p, limit)) as Box<dyn Fn(&str, usize) -> Vec<String>>
            });
            let prefix_ref: Option<&dyn Fn(&str, usize) -> Vec<String>> = prefix_fn.as_ref().map(|b| b.as_ref());

            if let (Some(model_arc), Some(pi)) = (&self.model, &self.prefix_index) {
                let mut model = model_arc.lock().unwrap();
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
                    self.baselines.as_ref(),
                    self.wordfreq.as_deref(),
                    fallback_ref,
                    prefix_ref,
                    self.embedding_store.as_ref(),
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

        // Grammar filter (borrows checker mutably) — only when grammar_completion enabled
        if let Some((results, ctx, bert_ms)) = raw_results {
            if self.grammar_completion {
                if let Some(checker) = &mut self.checker {
                    let t_gram = Instant::now();
                    let mut check_fn = |sentence: &str| -> GrammarCheckResult {
                        let errors = checker.check_sentence(sentence);
                        GrammarCheckResult {
                            ok: errors.is_empty(),
                            suggestions: errors.iter()
                                .filter(|e| !e.suggestion.is_empty())
                                .map(|e| e.suggestion.clone())
                                .collect(),
                        }
                    };
                    let filtered = grammar_filter(
                        &results, &ctx, prefix,
                        &mut check_fn,
                        5,
                    );
                    let gram_ms = t_gram.elapsed().as_millis();
                    let total_ms = t_total.elapsed().as_millis();
                    eprintln!("complete '...{}' bert={}ms gram={}ms total={}ms -> {}",
                        prefix, bert_ms, gram_ms, total_ms,
                        filtered.iter().map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "));
                    self.completions = filtered;
                } else {
                    self.completions = results.into_iter().take(5).collect();
                }
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
        if let Some(hwnd_val) = self.word_hwnd {
            use windows::Win32::Foundation::HWND;
            use windows::Win32::UI::WindowsAndMessaging::SetForegroundWindow;
            unsafe {
                let hwnd = HWND(hwnd_val as *mut _);
                let _ = SetForegroundWindow(hwnd);
                // Give the OS time to switch focus
                std::thread::sleep(std::time::Duration::from_millis(100));
            }
        }
    }
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
    sentences
}

fn get_screen_size() -> (f32, f32) {
    #[cfg(target_os = "windows")]
    unsafe {
        use windows::Win32::UI::WindowsAndMessaging::*;
        let w = GetSystemMetrics(SM_CXSCREEN);
        let h = GetSystemMetrics(SM_CYSCREEN);
        return (w as f32, h as f32);
    }
    #[allow(unreachable_code)]
    (1920.0, 1080.0)
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
        // Execute deferred find-and-replace
        if let Some((find, replace, context, doc_offset)) = self.pending_fix.take() {
            log!("pending_fix: bridge='{}' find='{}' replace='{}' offset={}",
                self.manager.active_bridge_name(),
                &find[..find.len().min(60)], &replace[..replace.len().min(60)], doc_offset);
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
                // Document changed — reset doc hash so next poll re-scans
                self.last_doc_hash = 0;
                // Clear grammar queue — document changed, stale sentences
                self.grammar_queue.clear();
                self.grammar_scanning = false;
                // Mark the replacement sentences as clean AND prolog-checked
                // so they don't get re-flagged or re-split
                let mark_clean = |text: &str, clean: &mut std::collections::HashSet<u64>, prolog: &mut std::collections::HashSet<u64>| {
                    let h = hash_str(text);
                    clean.insert(h);
                    prolog.insert(h);
                    // Also mark without trailing punctuation (Prolog strips it)
                    let stripped = text.trim_end_matches(|c: char| c == '.' || c == '!' || c == '?').trim();
                    if !stripped.is_empty() && stripped != text {
                        let sh = hash_str(stripped);
                        clean.insert(sh);
                        prolog.insert(sh);
                    }
                };
                // Mark the full replacement
                mark_clean(&replace, &mut self.clean_sentence_hashes, &mut self.prolog_checked_hashes);
                // Mark each sub-sentence within the replacement
                for sent in replace.split_inclusive(|c: char| c == '.' || c == '!' || c == '?') {
                    let trimmed = sent.trim();
                    if !trimmed.is_empty() {
                        mark_clean(trimmed, &mut self.clean_sentence_hashes, &mut self.prolog_checked_hashes);
                    }
                }
                // Remove only the specific error that was fixed (matching text + offset)
                let find_lower = find.to_lowercase();
                let mut removed_one = false;
                self.writing_errors.retain(|e| {
                    if removed_one { return true; }
                    if (e.word.to_lowercase() == find_lower || e.sentence_context.to_lowercase() == find_lower)
                        && e.doc_offset == doc_offset
                    {
                        removed_one = true;
                        return false;
                    }
                    true
                });
                log!("Fix applied: marked {} clean, {} prolog-checked",
                    self.clean_sentence_hashes.len(), self.prolog_checked_hashes.len());
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
                            tts::speak_word(&text);
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
                        self.model = model;
                        self.prefix_index = prefix_index;
                        self.baselines = baselines;
                        self.wordfreq = wordfreq;
                        self.embedding_store = embedding_store;
                        self.load_errors.extend(errors);
                        self.startup_done.push("NorBERT4".into());
                        eprintln!("Startup: NorBERT4 completer ready");
                    }
                    StartupItem::Whisper(result) => {
                        match result {
                            Ok(engine) => {
                                self.whisper_engine = Some(Arc::new(Mutex::new(engine)));
                                self.startup_done.push("Whisper".into());
                                eprintln!("Startup: Whisper engine ready");
                            }
                            Err(e) => {
                                self.load_errors.push(format!("Whisper: {}", e));
                                self.startup_done.push("Whisper (feil)".into());
                                eprintln!("Whisper engine failed to load: {}", e);
                            }
                        }
                    }
                }
            }
            if self.startup_done.len() >= self.startup_total {
                self.startup_rx = None;
            }
        }

        // Microphone: check if whisper transcription finished
        if let Some(handle) = &self.mic_handle {
            if let Ok(result) = handle.result_rx.try_recv() {
                eprintln!("Whisper transcription complete: '{}'", result.text);
                self.mic_result_text = Some(result.text);
                self.mic_handle = None;
                self.mic_transcribing = false;
            }
        }

        // Poll for new context
        if self.last_poll.elapsed() >= self.poll_interval {
            self.last_poll = Instant::now();

            if let Some(new_ctx) = self.manager.read_context() {
                // Save the foreground window only when we got useful context from it
                if !new_ctx.word.is_empty() || !new_ctx.sentence.is_empty() {
                    use windows::Win32::UI::WindowsAndMessaging::GetForegroundWindow;
                    let fg = unsafe { GetForegroundWindow() };
                    if fg.0 as isize != 0 {
                        let our_title = "NorskTale";
                        let mut buf = [0u16; 64];
                        let len = unsafe {
                            windows::Win32::UI::WindowsAndMessaging::GetWindowTextW(fg, &mut buf)
                        };
                        let title = String::from_utf16_lossy(&buf[..len as usize]);
                        if !title.contains(our_title)
                            && !title.starts_with("Forslag")
                            && !title.starts_with("Regelinfo")
                        {
                            self.word_hwnd = Some(fg.0 as isize);
                            self.manager.set_target_hwnd(fg.0 as isize);
                        }
                    }
                }
                if new_ctx.caret_pos.is_some() {
                    self.last_caret_pos = new_ctx.caret_pos;
                }
                // Only update context if we got something useful — don't overwrite
                // good context with empty when our own window is focused
                if !new_ctx.word.is_empty() || !new_ctx.sentence.is_empty() || new_ctx.masked_sentence.is_some() {
                    // Cursor moved — clear stale grammar queue
                    if new_ctx.masked_sentence != self.context.masked_sentence && !self.grammar_queue.is_empty() {
                        eprintln!("Cursor moved — clearing {} stale grammar queue items", self.grammar_queue.len());
                        self.grammar_queue.clear();
                        self.grammar_scanning = false;
                    }
                    // Update doc text cache from masked sentence (strip <mask> to get real text)
                    // Detect paste/cut/move: large jump in text length triggers full doc scan
                    if let Some(ref masked) = new_ctx.masked_sentence {
                        let doc_approx = masked.replace("<mask>", &new_ctx.word);
                        let old_len = self.last_doc_text.len();
                        let new_len = doc_approx.len();
                        let big_change = old_len == 0 || (new_len as isize - old_len as isize).unsigned_abs() > 20;
                        if doc_approx.len() > self.last_doc_text.len() / 2 {
                            self.last_doc_text = doc_approx;
                        }
                        if big_change {
                            // Paste/cut/move detected — trigger full document grammar scan
                            self.update_grammar_errors();
                        }
                    }
                    self.context = new_ctx;
                }
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
                // Word boundary: check spelling of the last finished word
                let sentence = self.context.sentence.clone();
                let spell_word = sentence.split_whitespace().last()
                    .map(|w| w.trim_matches(|c: char| c.is_ascii_punctuation() || c == '«' || c == '»').to_string());
                if let Some(ref w) = spell_word {
                    if !w.is_empty() {
                        self.check_spelling(w, &sentence, 0);
                    }
                }
                self.validate_consonant_checks();
                // Sentence boundary: run grammar check
                self.run_grammar_check();
                // Word boundary work: prune, upgrade, drain grammar queue
                // Only process grammar queue when no background forward in flight (avoids model mutex contention)
                self.prune_resolved_errors();
                self.upgrade_spelling_suggestions();
                if !self.grammar_queue.is_empty() && self.completion_rx.is_none() {
                    self.process_grammar_queue();
                }
            } else {
                // No word, no context: clear and run grammar
                self.completions.clear();
                self.open_completions.clear();
                self.last_completed_prefix.clear();
                // Check spelling + grammar on the last word/sentence
                let sentence = self.context.sentence.clone();
                let spell_word = sentence.split_whitespace().last()
                    .map(|w| w.trim_matches(|c: char| c.is_ascii_punctuation() || c == '«' || c == '»').to_string());
                if let Some(ref w) = spell_word {
                    if !w.is_empty() {
                        self.check_spelling(w, &sentence, 0);
                    }
                }
                self.validate_consonant_checks();
                self.run_grammar_check();
                // Word boundary work: prune, upgrade, drain grammar queue
                self.prune_resolved_errors();
                self.upgrade_spelling_suggestions();
                if !self.grammar_queue.is_empty() && self.completion_rx.is_none() {
                    self.process_grammar_queue();
                }
            }

            // Background completion: debounce + dispatch full completion + poll results
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
                    && self.completion_rx.is_none()
                    && self.last_context_change.elapsed() >= Duration::from_millis(300)
                {
                    if let Some(model_arc) = &self.model {
                        // Pre-fetch on main thread (fast, uses checker)
                        let matches: Vec<(u32, String)> = self.prefix_index.as_ref()
                            .and_then(|pi| pi.get(&prefix_lower))
                            .cloned()
                            .unwrap_or_default();
                        let mtag_candidates: Vec<String> = if matches.is_empty() && !prefix.is_empty() {
                            self.checker.as_ref().map_or(vec![], |c| c.prefix_lookup(&prefix_lower, 50))
                        } else {
                            vec![]
                        };
                        let nearby_words: std::collections::HashSet<String> = {
                            let before_mask = masked.split("<mask>").next().unwrap_or("");
                            // Only look at current sentence (after last sentence boundary)
                            // to avoid filtering words that appear in prior context sentences
                            let sent_start = before_mask.rfind(|c: char| ".!?".contains(c))
                                .map(|i| i + 1).unwrap_or(0);
                            let current_sent = &before_mask[sent_start..];
                            current_sent.split_whitespace()
                                .map(|w| w.trim_matches(|c: char| !c.is_alphanumeric()).to_lowercase())
                                .filter(|w| w.len() > 1)
                                .collect()
                        };
                        let wordfreq_clone = self.wordfreq.clone();
                        let model_clone = model_arc.clone();
                        let cancel = self.completion_cancel.clone();
                        cancel.store(false, std::sync::atomic::Ordering::Release);
                        let (tx, rx) = std::sync::mpsc::channel();
                        // Trim masked text to ~3 sentences around <mask> to keep forward fast
                        // (full doc context makes BERT 25x slower: 512 tokens vs ~30 tokens)
                        // Keep 3 sentence boundaries = current sentence + 2 previous for context
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
                                        if cuts >= 3 {
                                            start = i + 1;
                                            break;
                                        }
                                    }
                                }
                                before[start..].trim_start()
                            };
                            // Keep first sentence after mask
                            let trimmed_after = {
                                if let Some(pos) = after.find(|c: char| ".!?".contains(c)) {
                                    &after[..=pos]
                                } else {
                                    after
                                }
                            };
                            format!("{}<mask>{}", trimmed_before, trimmed_after)
                        };
                        let masked_clone = masked_trimmed;
                        let prefix_lower_clone = prefix_lower.clone();
                        let key_clone = cache_key.clone();
                        let capitalize = prefix.chars().next().map_or(false, |c| c.is_uppercase());

                        std::thread::spawn(move || {
                            let t_start = std::time::Instant::now();
                            let mut model = model_clone.lock().unwrap();
                            if cancel.load(std::sync::atomic::Ordering::Acquire) { return; }

                            // single_forward → logits (trimmed context for speed)
                            eprintln!("BERT context: {} chars", masked_clone.len());
                            let logits = match model.single_forward(&masked_clone) {
                                Ok((l, _)) => l,
                                Err(e) => { eprintln!("Background forward error: {}", e); return; }
                            };
                            if cancel.load(std::sync::atomic::Ordering::Acquire) { return; }

                            // Build left completions
                            let left = if matches.is_empty() && !prefix_lower_clone.is_empty() {
                                build_mtag_completions(&mut model, &masked_clone, &mtag_candidates, &logits, capitalize, &cancel)
                            } else if !prefix_lower_clone.is_empty() {
                                build_bpe_completions(&mut model, &masked_clone, &prefix_lower_clone, &matches, &logits, wordfreq_clone.as_deref(), &nearby_words, capitalize, &cancel)
                            } else {
                                vec![]
                            };
                            if cancel.load(std::sync::atomic::Ordering::Acquire) { return; }

                            // Build right completions
                            let left_words: std::collections::HashSet<String> = left.iter().map(|c| c.word.to_lowercase()).collect();
                            let right = build_right_completions(&model, &logits, wordfreq_clone.as_deref(), &nearby_words, &left_words);

                            let elapsed = t_start.elapsed().as_millis();
                            eprintln!("Background completion in {}ms: left=[{}] right=[{}]", elapsed,
                                left.iter().take(5).map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "),
                                right.iter().take(5).map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "));

                            if !cancel.load(std::sync::atomic::Ordering::Acquire) {
                                let _ = tx.send((key_clone, left, right));
                            }
                        });

                        self.completion_rx = Some(rx);
                        ctx.request_repaint_after(Duration::from_millis(50));
                    }
                } else if needs_completion && self.completion_rx.is_none() {
                    ctx.request_repaint_after(Duration::from_millis(50));
                }
            }

            // Poll background completion results
            if let Some(rx) = &self.completion_rx {
                match rx.try_recv() {
                    Err(std::sync::mpsc::TryRecvError::Disconnected) => {
                        // Background thread exited without sending (cancelled)
                        eprintln!("Background completion cancelled (sender dropped)");
                        self.completion_rx = None;
                    }
                    Err(std::sync::mpsc::TryRecvError::Empty) => {
                        // Still running — check again soon
                        ctx.request_repaint_after(Duration::from_millis(50));
                    }
                    Ok((key, left, right)) => {
                    eprintln!("Background completion received: {} left, {} right", left.len(), right.len());
                    // Apply grammar filter on main thread (checker isn't Send)
                    if self.grammar_completion {
                        if let Some(checker) = &mut self.checker {
                            let masked = self.context.masked_sentence.as_deref().unwrap_or("");
                            let before_mask = masked.split("<mask>").next().unwrap_or("");
                            let sent_start = before_mask.rfind(|c: char| ".!?".contains(c))
                                .map(|i| i + 1).unwrap_or(0);
                            let ctx_for_grammar = before_mask[sent_start..].trim().to_string();
                            let prefix = extract_prefix(&self.context.word);
                            // Pre-filter: remove words not in mtag dictionary (e.g. "sports")
                            // Grammar checker can't validate unknown words
                            let left_filtered: Vec<Completion> = left.into_iter()
                                .filter(|c| checker.has_word(&c.word.to_lowercase()))
                                .collect();
                            let right_filtered: Vec<Completion> = right.into_iter()
                                .filter(|c| checker.has_word(&c.word.to_lowercase()))
                                .collect();
                            let mut check_fn = |sentence: &str| -> GrammarCheckResult {
                                let errors = checker.check_sentence(sentence);
                                GrammarCheckResult {
                                    ok: errors.is_empty(),
                                    suggestions: errors.iter()
                                        .filter(|e| !e.suggestion.is_empty())
                                        .map(|e| e.suggestion.clone())
                                        .collect(),
                                }
                            };
                            self.completions = grammar_filter(&left_filtered, &ctx_for_grammar, prefix, &mut check_fn, 5);
                            self.open_completions = grammar_filter(&right_filtered, &ctx_for_grammar, "", &mut check_fn, 5);
                        } else {
                            self.completions = left.into_iter().take(5).collect();
                            self.open_completions = right.into_iter().take(5).collect();
                        }
                    } else {
                        self.completions = left.into_iter().take(5).collect();
                        self.open_completions = right.into_iter().take(5).collect();
                    }
                    self.last_completed_prefix = key;
                    self.completion_rx = None;
                    }
                }
            }
        }

        // Phase 1: Ctrl+Space while Word has focus → enter selection mode
        {
            use windows::Win32::UI::Input::KeyboardAndMouse::GetAsyncKeyState;
            let ctrl_down = unsafe { GetAsyncKeyState(0x11) } < 0;
            let space_down = unsafe { GetAsyncKeyState(0x20) } < 0;
            let both_held = ctrl_down && space_down;

            if both_held && !self.ctrl_space_held && !self.selection_mode
                && (!self.completions.is_empty() || !self.open_completions.is_empty())
            {
                self.ctrl_space_held = true;
                // Save Word's window handle before stealing focus
                use windows::Win32::UI::WindowsAndMessaging::GetForegroundWindow;
                let hwnd = unsafe { GetForegroundWindow() };
                self.word_hwnd = Some(hwnd.0 as isize);
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

        // Window sizing
        let has_content = !self.grammar_errors.is_empty() || !self.completions.is_empty() || !self.open_completions.is_empty();
        let recently_replaced = self.last_replace_time.elapsed() < Duration::from_secs(1);
        let win_h = if has_content || recently_replaced || self.selected_tab >= 1 { 250.0 } else { 110.0 };
        const WIN_W: f32 = 420.0;

        ctx.send_viewport_cmd(egui::ViewportCommand::Decorations(false));


        if self.follow_cursor {
            ctx.send_viewport_cmd(egui::ViewportCommand::InnerSize(egui::vec2(WIN_W, win_h)));
            if let Some((x, y)) = self.last_caret_pos {
                let (screen_w, screen_h) = get_screen_size();
                let pos_y = if (y as f32 + win_h) > screen_h {
                    y as f32 - win_h - 30.0
                } else {
                    y as f32
                };
                let pos_x = (x as f32).min(screen_w - WIN_W).max(0.0);

                ctx.send_viewport_cmd(egui::ViewportCommand::OuterPosition(
                    egui::pos2(pos_x, pos_y),
                ));
            }
        }

        // Request faster repaints while grammar queue is draining
        if self.grammar_scanning {
            ctx.request_repaint();
        } else {
            ctx.request_repaint_after(Duration::from_millis(100));
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
                    let loading: Vec<&str> = ["NorBERT4", "Whisper"]
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
                let tab_labels = ["Innhold", "Grammatikk", "Innst.", "Debug"];
                for (i, name) in tab_labels.iter().enumerate() {
                    // Draw colored dot for tabs 0 and 1
                    if i == 0 || i == 1 {
                        let dot_color = if i == 0 {
                            if has_completions { egui::Color32::from_rgb(0, 180, 60) }
                            else { egui::Color32::from_rgb(180, 180, 180) }
                        } else {
                            if has_grammar { egui::Color32::from_rgb(220, 50, 50) }
                            else { egui::Color32::from_rgb(0, 180, 60) }
                        };
                        let (dot_rect, _) = ui.allocate_exact_size(egui::vec2(10.0, 14.0), egui::Sense::hover());
                        let center = egui::pos2(dot_rect.min.x + 5.0, dot_rect.center().y);
                        ui.painter().circle_filled(center, 4.0, dot_color);
                    }

                    let is_selected = self.selected_tab == i;
                    let text = egui::RichText::new(*name).size(12.0);
                    let text = if is_selected {
                        text.strong().color(egui::Color32::from_rgb(0, 70, 160))
                    } else {
                        text.color(egui::Color32::from_rgb(100, 100, 100))
                    };
                    if ui.add(egui::Label::new(text).sense(egui::Sense::click())).clicked() {
                        self.selected_tab = i;
                    }
                    if i < tab_labels.len() - 1 {
                        ui.add_space(2.0);
                        ui.label(egui::RichText::new("|").size(12.0).color(egui::Color32::from_rgb(180, 170, 140)));
                        ui.add_space(2.0);
                    }
                }

                // TTS reading indicator + stop button
                if tts_speaking || ocr_is_busy {
                    ui.add_space(4.0);
                    ui.spinner();
                    if ui.add(egui::Button::new(
                        egui::RichText::new("■").size(12.0).color(egui::Color32::WHITE)
                    ).fill(egui::Color32::from_rgb(200, 40, 40))
                     .min_size(egui::vec2(18.0, 16.0))
                    ).clicked() {
                        tts::stop_speaking();
                        self.ocr_text = None;
                    }
                }

                // Microphone button / recording indicator
                let mic_recording = microphone::is_recording() || self.mic_transcribing;
                ui.add_space(4.0);
                if mic_recording {
                    ui.spinner();
                    let label = if self.mic_transcribing { "Transkriberer..." } else { "Lytter..." };
                    ui.label(egui::RichText::new(label).size(10.0).color(egui::Color32::from_rgb(200, 60, 60)));
                    if ui.add(egui::Button::new(
                        egui::RichText::new("■").size(12.0).color(egui::Color32::WHITE)
                    ).fill(egui::Color32::from_rgb(200, 40, 40))
                     .min_size(egui::vec2(18.0, 16.0))
                    ).clicked() {
                        if let Some(handle) = &self.mic_handle {
                            handle.stop();
                            self.mic_transcribing = true;
                        }
                    }
                } else {
                    let whisper_ready = self.whisper_engine.is_some();
                    let mic_btn = ui.add_enabled(whisper_ready, egui::Button::new(
                        egui::RichText::new("🎤").size(13.0)
                    ).min_size(egui::vec2(22.0, 16.0)));
                    if !whisper_ready {
                        mic_btn.on_hover_text("Whisper laster...");
                    } else if mic_btn.on_hover_text("Start talegjenkjenning").clicked() {
                        if let Some(engine) = &self.whisper_engine {
                            match microphone::start_recording(engine.clone()) {
                                Ok(handle) => {
                                    eprintln!("Microphone recording started");
                                    self.mic_handle = Some(handle);
                                    self.mic_result_text = None;
                                }
                                Err(e) => eprintln!("Microphone error: {}", e),
                            }
                        }
                    }
                }

                // Drag area for remaining space (leave room for close button)
                let remaining = ui.available_rect_before_wrap();
                let close_w = 20.0;
                let drag_rect = egui::Rect::from_min_max(
                    remaining.min,
                    egui::pos2(remaining.max.x - close_w, remaining.max.y),
                );
                let drag_resp = ui.allocate_rect(drag_rect, egui::Sense::drag());
                if drag_resp.drag_started() && !self.follow_cursor {
                    ctx.send_viewport_cmd(egui::ViewportCommand::StartDrag);
                }

                // Close button
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

            // === Tab: Innhold (0) ===
            if self.selected_tab == 0 {
                if !self.completions.is_empty() || !self.open_completions.is_empty() {
                    let header = if self.selection_mode {
                        "Forslag: (↑↓ velg, Enter godta, Esc avbryt)"
                    } else {
                        "Forslag: (Ctrl+Space for å velge)"
                    };
                    ui.label(
                        egui::RichText::new(header)
                            .size(11.0)
                            .color(egui::Color32::from_rgb(100, 100, 100)),
                    );
                    ui.add_space(2.0);

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
                        let col_w = (avail_w - 10.0) / 2.0;
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
                                ui.add_space(10.0);
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

                    // Copy button
                    ui.add_space(4.0);
                    if ui.small_button("Kopier").clicked() {
                        let mut text = String::new();
                        text.push_str(&format!("Ord: {}\n", self.context.word));
                        text.push_str("Venstre: ");
                        text.push_str(&self.completions.iter().map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "));
                        text.push_str("\nHøyre: ");
                        text.push_str(&self.open_completions.iter().map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "));
                        if let Some(masked) = &self.context.masked_sentence {
                            text.push_str(&format!("\nMaskert: {}", masked));
                        }
                        ctx.copy_text(text);
                    }

                    if let Some(word) = clicked_word {
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
                // Sort: SentenceBoundary first, then Grammar, then Spelling
                active_errors.sort_by_key(|&i| match self.writing_errors[i].category {
                    ErrorCategory::SentenceBoundary => 0,
                    ErrorCategory::Grammar => 1,
                    ErrorCategory::Spelling => 2,
                });

                if active_errors.is_empty() {
                    ui.label(
                        egui::RichText::new("Ingen feil funnet.")
                            .size(12.0)
                            .color(egui::Color32::from_rgb(0, 140, 60)),
                    );
                } else {
                    ui.label(
                        egui::RichText::new(format!("Mulige feil? ({})", active_errors.len()))
                            .size(12.0)
                            .strong()
                            .color(egui::Color32::from_rgb(80, 80, 80)),
                    );
                    ui.add_space(4.0);

                    let mut action: Option<(usize, &str)> = None;

                    // Group grammar errors by (sentence_context, doc_offset)
                    let mut shown_contexts: std::collections::HashSet<(String, usize)> = std::collections::HashSet::new();

                    egui::ScrollArea::vertical().max_height(ui.available_height() - 4.0).show(ui, |ui| {
                    for &idx in &active_errors {
                        let error = &self.writing_errors[idx];

                        // For grammar errors with position > 0, skip — they're shown as alternatives
                        if matches!(error.category, ErrorCategory::Grammar) && error.position > 0 {
                            if shown_contexts.contains(&(error.sentence_context.clone(), error.doc_offset)) {
                                continue;
                            }
                        }

                        ui.separator();
                        ui.scope(|ui| {
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
                    }
                    }); // end ScrollArea

                    // Handle actions after rendering
                    if let Some((idx, act)) = action {
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
                                    &word[..word.len().min(60)], &suggestion[..suggestion.len().min(60)]);
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
                                let existing = self.writing_errors[idx].suggestion.clone();
                                let mut suggestions = self.trigram_suggestions(&word, &sentence_ctx);
                                // Boost existing suggestion (compound/confirmed) to top
                                if !existing.is_empty() {
                                    let et = Self::trigrams(&existing.to_lowercase());
                                    let wt = Self::trigrams(&word.to_lowercase());
                                    let common = wt.iter().filter(|t| et.contains(t)).count();
                                    let total = wt.len().max(et.len()).max(1);
                                    let sim = common as f32 / total as f32;
                                    let boost_score = 2.0 + sim;
                                    if let Some(pos) = suggestions.iter().position(|(w, _)| w == &existing) {
                                        suggestions[pos].1 = suggestions[pos].1.max(boost_score);
                                    } else {
                                        suggestions.push((existing.clone(), boost_score));
                                    }
                                    suggestions.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
                                }
                                // Grammar-filter: substitute each candidate into the sentence and check
                                if let Some(checker) = &mut self.checker {
                                    suggestions.retain(|(candidate, _score)| {
                                        let test_sentence = sentence_ctx.to_lowercase()
                                            .replacen(&word.to_lowercase(), candidate, 1);
                                        let errors = checker.check_sentence(&test_sentence);
                                        errors.is_empty()
                                    });
                                }
                                self.suggestion_window = Some((word, suggestions));
                            }
                            "ignore" => {
                                let error = &self.writing_errors[idx];
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
                                    &error.sentence_context[..error.sentence_context.len().min(50)]);
                                self.manager.select_range(start, end);
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

            // === OCR: screenshot detected prompt ===
            let ocr_has_pending = self.ocr.as_ref().map_or(false, |o| o.has_pending_image());
            if ocr_has_pending && !ocr_is_busy {
                let mut do_ocr = false;
                let mut do_dismiss = false;
                egui::Window::new("Skjermbilde oppdaget")
                    .collapsible(false)
                    .resizable(false)
                    .default_width(300.0)
                    .anchor(egui::Align2::CENTER_CENTER, [0.0, 0.0])
                    .show(ctx, |ui| {
                        ui.label(
                            egui::RichText::new("Vil du lese teksten fra skjermbildet?")
                                .size(14.0)
                        );
                        ui.add_space(8.0);
                        ui.horizontal(|ui| {
                            if ui.button(egui::RichText::new("Ja, les teksten").size(13.0)).clicked() {
                                do_ocr = true;
                            }
                            if ui.button(egui::RichText::new("Nei").size(13.0)).clicked() {
                                do_dismiss = true;
                            }
                        });
                    });
                if do_ocr {
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

            // === Whisper transcription result window ===
            if self.mic_result_text.is_some() {
                let mut do_close = false;
                let text_clone = self.mic_result_text.clone().unwrap_or_default();
                egui::Window::new("Talegjenkjenning")
                    .collapsible(false)
                    .resizable(true)
                    .default_width(400.0)
                    .default_height(200.0)
                    .anchor(egui::Align2::CENTER_CENTER, [0.0, 0.0])
                    .show(ctx, |ui| {
                        egui::ScrollArea::vertical().max_height(300.0).show(ui, |ui| {
                            ui.label(
                                egui::RichText::new(&text_clone)
                                    .size(14.0)
                                    .color(egui::Color32::from_rgb(30, 30, 30)),
                            );
                        });
                        ui.add_space(8.0);
                        ui.horizontal(|ui| {
                            if ui.button(egui::RichText::new("Kopier").size(13.0)).clicked() {
                                ctx.copy_text(text_clone.clone());
                            }
                            if ui.button(egui::RichText::new("Lukk").size(13.0)).clicked() {
                                do_close = true;
                            }
                        });
                    });
                if do_close {
                    self.mic_result_text = None;
                }
            }

            // === Tab: Innstillinger (2) ===
            if self.selected_tab == 2 {
                ui.checkbox(&mut self.follow_cursor,
                    egui::RichText::new("Følg cursor").size(13.0)
                        .color(egui::Color32::from_rgb(60, 60, 55))
                );
                ui.add_space(4.0);
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
                // Load errors
                for err in &self.load_errors {
                    ui.label(
                        egui::RichText::new(err)
                            .size(10.0)
                            .color(egui::Color32::from_rgb(200, 50, 50)),
                    );
                }
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
    }
}

fn main() -> eframe::Result {
    // Set ORT_DYLIB_PATH if not already set
    if std::env::var("ORT_DYLIB_PATH").is_err() {
        // Try System32 first, then other known locations
        let candidates = [
            concat!(env!("CARGO_MANIFEST_DIR"), "/../../onnxruntime/onnxruntime-win-x64-1.23.0/lib/onnxruntime.dll"),
            "C:\\Windows\\System32\\onnxruntime.dll",
        ];
        for path in &candidates {
            if std::path::Path::new(path).exists() {
                unsafe { std::env::set_var("ORT_DYLIB_PATH", path); }
                break;
            }
        }
    }

    // Console spelling test mode — exercises exact same code as GUI
    if std::env::args().any(|a| a == "--test-spelling") {
        eprintln!("=== Spelling test mode ===");
        let mut app = ContextApp::new(true, true, 2);
        let mut pass = 0;
        let mut fail = 0;

        // --- Spelling suggestion tests (word NOT in dictionary) ---
        let spelling_tests: Vec<(&str, &str, &str)> = vec![
            ("fotbal", "jeg spiller og fotbal", "fotball"),
            ("blåsjell", "vi spiser blåsjell", "blåskjell"),
            ("spitlt", "Jeg hadde spitlt fotball.", "spilt"),
        ];
        for (word, sentence, expected) in &spelling_tests {
            app.last_spell_checked_word.clear();
            app.writing_errors.clear();
            app.pending_consonant_checks.clear();
            app.check_spelling(word, sentence, 0);
            app.validate_consonant_checks();
            app.upgrade_spelling_suggestions();
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
            app.check_spelling(word, sentence, 0);
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

    // Initialize Acapela TTS
    // Initialize Acapela TTS - look for SDK in user's Downloads
    if let Some(home) = std::env::var_os("USERPROFILE") {
        let sdk_dir = std::path::Path::new(&home)
            .join("Downloads/Sdk-Amul-Cogni-TTS-WIN_14-000_AIO");
        if sdk_dir.exists() {
            tts::init_tts(sdk_dir.to_str().unwrap(), "Kari22k_NV");
        } else {
            eprintln!("Acapela SDK not found at {:?}", sdk_dir);
        }
    }

    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([420.0, 250.0])
            .with_always_on_top()
            .with_decorations(false)
            .with_title("NorskTale")
            .with_close_button(false),  // prevent Alt+F4 and system close
        ..Default::default()
    };

    eframe::run_native(
        "NorskTale",
        options,
        Box::new(move |cc| {
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
                // Add Segoe UI Emoji as fallback for emoji glyphs (👍👎🔊 etc.)
                if let Ok(emoji_data) = std::fs::read("C:/Windows/Fonts/seguiemj.ttf") {
                    fonts.font_data.insert(
                        "SegoeEmoji".to_owned(),
                        egui::FontData::from_owned(emoji_data).into(),
                    );
                    fonts.families.get_mut(&egui::FontFamily::Proportional).unwrap()
                        .push("SegoeEmoji".to_owned());
                }
                cc.egui_ctx.set_fonts(fonts);
                eprintln!("Loaded Open Sans font");
            } else {
                eprintln!("Warning: Open Sans font not found at {}", font_path);
            }
            Ok(Box::new(ContextApp::new(grammar_completion, use_swipl, quality)))
        }),
    )
}
