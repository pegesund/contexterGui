mod bridge;
mod tts;

use bridge::{CursorContext, TextBridge};
use std::collections::HashMap;
use std::path::PathBuf;
use std::time::{Duration, Instant};

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

enum AnyChecker {
    Neo(GrammarChecker),
    Swi(SwiGrammarChecker),
}

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

}

// --- Error list for spelling and grammar ---

#[derive(Clone, Debug)]
enum ErrorCategory {
    Spelling,
    Grammar,
}

#[derive(Clone, Debug)]
struct WritingError {
    category: ErrorCategory,
    word: String,
    suggestion: String,
    explanation: String,
    rule_name: String,
    /// Approximate character offset in the sentence where the error occurred
    sentence_context: String,
    /// Position within the sentence (token index or char offset)
    position: usize,
    /// true if user clicked "Ignorer"
    ignored: bool,
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

        for bridge in &self.bridges {
            if bridge.is_available() {
                if let Some(ctx) = bridge.read_context() {
                    return Some(ctx);
                }
            }
        }
        None
    }

    fn active_bridge_name(&self) -> &str {
        for bridge in &self.bridges {
            if bridge.is_available() {
                return bridge.name();
            }
        }
        "none"
    }

    #[allow(dead_code)]
    fn replace_word(&self, new_text: &str) -> bool {
        for bridge in &self.bridges {
            if bridge.is_available() {
                return bridge.replace_word(new_text);
            }
        }
        false
    }

    fn find_and_replace(&self, find: &str, replace: &str) -> bool {
        for bridge in &self.bridges {
            if bridge.is_available() {
                return bridge.find_and_replace(find, replace);
            }
        }
        false
    }

    fn find_and_replace_in_context(&self, find: &str, replace: &str, context: &str) -> bool {
        for bridge in &self.bridges {
            if bridge.is_available() {
                return bridge.find_and_replace_in_context(find, replace, context);
            }
        }
        false
    }

    fn read_document_context(&self) -> Option<String> {
        for bridge in &self.bridges {
            if bridge.is_available() {
                return bridge.read_document_context();
            }
        }
        None
    }

    fn read_full_document(&self) -> Option<String> {
        for bridge in &self.bridges {
            if bridge.is_available() {
                return bridge.read_full_document();
            }
        }
        None
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

struct ContextApp {
    manager: BridgeManager,
    context: CursorContext,
    last_poll: Instant,
    poll_interval: Duration,
    follow_cursor: bool,
    last_caret_pos: Option<(i32, i32)>,
    // Grammar checker
    checker: Option<AnyChecker>,
    grammar_errors: Vec<GrammarError>,
    last_checked_sentence: String,
    // Word completer
    model: Option<Model>,
    prefix_index: Option<PrefixIndex>,
    baselines: Option<Baselines>,
    wordfreq: Option<HashMap<String, u64>>,
    embedding_store: Option<EmbeddingStore>,
    completions: Vec<Completion>,
    /// Open suggestions (any word) for fill-in-the-blank mode
    open_completions: Vec<Completion>,
    last_completed_prefix: String,
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
    /// Word's HWND to return focus to
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
    /// Sentences from last doc text (for diffing against current)
    last_doc_sentences: Vec<String>,
    /// Deferred find-and-replace (word, replacement, optional sentence context) — executed next frame
    pending_fix: Option<(String, String, String)>,
    /// Suggestion window: (misspelled_word, candidates)
    suggestion_window: Option<(String, Vec<(String, f32)>)>,
}

impl ContextApp {
    fn new(grammar_completion: bool, use_swipl: bool, quality: u8) -> Self {
        #[cfg(target_os = "windows")]
        unsafe {
            use windows::Win32::System::Com::*;
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok();
        }

        let mut load_errors = Vec::new();

        // Load grammar checker
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
                    // Fallback to neorusticus
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

        // Load NorBERT4 model + completer infrastructure
        let data = data_dir();
        let onnx_path = data.join("onnx/norbert4_base_int8.onnx");
        let tokenizer_path = data.join("onnx/tokenizer.json");
        let wordfreq_path = data.join("wordfreq.tsv");
        let minilm_onnx = data.join("minilm-onnx/model_optimized.onnx");
        let minilm_tok = data.join("minilm-onnx/tokenizer.json");
        let embed_cache = data.join("word_embeddings.bin");

        let (model, prefix_index, baselines, wf, embedding_store) =
            match Self::load_completer(
                &onnx_path, &tokenizer_path, &wordfreq_path,
                &minilm_onnx, &minilm_tok, &embed_cache,
            ) {
                Ok(parts) => parts,
                Err(e) => {
                    let msg = format!("Completer: {}", e);
                    eprintln!("{}", msg);
                    load_errors.push(msg);
                    (None, None, None, None, None)
                }
            };

        ContextApp {
            manager: BridgeManager::new(),
            context: CursorContext::default(),
            last_poll: Instant::now(),
            poll_interval: Duration::from_millis(300),
            follow_cursor: true,
            last_caret_pos: None,
            checker,
            grammar_errors: Vec::new(),
            last_checked_sentence: String::new(),
            model,
            prefix_index,
            baselines,
            wordfreq: wf,
            embedding_store,
            completions: Vec::new(),
            open_completions: Vec::new(),
            last_completed_prefix: String::new(),
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
            last_doc_sentences: Vec::new(),
            pending_fix: None,
            suggestion_window: None,
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
        Option<HashMap<String, u64>>,
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

        // Load MiniLM + embedding store
        let embedding_store = if minilm_onnx.exists() && minilm_tok.exists() {
            eprintln!("Loading MiniLM...");
            match nostos_cognio::embeddings::Embedder::load(
                minilm_onnx.to_str().unwrap(),
                minilm_tok.to_str().unwrap(),
            ) {
                Ok(embedder) => {
                    let store = EmbeddingStore::new(
                        embedder,
                        &wf,
                        if embed_cache.exists() { Some(embed_cache.as_path()) } else { None },
                    )?;
                    eprintln!("Embedding store ready.");
                    Some(store)
                }
                Err(e) => {
                    eprintln!("MiniLM load error: {}", e);
                    None
                }
            }
        } else {
            None
        };

        Ok((Some(model), Some(pi), Some(baselines), Some(wf), embedding_store))
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

    /// Check spelling of the just-finished word. Called when user types space/punctuation.
    fn check_spelling(&mut self, word: &str) {
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

        let checker = match &self.checker {
            Some(c) => c,
            None => return,
        };

        // Check if word exists in dictionary (with normert reading)
        let found = checker.has_word(&clean);
        eprintln!("  has_word('{}') = {}", clean, found);
        if found {
            return;
        }

        // Also check the original case (proper nouns)
        let original = word.trim();
        if checker.has_word(original) {
            return;
        }

        // Word not found — get fuzzy suggestions
        let fuzzy = checker.fuzzy_lookup(&clean, 2);

        let mut candidates: Vec<(String, u32)> = fuzzy.into_iter()
            .filter(|(w, _)| w != &clean) // exclude exact match
            .take(10)
            .collect();

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

        // Don't add duplicate errors for the same word
        if self.writing_errors.iter().any(|e| {
            matches!(e.category, ErrorCategory::Spelling) && e.word == clean && !e.ignored
        }) {
            return;
        }

        let error = WritingError {
            category: ErrorCategory::Spelling,
            word: clean.clone(),
            suggestion: best,
            explanation: format!("«{}» finnes ikke i ordboken.", clean),
            rule_name: "stavefeil".to_string(),
            sentence_context: self.context.sentence.clone(),
            position: 0,
            ignored: false,
        };
        self.writing_errors.push(error);
        eprintln!("Spelling: '{}' not found, suggesting '{}'",
            clean, self.writing_errors.last().unwrap().suggestion);
    }

    /// Find suggestion candidates using BERT semantic predictions + trigram ranking.
    /// Uses complete_word with empty prefix to get BPE-extended candidates (e.g. "håndball"),
    /// then re-ranks by trigram similarity to the misspelled word.
    fn trigram_suggestions(&mut self, word: &str, sentence_ctx: &str) -> Vec<(String, f32)> {
        let word_lower = word.to_lowercase();
        let word_trigrams = Self::trigrams(&word_lower);
        let word_first = word_lower.chars().next().unwrap_or(' ');

        // Step 1: Get semantic candidates via complete_word (handles BPE extension)
        let mut semantic_words: Vec<(String, f32)> = Vec::new();

        // Build masked context: everything before the misspelled word + <mask>
        let sentence_lower = sentence_ctx.to_lowercase();
        let masked_context = if let Some(pos) = sentence_lower.find(&word_lower) {
            let before = &sentence_ctx[..pos];
            format!("{}<mask>", before.trim_end())
        } else {
            // Fallback: use full sentence as context
            format!("{} <mask>", sentence_ctx)
        };
        eprintln!("Forslag: masked context = '{}'", masked_context);

        if let (Some(model), Some(pi), Some(bl)) = (&mut self.model, &self.prefix_index, &self.baselines) {
            // Use empty prefix to get all BERT predictions with BPE extension
            if let Ok(completions) = complete_word(
                model,
                &masked_context,
                "",  // empty prefix — get all candidates
                pi,
                Some(bl),
                self.wordfreq.as_ref(),
                None, // no dict filter on prefix — we want all words
                None,
                None, // no embedding re-ranking
                0.0, 0.0,
                50,  // get 50 candidates
                10,  // max BPE steps
            ) {
                for c in &completions {
                    let w = c.word.to_lowercase();
                    if w.len() < 2 || w == word_lower { continue; }
                    semantic_words.push((w, c.score));
                }
                eprintln!("Forslag: got {} BERT candidates", semantic_words.len());
            }
        }

        // Build BERT score lookup for semantic tiebreaking
        let bert_scores: HashMap<String, f32> = semantic_words.iter()
            .cloned()
            .collect();

        // Step 2: Score semantic candidates by trigram similarity to misspelled word
        // Only keep candidates with at least 1 common trigram (semantic + orthographic match)
        let mut scored: Vec<(String, f32)> = semantic_words.into_iter()
            .filter_map(|(w, bert_score)| {
                let w_trigrams = Self::trigrams(&w);
                let common = word_trigrams.iter()
                    .filter(|t| w_trigrams.contains(t))
                    .count();
                if common == 0 { return None; } // no orthographic overlap → skip
                let max_trigrams = word_trigrams.len().max(w_trigrams.len()).max(1);
                let trigram_score = common as f32 / max_trigrams as f32;
                // Combine: trigram similarity (0-1) * 2 + first letter bonus + BERT context bonus
                let mut score = trigram_score * 2.0;
                if w.chars().next().unwrap_or(' ') == word_first {
                    score += 0.5;
                }
                // BERT bonus: contextually relevant words get a boost
                score += (bert_score / 50.0).min(0.5);
                Some((w, score))
            })
            .collect();

        // Step 3: Also add fuzzy Levenshtein matches (distance 4) that share trigrams
        // Use BERT score as semantic tiebreaker for same-distance matches
        if let Some(checker) = &self.checker {
            let fuzzy = checker.fuzzy_lookup(&word_lower, 4);
            eprintln!("Forslag: fuzzy(4) returned {} matches", fuzzy.len());
            for (w, dist) in fuzzy {
                if w == word_lower { continue; }
                let w_trigrams = Self::trigrams(&w);
                let common = word_trigrams.iter()
                    .filter(|t| w_trigrams.contains(t))
                    .count();
                if common == 0 { continue; }
                if scored.iter().any(|(s, _)| s == &w) { continue; }
                let mut score = 1.0 - (dist as f32 * 0.15);
                if w.chars().next().unwrap_or(' ') == word_first {
                    score += 0.3;
                }
                // Semantic tiebreaker: if BERT also predicted this word, boost it
                if let Some(&bs) = bert_scores.get(&w) {
                    score += (bs / 50.0).min(0.5);
                }
                scored.push((w, score));
            }
        }

        // Step 4: Wordfreq trigram search — finds loanwords like "volleyball"
        // that may be missing from mtag but exist in common word lists
        if let Some(wf) = &self.wordfreq {
            for (w, _freq) in wf.iter() {
                let wl = w.to_lowercase();
                if wl == word_lower { continue; }
                if scored.iter().any(|(s, _)| *s == wl) { continue; }
                // Quick filter: must share first letter
                if wl.chars().next().unwrap_or(' ') != word_first { continue; }
                let w_trigrams = Self::trigrams(&wl);
                let common = word_trigrams.iter()
                    .filter(|t| w_trigrams.contains(t))
                    .count();
                if common < 2 { continue; }
                let max_trigrams = word_trigrams.len().max(w_trigrams.len()).max(1);
                let trigram_score = common as f32 / max_trigrams as f32;
                let mut score = trigram_score * 2.0;
                score += 0.3; // first letter always matches (filtered above)
                if let Some(&bs) = bert_scores.get(&wl) {
                    score += (bs / 50.0).min(0.5);
                }
                scored.push((wl, score));
            }
        }

        eprintln!("Forslag: top results for '{}': {:?}", word_lower,
            scored.iter().take(10).collect::<Vec<_>>());

        scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
        scored.dedup_by(|a, b| a.0 == b.0);
        scored.truncate(10);
        scored
    }

    /// Score a sentence using BERT pseudo-log-likelihood.
    /// For each word position, mask it and check how well BERT predicts the actual word.
    fn bert_sentence_score(&mut self, sentence: &str) -> f32 {
        let words: Vec<&str> = sentence.split_whitespace().collect();
        if words.is_empty() { return f32::NEG_INFINITY; }

        let model = match &mut self.model {
            Some(m) => m,
            None => return 0.0,
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
    fn best_sentence_corrections(&mut self, sentence: &str, errors: &[GrammarError]) -> Vec<(String, String, f32)> {
        let mut candidates: Vec<(String, String)> = Vec::new(); // (corrected_sentence, explanation)

        // 1. Apply each individual grammar suggestion
        for e in errors {
            if !e.suggestion.is_empty() {
                let fixed = replace_word_at_position(sentence, &e.word, &e.suggestion);
                let expl = format!("«{}» → «{}»: {}", e.word, e.suggestion, e.explanation);
                candidates.push((fixed, expl));
            }
        }

        // 2. Try removing each error word — only if no substitution suggestion exists
        let has_substitution = errors.iter().any(|e| !e.suggestion.is_empty());
        if !has_substitution {
            for e in errors {
                // Try removing the error word itself
                let removed = remove_word_from_sentence(sentence, &e.word);
                if removed != sentence {
                    candidates.push((removed, format!("Fjernet «{}».", e.word)));
                }
                // For "å + noun" errors: try removing "å" before the word instead
                let words: Vec<&str> = sentence.split_whitespace().collect();
                if let Some(pos) = words.iter().position(|w| {
                    w.trim_matches(|c: char| c.is_ascii_punctuation()).eq_ignore_ascii_case(&e.word)
                }) {
                    if pos > 0 {
                        let prev = words[pos - 1].trim_matches(|c: char| c.is_ascii_punctuation());
                        if prev == "å" {
                            let removed_aa = remove_word_from_sentence(sentence, "å");
                            if removed_aa != sentence {
                                candidates.push((removed_aa, format!("Fjernet «å» foran «{}».", e.word)));
                            }
                        }
                    }
                }
            }
        }

        // 3. Apply all suggestions together
        if errors.len() > 1 {
            let mut all_fixed = sentence.to_string();
            let mut all_expl = Vec::new();
            for e in errors {
                if !e.suggestion.is_empty() {
                    all_fixed = replace_word_at_position(&all_fixed, &e.word, &e.suggestion);
                    all_expl.push(format!("«{}» → «{}»", e.word, e.suggestion));
                }
            }
            if all_fixed != sentence {
                candidates.push((all_fixed, all_expl.join(", ")));
            }
        }

        // Deduplicate by corrected sentence
        candidates.dedup_by(|a, b| a.0 == b.0);
        // Also remove any remaining duplicates (not just adjacent)
        {
            let mut seen = std::collections::HashSet::new();
            candidates.retain(|(c, _)| seen.insert(c.clone()));
        }

        if candidates.is_empty() {
            return Vec::new();
        }

        // Grammar-check each candidate: discard ones that still have errors
        let valid_candidates: Vec<(String, String)> = if let Some(checker) = &mut self.checker {
            candidates.into_iter()
                .filter(|(c, _)| {
                    let errs = checker.check_sentence(c);
                    let ok = errs.is_empty();
                    if !ok {
                        eprintln!("    REJECTED (grammar): '{}' — {} errors", c, errs.len());
                    }
                    ok
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
        let mut scored: Vec<(String, String, f32)> = valid_candidates.into_iter()
            .map(|(c, e)| {
                let score = self.bert_sentence_score(&c);
                eprintln!("    {:.1}: '{}'", score, c);
                (c, e, score)
            })
            .collect();

        // Sort: substitutions first (by BERT score), then removals (by BERT score)
        scored.sort_by(|a, b| {
            let a_removal = a.1.starts_with("Fjernet");
            let b_removal = b.1.starts_with("Fjernet");
            match (a_removal, b_removal) {
                (false, true) => std::cmp::Ordering::Less,   // substitution before removal
                (true, false) => std::cmp::Ordering::Greater,
                _ => b.2.partial_cmp(&a.2).unwrap_or(std::cmp::Ordering::Equal),
            }
        });
        scored.truncate(3);
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
        // Get full document text to check all sentences
        let doc_text = self.manager.read_full_document()
            .unwrap_or_default().to_lowercase();
        self.writing_errors.retain(|e| {
            if e.ignored {
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
            };
            if !still_present {
                eprintln!("Error resolved: '{}' no longer in document", e.word);
            }
            still_present
        });
    }

    /// Update the error list with grammar errors from the current sentence.
    /// Called when a sentence boundary is detected.
    fn update_grammar_errors(&mut self) {
        // Read document text and check all complete sentences
        let doc_text = match self.manager.read_full_document() {
            Some(t) => t,
            None => return,
        };

        let sentences = split_sentences(&doc_text);
        if sentences.is_empty() {
            return;
        }

        // Build set of current trimmed sentences
        let current: Vec<String> = sentences.iter()
            .map(|s| s.trim().to_string())
            .filter(|s| !s.is_empty())
            .collect();

        // Find sentences that are NEW (not in previous doc)
        let prev_set: std::collections::HashSet<&str> = self.last_doc_sentences.iter()
            .map(|s| s.as_str()).collect();
        let new_sentences: Vec<String> = current.iter()
            .filter(|s| !prev_set.contains(s.as_str()))
            .cloned()
            .collect();

        // Update stored sentences for next poll
        self.last_doc_sentences = current;

        if new_sentences.is_empty() {
            return;
        }

        for trimmed in &new_sentences {
            eprintln!("Grammar check: '{}'", trimmed);

            let checker = match &mut self.checker {
                Some(c) => c,
                None => return,
            };

            let errors = checker.check_sentence(trimmed);
            if errors.is_empty() { continue; }

            // Don't add duplicate grammar errors for this sentence
            if self.writing_errors.iter().any(|e| {
                matches!(e.category, ErrorCategory::Grammar)
                    && e.sentence_context == *trimmed
                    && !e.ignored
            }) {
                continue;
            }

            for ge in &errors {
                eprintln!("  Grammar error: '{}' → '{}' ({})", ge.word, ge.suggestion, ge.rule_name);
            }

            // Score candidates with BERT (only runs when Prolog found errors)
            let corrections = self.best_sentence_corrections(trimmed, &errors);

            for (i, (corrected, explanation, score)) in corrections.iter().enumerate() {
                eprintln!("  Correction #{}: ({:.1}) '{}' → '{}'", i+1, score, trimmed, corrected);
                self.writing_errors.push(WritingError {
                    category: ErrorCategory::Grammar,
                    word: trimmed.to_string(),
                    suggestion: corrected.clone(),
                    explanation: explanation.clone(),
                    rule_name: "sentence_correction".to_string(),
                    sentence_context: trimmed.to_string(),
                    position: i,
                    ignored: false,
                });
            }
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

        // Fill-in-the-blank using full sentence context
        // Works with prefix (typed letters) or without (just pressed space)
        if let Some(masked) = &self.context.masked_sentence.clone() {
            if let (Some(model), Some(pi)) = (&mut self.model, &self.prefix_index) {
                let t_bert = Instant::now();
                let prefix_lower = prefix.to_lowercase();
                // Collect nearby words before <mask> to filter repetition
                let nearby_words: std::collections::HashSet<String> = {
                    let before_mask = masked.split("<mask>").next().unwrap_or("");
                    before_mask.split_whitespace()
                        .rev().take(5)
                        .map(|w| w.trim_matches(|c: char| !c.is_alphanumeric()).to_lowercase())
                        .filter(|w| w.len() > 1)
                        .collect()
                };

                // Cache the initial forward pass: masked sentence is identical across keystrokes
                // (only the prefix changes), so we reuse logits when masked text hasn't changed.
                let forward_result = if let Some((ref cached_masked, ref cached_logits)) = self.cached_forward {
                    if cached_masked == masked {
                        eprintln!("fill-blank: reusing cached forward pass");
                        Ok((cached_logits.clone(), 0.0))
                    } else {
                        model.single_forward(masked)
                    }
                } else {
                    model.single_forward(masked)
                };
                match forward_result {
                    Ok((logits, _ms)) => {
                        self.cached_forward = Some((masked.clone(), logits.clone()));
                        let checker_ref = self.checker.as_ref();
                        let wf_ref = self.wordfreq.as_ref();
                        let is_valid = |w: &str| -> bool {
                            let key = w.to_lowercase();
                            if nearby_words.contains(&key) { return false; }
                            if let Some(c) = checker_ref { c.has_word(&key) }
                            else { wf_ref.map_or(true, |wf| wf.contains_key(&key)) }
                        };

                        // Left list: first-letter matches, expanded via BPE extension
                        // Skip when no prefix (space-only trigger shows right column only)
                        let matches: Vec<(u32, String)> = if prefix.is_empty() {
                            Vec::new()
                        } else {
                            pi.get(&prefix_lower).cloned().unwrap_or_default()
                        };

                        // If BPE has no matches, try mtag dictionary prefix search
                        // Score candidates using BERT logits for contextual ranking
                        if matches.is_empty() && !prefix.is_empty() {
                            if let Some(checker) = &self.checker {
                                let mtag_hits = checker.prefix_lookup(&prefix_lower, 50);
                                if !mtag_hits.is_empty() {
                                    let capitalize = prefix.chars().next().map_or(false, |c| c.is_uppercase());
                                    let cap = |s: &str| -> String {
                                        let mut c = s.chars();
                                        match c.next() {
                                            None => String::new(),
                                            Some(f) => f.to_uppercase().to_string() + c.as_str(),
                                        }
                                    };
                                    // Score candidates using BERT batched forward
                                    let mask_parts: Vec<&str> = masked.splitn(2, "<mask>").collect();
                                    let ctx_before_m = mask_parts[0].trim_end();
                                    let ctx_after_m = mask_parts.get(1).map(|s| s.trim_start()).unwrap_or(".");

                                    // Tokenize all candidates
                                    let candidates_with_tokens: Vec<(String, Vec<u32>)> = mtag_hits.into_iter()
                                        .filter_map(|w| {
                                            let enc = model.tokenizer.encode(format!(" {}", w).as_str(), false).ok()?;
                                            let ids: Vec<u32> = enc.get_ids().to_vec();
                                            if ids.is_empty() { return None; }
                                            Some((w, ids))
                                        })
                                        .collect();

                                    // First-token score from existing mask logits
                                    let mut scores: Vec<f32> = candidates_with_tokens.iter()
                                        .map(|(_, ids)| logits[ids[0] as usize])
                                        .collect();

                                    // Batched extension with dedup: candidates sharing the same
                                    // token prefix produce identical masked texts — run once, reuse logits
                                    let max_tokens = candidates_with_tokens.iter().map(|(_, ids)| ids.len()).max().unwrap_or(1);
                                    for t in 1..max_tokens {
                                        let to_score: Vec<usize> = candidates_with_tokens.iter().enumerate()
                                            .filter(|(_, (_, ids))| ids.len() > t)
                                            .map(|(i, _)| i)
                                            .collect();
                                        if to_score.is_empty() { break; }

                                        // Group by token prefix (ids[..t]) to deduplicate
                                        let mut unique_prefixes: Vec<Vec<u32>> = Vec::new();
                                        let mut prefix_to_idx: std::collections::HashMap<Vec<u32>, usize> = std::collections::HashMap::new();
                                        let mut candidate_to_prefix: Vec<usize> = Vec::new(); // maps to_score index → unique prefix index

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

                                        // One forward pass per unique prefix
                                        let batch_texts: Vec<String> = unique_prefixes.iter()
                                            .map(|ids| {
                                                let partial = model.tokenizer.decode(ids, false).unwrap_or_default();
                                                format!("{} {}<mask> {}", ctx_before_m, partial.trim(), ctx_after_m)
                                            })
                                            .collect();

                                        if let Ok((batch_logits, _)) = model.batched_forward(&batch_texts) {
                                            for (k, &i) in to_score.iter().enumerate() {
                                                let pidx = candidate_to_prefix[k];
                                                scores[i] += batch_logits[pidx][candidates_with_tokens[i].1[t] as usize];
                                            }
                                        }
                                    }

                                    // Average per token so long words aren't penalized
                                    let mut scored: Vec<(String, f32)> = candidates_with_tokens.iter().enumerate()
                                        .map(|(i, (w, ids))| (w.clone(), scores[i] / ids.len() as f32))
                                        .collect();
                                    scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
                                    let left: Vec<Completion> = scored.into_iter()
                                        .take(25)
                                        .map(|(w, s)| Completion {
                                            word: if capitalize { cap(&w) } else { w },
                                            score: s,
                                            elapsed_ms: 0.0,
                                        })
                                        .collect();
                                    let bert_ms = t_bert.elapsed().as_millis();
                                    // Right list
                                    let mut all_scored: Vec<(String, f32)> = model.id_to_token.iter()
                                        .enumerate()
                                        .filter(|(_, tok)| tok.starts_with('Ġ'))
                                        .map(|(i, _)| {
                                            let decoded = model.tokenizer
                                                .decode(&[i as u32], false)
                                                .unwrap_or_default().trim().to_string();
                                            (decoded, logits[i])
                                        })
                                        .filter(|(w, _)| !w.is_empty() && w.len() > 1 && is_valid(w))
                                        .collect();
                                    all_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
                                    let right: Vec<Completion> = all_scored.iter()
                                        .take(10)
                                        .map(|(w, s)| Completion { word: w.clone(), score: *s, elapsed_ms: 0.0 })
                                        .collect();
                                    eprintln!("mtag fallback (BERT-ranked): left=[{}] bert={}ms",
                                        left.iter().take(10).map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "),
                                        bert_ms);

                                    // Grammar filter
                                    if self.grammar_completion {
                                        if let Some(checker) = &mut self.checker {
                                            let ctx_for_grammar = masked.replace("<mask>", "").trim().to_string();
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
                                            self.completions = grammar_filter(&left, &ctx_for_grammar, prefix, &mut check_fn, 5);
                                            self.open_completions = grammar_filter(&right, &ctx_for_grammar, "", &mut check_fn, 5);
                                        } else {
                                            self.completions = left.into_iter().take(5).collect();
                                            self.open_completions = right.into_iter().take(5).collect();
                                        }
                                    } else {
                                        self.completions = left.into_iter().take(5).collect();
                                        self.open_completions = right.into_iter().take(5).collect();
                                    }
                                    return;
                                }
                            }
                        }
                        let mut token_scored: Vec<(String, f32)> = matches.iter()
                            .map(|(tid, word)| (word.clone(), logits[*tid as usize]))
                            .collect();
                        token_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

                        // Iterative BPE extension (like Python cognio_demo)
                        // Diverse candidates: top 20 by logit + top 20 long tokens (≥5 chars)
                        // Short tokens dominate logits but long tokens carry semantic meaning
                        let mut left_scored: Vec<(String, f32)> = Vec::new();
                        let mut seen_words: std::collections::HashSet<String> = std::collections::HashSet::new();
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
                        // Top 20 long tokens (≥5 chars) — these are word stems like "menneske"
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

                        // Extract context parts from masked sentence
                        let mask_parts: Vec<&str> = masked.splitn(2, "<mask>").collect();
                        let ctx_before = mask_parts[0].trim_end();
                        let ctx_after = mask_parts.get(1).map(|s| s.trim_start()).unwrap_or(".");

                        // Adaptive max_steps: longer prefixes need fewer BPE extensions
                        // since BPE tokens already cover most of the word.
                        // 1-3 chars → 1 step (needed for compounds like for→forut→forutsi)
                        // 4+ → 0 steps
                        let max_steps = if prefix_lower.len() <= 3 { 1 } else { 0 };
                        for _step in 0..max_steps {
                            let best_score = candidates.iter()
                                .filter(|c| !c.done)
                                .map(|c| c.score)
                                .fold(f32::NEG_INFINITY, f32::max);
                            let threshold = best_score - 15.0;
                            let mut to_extend: Vec<usize> = candidates.iter().enumerate()
                                .filter(|(_, c)| !c.done && c.score >= threshold)
                                .map(|(i, _)| i)
                                .collect();
                            // Mark low-scoring candidates as done
                            for c in candidates.iter_mut() {
                                if !c.done && c.score < threshold {
                                    c.done = true;
                                }
                            }
                            // Cap batch size: 5 for short prefixes (speed), 10 otherwise
                            let batch_cap = if prefix_lower.len() <= 2 { 5 } else { 10 };
                            to_extend.truncate(batch_cap);
                            if to_extend.is_empty() { break; }

                            // Build batch: "{ctx} {accumulated}<mask> {after}"
                            // NO space before <mask> → forces continuation token prediction
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

                                        // Continuation token = no Ġ prefix and not punctuation
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

                        // Collect extended words into left_scored
                        for c in &candidates {
                            let key = c.word.to_lowercase();
                            if is_valid(&key) && seen_words.insert(key.clone()) {
                                left_scored.push((key, c.score));
                            }
                        }
                        eprintln!("BPE candidates: [{}]",
                            candidates.iter().take(10).map(|c| format!("{}({:.1},done={})", c.word, c.score, c.done))
                            .collect::<Vec<_>>().join(", "));

                        // Supplement BPE with mtag prefix search + BERT batch scoring.
                        // Cached per masked_sentence — subsequent keystrokes just filter by prefix.
                        if !prefix.is_empty() && prefix_lower.len() >= 3 {
                            let same_masked = self.cached_mtag_supplement.as_ref()
                                .map_or(false, |(m, _)| m == masked);

                            // Collect new words not already in cache
                            let already_cached: std::collections::HashSet<String> = if same_masked {
                                self.cached_mtag_supplement.as_ref()
                                    .map_or_else(std::collections::HashSet::new, |(_, v)| v.iter().map(|(w, _)| w.clone()).collect())
                            } else {
                                std::collections::HashSet::new()
                            };

                            {
                                if let Some(checker) = &self.checker {
                                    let mut mtag_set: std::collections::HashSet<String> = std::collections::HashSet::new();
                                    // Search from extended BPE words and sub-prefixes
                                    for c in &candidates {
                                        let key = c.word.to_lowercase();
                                        if key.len() > prefix_lower.len() {
                                            for w in checker.prefix_lookup(&key, 10) {
                                                if !already_cached.contains(&w) {
                                                    mtag_set.insert(w);
                                                }
                                            }
                                            let chars: Vec<char> = key.chars().collect();
                                            let plen = prefix_lower.chars().count();
                                            for end in (plen + 1)..chars.len() {
                                                let sub: String = chars[..end].iter().collect();
                                                for w in checker.prefix_lookup(&sub, 5) {
                                                    if !already_cached.contains(&w) {
                                                        mtag_set.insert(w);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    for w in checker.prefix_lookup(&prefix_lower, 10) {
                                        if !already_cached.contains(&w) {
                                            mtag_set.insert(w);
                                        }
                                    }

                                    if !mtag_set.is_empty() {
                                        // Phase 1: first-token score + frequency boost
                                        let mut all_cands: Vec<(String, Vec<u32>, f32)> = mtag_set.into_iter()
                                            .filter_map(|w| {
                                                let enc = model.tokenizer.encode(format!(" {}", w).as_str(), false).ok()?;
                                                let ids: Vec<u32> = enc.get_ids().to_vec();
                                                if ids.is_empty() { return None; }
                                                let mut score = logits[ids[0] as usize];
                                                // Frequency boost: common words get +2, unknown get -2
                                                if let Some(wf) = wf_ref {
                                                    if wf.contains_key(&w) {
                                                        score += 2.0;
                                                    } else {
                                                        score -= 2.0;
                                                    }
                                                }
                                                Some((w, ids, score))
                                            })
                                            .collect();


                                        all_cands.sort_by(|a, b| b.2.partial_cmp(&a.2).unwrap_or(std::cmp::Ordering::Equal));
                                        all_cands.truncate(10);

                                        let cands: Vec<(String, Vec<u32>)> = all_cands.iter()
                                            .map(|(w, ids, _)| (w.clone(), ids.clone())).collect();
                                        let mut scores: Vec<f32> = all_cands.iter().map(|(_, _, s)| *s).collect();

                                        // Phase 2: multi-token batch scoring with dedup
                                        let max_tokens = cands.iter().map(|(_, ids)| ids.len()).max().unwrap_or(1).min(3);
                                        for t in 1..max_tokens {
                                            let to_score: Vec<usize> = cands.iter().enumerate()
                                                .filter(|(_, (_, ids))| ids.len() > t)
                                                .map(|(i, _)| i).collect();
                                            if to_score.is_empty() { break; }

                                            let mut unique_prefixes: Vec<Vec<u32>> = Vec::new();
                                            let mut prefix_map: std::collections::HashMap<Vec<u32>, usize> = std::collections::HashMap::new();
                                            let mut cand_to_prefix: Vec<usize> = Vec::new();
                                            for &i in &to_score {
                                                let tp = cands[i].1[..t].to_vec();
                                                let pidx = *prefix_map.entry(tp.clone()).or_insert_with(|| {
                                                    let idx = unique_prefixes.len();
                                                    unique_prefixes.push(tp);
                                                    idx
                                                });
                                                cand_to_prefix.push(pidx);
                                            }

                                            let batch_texts: Vec<String> = unique_prefixes.iter()
                                                .map(|ids| {
                                                    let partial = model.tokenizer.decode(ids, false).unwrap_or_default();
                                                    format!("{} {}<mask> {}", ctx_before, partial.trim(), ctx_after)
                                                }).collect();

                                            if let Ok((batch_logits, _)) = model.batched_forward(&batch_texts) {
                                                for (k, &i) in to_score.iter().enumerate() {
                                                    let pidx = cand_to_prefix[k];
                                                    scores[i] += batch_logits[pidx][cands[i].1[t] as usize];
                                                }
                                            }
                                        }

                                        let mtag_scored: Vec<(String, f32)> = cands.iter().enumerate()
                                            .map(|(i, (w, ids))| (w.clone(), scores[i] / ids.len() as f32))
                                            .collect();
                                        // Accumulate into cache
                                        if same_masked {
                                            if let Some((_, ref mut existing)) = self.cached_mtag_supplement {
                                                eprintln!("mtag supplement: +{} new (total {})", mtag_scored.len(), existing.len() + mtag_scored.len());
                                                existing.extend(mtag_scored);
                                            }
                                        } else {
                                            eprintln!("mtag supplement (BERT-ranked, {} cands): computed and cached", mtag_scored.len());
                                            self.cached_mtag_supplement = Some((masked.clone(), mtag_scored));
                                        }
                                    }
                                }
                            }

                            // Merge cached results filtered by current prefix
                            if let Some((_, ref scored)) = self.cached_mtag_supplement {
                                let filtered: Vec<(String, f32)> = scored.iter()
                                    .filter(|(w, _)| w.starts_with(&prefix_lower) && !seen_words.contains(w.as_str()))
                                    .take(10)
                                    .cloned()
                                    .collect();
                                for (w, score) in filtered {
                                    seen_words.insert(w.clone());
                                    left_scored.push((w, score));
                                }
                            }
                        }

                        left_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
                        let left: Vec<Completion> = left_scored.iter()
                            .take(25)
                            .map(|(w, s)| Completion { word: w.clone(), score: *s, elapsed_ms: 0.0 })
                            .collect();

                        // Right list: cached since it only depends on masked sentence, not prefix
                        let right_scored = if let Some((ref cached_m, ref cached_right)) = self.cached_right_column {
                            if cached_m == masked {
                                cached_right.clone()
                            } else {
                                let mut all_scored: Vec<(String, f32)> = model.id_to_token.iter()
                                    .enumerate()
                                    .filter(|(_, tok)| tok.starts_with('Ġ'))
                                    .map(|(i, _)| {
                                        let decoded = model.tokenizer
                                            .decode(&[i as u32], false)
                                            .unwrap_or_default().trim().to_string();
                                        (decoded, logits[i])
                                    })
                                    .filter(|(w, _)| !w.is_empty() && w.len() > 1 && is_valid(w))
                                    .collect();
                                all_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
                                all_scored.truncate(20);
                                self.cached_right_column = Some((masked.clone(), all_scored.clone()));
                                all_scored
                            }
                        } else {
                            let mut all_scored: Vec<(String, f32)> = model.id_to_token.iter()
                                .enumerate()
                                .filter(|(_, tok)| tok.starts_with('Ġ'))
                                .map(|(i, _)| {
                                    let decoded = model.tokenizer
                                        .decode(&[i as u32], false)
                                        .unwrap_or_default().trim().to_string();
                                    (decoded, logits[i])
                                })
                                .filter(|(w, _)| !w.is_empty() && w.len() > 1 && is_valid(w))
                                .collect();
                            all_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
                            all_scored.truncate(20);
                            self.cached_right_column = Some((masked.clone(), all_scored.clone()));
                            all_scored
                        };
                        let left_words: std::collections::HashSet<&str> = left.iter().map(|c| c.word.as_str()).collect();
                        let right: Vec<Completion> = right_scored.iter()
                            .filter(|(w, _)| !left_words.contains(w.as_str()))
                            .take(10)
                            .map(|(w, s)| Completion { word: w.clone(), score: *s, elapsed_ms: 0.0 })
                            .collect();

                        let bert_ms = t_bert.elapsed().as_millis();

                        // Grammar filter both lists
                        if self.grammar_completion {
                            if let Some(checker) = &mut self.checker {
                                // Build context: sentence fragment before <mask>
                                let before_mask = masked.split("<mask>").next().unwrap_or("");
                                let sent_start = before_mask.rfind(|c: char| ".!?".contains(c))
                                    .map(|i| i + 1).unwrap_or(0);
                                let ctx_for_grammar = before_mask[sent_start..].trim().to_string();
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
                                self.completions = grammar_filter(&left, &ctx_for_grammar, prefix, &mut check_fn, 5);
                                self.open_completions = grammar_filter(&right, &ctx_for_grammar, "", &mut check_fn, 5);
                            } else {
                                self.completions = left.into_iter().take(5).collect();
                                self.open_completions = right.into_iter().take(5).collect();
                            }
                        } else {
                            self.completions = left.into_iter().take(5).collect();
                            self.open_completions = right.into_iter().take(5).collect();
                        }

                        let total_ms = t_total.elapsed().as_millis();
                        eprintln!("fill-blank bert={}ms total={}ms left=[{}] right=[{}]",
                            bert_ms, total_ms,
                            self.completions.iter().map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "),
                            self.open_completions.iter().map(|c| format!("{}({:.1})", c.word, c.score)).collect::<Vec<_>>().join(", "));
                    }
                    Err(e) => eprintln!("Fill-blank error: {}", e),
                }
            }
            return;
        }

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

            if let (Some(model), Some(pi)) = (&mut self.model, &self.prefix_index) {
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
                    model,
                    ctx.as_str(),
                    prefix,
                    pi,
                    self.baselines.as_ref(),
                    self.wordfreq.as_ref(),
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

    fn return_focus_to_word(&self) {
        if let Some(hwnd_val) = self.word_hwnd {
            use windows::Win32::Foundation::HWND;
            use windows::Win32::UI::WindowsAndMessaging::SetForegroundWindow;
            unsafe {
                let hwnd = HWND(hwnd_val as *mut _);
                let _ = SetForegroundWindow(hwnd);
            }
        }
    }
}

/// Split text into sentences for embedding.
/// Replace first occurrence of a word (whole word match) in a sentence.
fn replace_word_at_position(sentence: &str, word: &str, replacement: &str) -> String {
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


fn split_sentences(text: &str) -> Vec<String> {
    let mut sentences = Vec::new();
    let mut current = String::new();
    for c in text.chars() {
        current.push(c);
        if c == '.' || c == '!' || c == '?' {
            let trimmed = current.trim().to_string();
            if !trimmed.is_empty() && trimmed.len() > 5 {
                sentences.push(trimmed);
            }
            current.clear();
        }
    }
    // Don't embed the trailing incomplete sentence — only complete sentences
    // ending with .!? produce stable embeddings across sync cycles.
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
        if let Some((find, replace, context)) = self.pending_fix.take() {
            if context.is_empty() {
                self.manager.find_and_replace(&find, &replace);
            } else {
                self.manager.find_and_replace_in_context(&find, &replace, &context);
            }
        }

        // Poll for new context
        if self.last_poll.elapsed() >= self.poll_interval {
            self.last_poll = Instant::now();
            if let Some(new_ctx) = self.manager.read_context() {
                if new_ctx.caret_pos.is_some() {
                    self.last_caret_pos = new_ctx.caret_pos;
                }
                self.context = new_ctx;
                // Remove errors for words that have been manually corrected
                self.prune_resolved_errors();
            }

            // Sync document sentences for topic-aware completion
            self.sync_embeddings();

            // Always run grammar check (not just at word boundaries)
            self.update_grammar_errors();

            let mid = is_mid_word(&self.context.word);
            if mid {
                // Mid-word: mark prefix change for debouncing
                let prefix = extract_prefix(&self.context.word);
                if prefix != self.last_completed_prefix {
                    self.last_prefix_change = Instant::now();
                    self.pending_completion = true;
                    if !self.selection_mode {
                        self.selected_completion = None;
                    }
                }
            } else if self.context.masked_sentence.is_some() {
                // No prefix but have context (e.g. after space): suggest next word
                let cache_key = format!("__noprefix__{}", self.context.masked_sentence.as_deref().unwrap_or(""));
                if cache_key != self.last_completed_prefix {
                    self.last_prefix_change = Instant::now();
                    self.pending_completion = true;
                    if !self.selection_mode {
                        self.selected_completion = None;
                    }
                }
                // Word boundary: check spelling of the last finished word
                let spell_word = self.context.sentence.split_whitespace().last()
                    .map(|w| w.trim_matches(|c: char| c.is_ascii_punctuation() || c == '«' || c == '»').to_string());
                if let Some(ref w) = spell_word {
                    if !w.is_empty() {
                        self.check_spelling(w);
                    }
                }
                // Sentence boundary: run grammar check
                self.run_grammar_check();
            } else {
                // No word, no context: clear and run grammar
                self.completions.clear();
                self.open_completions.clear();
                self.last_completed_prefix.clear();
                // Check spelling + grammar on the last word/sentence
                let spell_word = self.context.sentence.split_whitespace().last()
                    .map(|w| w.trim_matches(|c: char| c.is_ascii_punctuation() || c == '«' || c == '»').to_string());
                if let Some(ref w) = spell_word {
                    if !w.is_empty() {
                        self.check_spelling(w);
                    }
                }
                self.run_grammar_check();
            }

            // Pre-warm forward cache: when masked_sentence changes, run single_forward
            // eagerly so the first keystroke hits the cache instead of waiting ~150ms.
            if let Some(masked) = &self.context.masked_sentence {
                let needs_warmup = match &self.cached_forward {
                    Some((cached_masked, _)) => cached_masked != masked,
                    None => true,
                };
                if needs_warmup {
                    if let Some(model) = &mut self.model {
                        let t = Instant::now();
                        if let Ok((logits, _)) = model.single_forward(masked) {
                            let fwd_ms = t.elapsed().as_millis();
                            self.cached_forward = Some((masked.clone(), logits.clone()));

                            // Pre-warm mtag supplement: use top word-initial tokens as stems
                            // to discover and score dictionary words before user types
                            if let Some(checker) = &self.checker {
                                let t2 = Instant::now();
                                // Get top 15 word-initial tokens as stems
                                let mut stems: Vec<(String, f32)> = model.id_to_token.iter()
                                    .enumerate()
                                    .filter(|(_, tok)| tok.starts_with('Ġ'))
                                    .map(|(i, _)| {
                                        let decoded = model.tokenizer
                                            .decode(&[i as u32], false)
                                            .unwrap_or_default().trim().to_string();
                                        (decoded, logits[i])
                                    })
                                    .filter(|(w, _)| w.len() >= 3)
                                    .collect();
                                stems.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
                                stems.truncate(15);

                                // Find mtag words from these stems + sub-prefixes
                                let mut mtag_set: std::collections::HashSet<String> = std::collections::HashSet::new();
                                for (stem, _) in &stems {
                                    let lower = stem.to_lowercase();
                                    for w in checker.prefix_lookup(&lower, 5) {
                                        mtag_set.insert(w);
                                    }
                                    // Sub-prefixes for compound discovery
                                    let chars: Vec<char> = lower.chars().collect();
                                    for end in 2..chars.len() {
                                        let sub: String = chars[..end].iter().collect();
                                        for w in checker.prefix_lookup(&sub, 3) {
                                            mtag_set.insert(w);
                                        }
                                    }
                                }

                                if !mtag_set.is_empty() {
                                    // Tokenize + first-token filter + frequency boost
                                    let wf = self.wordfreq.as_ref();
                                    let mut all_cands: Vec<(String, Vec<u32>, f32)> = mtag_set.into_iter()
                                        .filter_map(|w| {
                                            let enc = model.tokenizer.encode(format!(" {}", w).as_str(), false).ok()?;
                                            let ids: Vec<u32> = enc.get_ids().to_vec();
                                            if ids.is_empty() { return None; }
                                            let mut score = logits[ids[0] as usize];
                                            if let Some(wf) = wf {
                                                if wf.contains_key(&w) { score += 2.0; } else { score -= 2.0; }
                                            }
                                            Some((w, ids, score))
                                        })
                                        .collect();
                                    all_cands.sort_by(|a, b| b.2.partial_cmp(&a.2).unwrap_or(std::cmp::Ordering::Equal));
                                    all_cands.truncate(10);

                                    let cands: Vec<(String, Vec<u32>)> = all_cands.iter()
                                        .map(|(w, ids, _)| (w.clone(), ids.clone())).collect();
                                    let mut scores: Vec<f32> = all_cands.iter().map(|(_, _, s)| *s).collect();

                                    // Multi-token scoring (max 2 rounds)
                                    let mask_parts: Vec<&str> = masked.splitn(2, "<mask>").collect();
                                    let ctx_b = mask_parts[0].trim_end();
                                    let ctx_a = mask_parts.get(1).map(|s| s.trim_start()).unwrap_or(".");
                                    let max_t = cands.iter().map(|(_, ids)| ids.len()).max().unwrap_or(1).min(3);
                                    for t in 1..max_t {
                                        let to_score: Vec<usize> = cands.iter().enumerate()
                                            .filter(|(_, (_, ids))| ids.len() > t)
                                            .map(|(i, _)| i).collect();
                                        if to_score.is_empty() { break; }
                                        let mut unique_pfx: Vec<Vec<u32>> = Vec::new();
                                        let mut pfx_map: std::collections::HashMap<Vec<u32>, usize> = std::collections::HashMap::new();
                                        let mut c2p: Vec<usize> = Vec::new();
                                        for &i in &to_score {
                                            let tp = cands[i].1[..t].to_vec();
                                            let pidx = *pfx_map.entry(tp.clone()).or_insert_with(|| {
                                                let idx = unique_pfx.len(); unique_pfx.push(tp); idx
                                            });
                                            c2p.push(pidx);
                                        }
                                        let texts: Vec<String> = unique_pfx.iter()
                                            .map(|ids| {
                                                let p = model.tokenizer.decode(ids, false).unwrap_or_default();
                                                format!("{} {}<mask> {}", ctx_b, p.trim(), ctx_a)
                                            }).collect();
                                        if let Ok((bl, _)) = model.batched_forward(&texts) {
                                            for (k, &i) in to_score.iter().enumerate() {
                                                scores[i] += bl[c2p[k]][cands[i].1[t] as usize];
                                            }
                                        }
                                    }

                                    let mtag_scored: Vec<(String, f32)> = cands.iter().enumerate()
                                        .map(|(i, (w, ids))| (w.clone(), scores[i] / ids.len() as f32))
                                        .collect();
                                    self.cached_mtag_supplement = Some((masked.clone(), mtag_scored));
                                    self.cached_right_column = None; // will be recomputed on first keystroke
                                }
                                eprintln!("pre-warmed forward+mtag in {}ms+{}ms", fwd_ms, t2.elapsed().as_millis());
                            } else {
                                eprintln!("pre-warmed forward cache in {}ms", fwd_ms);
                            }
                        }
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
                        self.return_focus_to_word();
                        self.manager.replace_word(&word);
                        self.completions.clear();
            self.open_completions.clear();
                        self.last_completed_prefix.clear();
                        // Force immediate context refresh after replace
                        self.last_poll = Instant::now() - self.poll_interval;
                    }
                }
                self.selection_mode = false;
                self.selected_completion = None;
            }
            if cancel {
                self.return_focus_to_word();
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
        let win_h = if has_content { 250.0 } else { 110.0 };
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

        ctx.request_repaint_after(Duration::from_millis(100));

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

        egui::CentralPanel::default().frame(panel_frame).show(ctx, |ui| {
            // Tab bar with painted dot indicators
            ui.horizontal(|ui| {
                let tab_labels = ["Innhold", "Grammatikk", "Innstillinger", "Debug"];
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
                if close_resp.double_clicked() {
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
                        self.manager.replace_word(&word);
                        self.completions.clear();
                        self.open_completions.clear();
                        self.selected_completion = None;
                        self.selection_mode = false;
                        self.last_completed_prefix.clear();
                        self.last_poll = Instant::now() - self.poll_interval;
                        self.return_focus_to_word();
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
                let active_errors: Vec<usize> = self.writing_errors.iter()
                    .enumerate()
                    .filter(|(_, e)| !e.ignored)
                    .map(|(i, _)| i)
                    .collect();

                if active_errors.is_empty() {
                    ui.label(
                        egui::RichText::new("Ingen feil funnet.")
                            .size(12.0)
                            .color(egui::Color32::from_rgb(0, 140, 60)),
                    );
                } else {
                    ui.label(
                        egui::RichText::new(format!("Feil funnet ({})", active_errors.len()))
                            .size(12.0)
                            .strong()
                            .color(egui::Color32::from_rgb(80, 80, 80)),
                    );
                    ui.add_space(4.0);

                    let mut action: Option<(usize, &str)> = None;

                    // Group grammar errors by sentence_context
                    let mut shown_contexts: std::collections::HashSet<String> = std::collections::HashSet::new();

                    for &idx in &active_errors {
                        let error = &self.writing_errors[idx];

                        // For grammar errors with position > 0, skip — they're shown as alternatives
                        if matches!(error.category, ErrorCategory::Grammar) && error.position > 0 {
                            if shown_contexts.contains(&error.sentence_context) {
                                continue;
                            }
                        }

                        let (icon, color) = match error.category {
                            ErrorCategory::Spelling => (
                                "🔴",
                                egui::Color32::from_rgb(200, 40, 40),
                            ),
                            ErrorCategory::Grammar => (
                                "🟢",
                                rule_color(&error.rule_name),
                            ),
                        };

                        ui.group(|ui| {
                            if matches!(error.category, ErrorCategory::Grammar) {
                                shown_contexts.insert(error.sentence_context.clone());
                                // Sentence-level grammar correction
                                ui.horizontal(|ui| {
                                    ui.label(egui::RichText::new(icon).size(12.0));
                                    ui.label(
                                        egui::RichText::new("Grammatikkfeil")
                                            .strong()
                                            .size(12.0)
                                            .color(color),
                                    );
                                });
                                // Show original with strikethrough
                                ui.label(
                                    egui::RichText::new(&error.word)
                                        .size(11.0)
                                        .strikethrough()
                                        .color(egui::Color32::from_rgb(150, 80, 80)),
                                );
                                // Show all alternatives for this sentence
                                let ctx = error.sentence_context.clone();
                                let alternatives: Vec<usize> = active_errors.iter()
                                    .filter(|&&i| {
                                        let e = &self.writing_errors[i];
                                        matches!(e.category, ErrorCategory::Grammar)
                                            && e.sentence_context == ctx
                                            && !e.suggestion.is_empty()
                                    })
                                    .copied()
                                    .collect();

                                for (alt_num, &alt_idx) in alternatives.iter().enumerate() {
                                    let alt = &self.writing_errors[alt_idx];
                                    ui.horizontal(|ui| {
                                        ui.label(
                                            egui::RichText::new(format!("{}.", alt_num + 1))
                                                .size(11.0)
                                                .strong()
                                                .color(egui::Color32::from_rgb(60, 60, 60)),
                                        );
                                        ui.label(
                                            egui::RichText::new(&alt.suggestion)
                                                .size(11.0)
                                                .strong()
                                                .color(egui::Color32::from_rgb(0, 120, 60)),
                                        );
                                        if ui.small_button("Rett opp").clicked() {
                                            action = Some((alt_idx, "fix"));
                                        }
                                    });
                                    ui.label(
                                        egui::RichText::new(&alt.explanation)
                                            .size(10.0)
                                            .color(egui::Color32::from_rgb(100, 100, 100)),
                                    );
                                }

                                // Ignore button for the whole group
                                ui.horizontal(|ui| {
                                    if ui.button(
                                        egui::RichText::new("Ignorer").size(11.0)
                                    ).clicked() {
                                        // Ignore all alternatives for this sentence
                                        action = Some((idx, "ignore_group"));
                                    }
                                });
                            } else {
                                // Spelling error: word-level display
                                ui.horizontal(|ui| {
                                    ui.label(egui::RichText::new(icon).size(12.0));
                                    ui.label(
                                        egui::RichText::new(format!("«{}»", error.word))
                                            .strong()
                                            .size(12.0)
                                            .color(color),
                                    );
                                    if !error.suggestion.is_empty() {
                                        ui.label(
                                            egui::RichText::new(format!("→ {}", error.suggestion))
                                                .size(12.0)
                                                .strong()
                                                .color(egui::Color32::from_rgb(0, 120, 60)),
                                        );
                                    }
                                });
                                ui.label(
                                    egui::RichText::new(&error.explanation)
                                        .size(11.0)
                                        .color(egui::Color32::from_rgb(80, 80, 80)),
                                );

                                // Action buttons for spelling
                                ui.horizontal(|ui| {
                                    if !error.suggestion.is_empty() {
                                        if ui.button(
                                            egui::RichText::new("Rett opp").size(11.0)
                                        ).clicked() {
                                            action = Some((idx, "fix"));
                                        }
                                    }
                                    if ui.button(
                                        egui::RichText::new("Forslag").size(11.0)
                                    ).clicked() {
                                        action = Some((idx, "suggest"));
                                    }
                                    if ui.button(
                                        egui::RichText::new("Ignorer").size(11.0)
                                    ).clicked() {
                                        action = Some((idx, "ignore"));
                                    }
                                });
                            }
                        });
                        ui.add_space(2.0);
                    }

                    // Handle actions after rendering
                    if let Some((idx, act)) = action {
                        match act {
                            "fix" => {
                                let error = &self.writing_errors[idx];
                                let suggestion = error.suggestion.clone();
                                let word = error.word.clone();
                                let context = error.sentence_context.clone();
                                self.pending_fix = Some((word, suggestion, context));
                                // Mark all alternatives for this sentence as ignored
                                let ctx = self.writing_errors[idx].sentence_context.clone();
                                for e in &mut self.writing_errors {
                                    if e.sentence_context == ctx && matches!(e.category, ErrorCategory::Grammar) {
                                        e.ignored = true;
                                    }
                                }
                            }
                            "suggest" => {
                                let word = self.writing_errors[idx].word.clone();
                                let sentence_ctx = self.writing_errors[idx].sentence_context.clone();
                                let suggestions = self.trigram_suggestions(&word, &sentence_ctx);
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
                                for e in &mut self.writing_errors {
                                    if e.sentence_context == ctx && matches!(e.category, ErrorCategory::Grammar) {
                                        e.ignored = true;
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                }
            }

            // === Suggestion window ===
            if self.suggestion_window.is_some() {
                let mut selected: Option<String> = None;
                let mut open = true;
                let (word, candidates) = self.suggestion_window.as_ref().unwrap();
                let word_clone = word.clone();
                let candidates_clone: Vec<(String, f32)> = candidates.clone();

                egui::Window::new(format!("Forslag for «{}»", word_clone))
                    .open(&mut open)
                    .collapsible(false)
                    .resizable(false)
                    .default_width(250.0)
                    .show(ctx, |ui| {
                        if candidates_clone.is_empty() {
                            ui.label("Ingen forslag funnet.");
                        } else {
                            for (candidate, _score) in candidates_clone.iter().take(10) {
                                ui.horizontal(|ui| {
                                    if ui.small_button("🔊").clicked() {
                                        tts::speak_word(candidate);
                                    }
                                    if ui.button(
                                        egui::RichText::new(candidate).size(12.0).strong()
                                    ).clicked() {
                                        selected = Some(candidate.clone());
                                    }
                                });
                            }
                        }
                    });

                if let Some(replacement) = selected {
                    // Replace in document and mark error as fixed
                    let ctx = self.writing_errors.iter()
                        .find(|e| e.word == word_clone && !e.ignored)
                        .map(|e| e.sentence_context.clone())
                        .unwrap_or_default();
                    self.pending_fix = Some((word_clone.clone(), replacement.clone(), ctx));
                    // Update the error's suggestion and mark ignored
                    for e in &mut self.writing_errors {
                        if e.word == word_clone && !e.ignored {
                            e.suggestion = replacement;
                            e.ignored = true;
                            break;
                        }
                    }
                    self.suggestion_window = None;
                }
                if !open {
                    self.suggestion_window = None;
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
                ui.label(egui::RichText::new("Ord:").size(11.0).strong().color(egui::Color32::from_rgb(100, 100, 100)));
                ui.label(
                    egui::RichText::new(if self.context.word.is_empty() { "(tomt)" } else { &self.context.word })
                        .size(13.0)
                        .color(egui::Color32::from_rgb(0, 70, 160)),
                );
                ui.add_space(4.0);
                ui.label(egui::RichText::new("Setning:").size(11.0).strong().color(egui::Color32::from_rgb(100, 100, 100)));
                ui.label(
                    egui::RichText::new(if self.context.sentence.is_empty() { "(tom)" } else { &self.context.sentence })
                        .size(12.0)
                        .color(egui::Color32::from_rgb(50, 50, 50)),
                );
                if let Some(masked) = &self.context.masked_sentence {
                    ui.add_space(4.0);
                    ui.label(egui::RichText::new("Maskert:").size(11.0).strong().color(egui::Color32::from_rgb(100, 100, 100)));
                    let display = if masked.len() > 200 {
                        format!("{}...", &masked[..200])
                    } else {
                        masked.clone()
                    };
                    ui.label(
                        egui::RichText::new(display)
                            .size(10.0)
                            .color(egui::Color32::from_rgb(80, 80, 80)),
                    );
                }
                ui.add_space(6.0);
                if ui.small_button("Kopier til utklippstavle").clicked() {
                    let mut text = format!("Ord: {}\nSetning: {}", self.context.word, self.context.sentence);
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
                eprintln!("ORT_DYLIB_PATH={}", path);
                break;
            }
        }
    }

    let grammar_completion = !std::env::args().any(|a| a == "--no-grammar");
    let use_swipl = std::env::args().any(|a| a == "--swipl");
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
                cc.egui_ctx.set_fonts(fonts);
                eprintln!("Loaded Open Sans font");
            } else {
                eprintln!("Warning: Open Sans font not found at {}", font_path);
            }
            Ok(Box::new(ContextApp::new(grammar_completion, use_swipl, quality)))
        }),
    )
}
