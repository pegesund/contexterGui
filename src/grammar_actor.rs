use std::sync::mpsc;
use std::sync::{Arc, Mutex};
use std::thread;

use nostos_cognio::grammar::types::GrammarError;
use nostos_cognio::model::Model;

use crate::AnyChecker;

/// Messages sent from main thread to grammar actor
pub enum GrammarRequest {
    /// Check a single sentence (synchronous — caller waits for response)
    CheckSentence {
        text: String,
        reply: mpsc::Sender<Vec<GrammarError>>,
    },
    /// Split text by Prolog sentence boundaries (synchronous)
    SplitByProlog {
        text: String,
        reply: mpsc::Sender<Option<Vec<String>>>,
    },
    /// Suggest compound correction (synchronous)
    SuggestCompound {
        word: String,
        reply: mpsc::Sender<Option<String>>,
    },
    /// Run a full document grammar scan (asynchronous — results sent back via channel)
    ScanDocument {
        doc_text: String,
        sentences: Vec<(String, usize)>,
        clean_hashes: std::collections::HashSet<u64>,
        existing_errors: Vec<crate::WritingError>,
    },
    /// Shutdown the actor
    Shutdown,
}

/// Results from a full document scan
pub struct ScanResult {
    pub errors: Vec<crate::WritingError>,
    pub new_clean_hashes: std::collections::HashSet<u64>,
}

/// Handle to communicate with the grammar actor
pub struct GrammarActorHandle {
    pub sender: mpsc::Sender<GrammarRequest>,
    pub scan_receiver: mpsc::Receiver<ScanResult>,
    pub scanning: Arc<Mutex<bool>>,
}

impl GrammarActorHandle {
    /// Check a single sentence synchronously (blocks until response)
    pub fn check_sentence(&self, text: &str) -> Vec<GrammarError> {
        let (reply_tx, reply_rx) = mpsc::channel();
        let _ = self.sender.send(GrammarRequest::CheckSentence {
            text: text.to_string(),
            reply: reply_tx,
        });
        reply_rx.recv().unwrap_or_default()
    }

    /// Split text by Prolog sentence boundaries synchronously
    pub fn split_by_prolog(&self, text: &str) -> Option<Vec<String>> {
        let (reply_tx, reply_rx) = mpsc::channel();
        let _ = self.sender.send(GrammarRequest::SplitByProlog {
            text: text.to_string(),
            reply: reply_tx,
        });
        reply_rx.recv().ok().flatten()
    }

    /// Suggest compound correction synchronously
    pub fn suggest_compound(&self, word: &str) -> Option<String> {
        let (reply_tx, reply_rx) = mpsc::channel();
        let _ = self.sender.send(GrammarRequest::SuggestCompound {
            word: word.to_string(),
            reply: reply_tx,
        });
        reply_rx.recv().ok().flatten()
    }

    /// Start an async document scan. Returns immediately.
    pub fn start_scan(
        &self,
        doc_text: String,
        sentences: Vec<(String, usize)>,
        clean_hashes: std::collections::HashSet<u64>,
        existing_errors: Vec<crate::WritingError>,
    ) {
        if let Ok(mut s) = self.scanning.lock() {
            *s = true;
        }
        let _ = self.sender.send(GrammarRequest::ScanDocument {
            doc_text,
            sentences,
            clean_hashes,
            existing_errors,
        });
    }

    /// Check if a scan result is ready (non-blocking)
    pub fn try_recv_scan(&self) -> Option<ScanResult> {
        match self.scan_receiver.try_recv() {
            Ok(result) => {
                if let Ok(mut s) = self.scanning.lock() {
                    *s = false;
                }
                Some(result)
            }
            Err(_) => None,
        }
    }

    /// Is a scan currently running?
    pub fn is_scanning(&self) -> bool {
        self.scanning.lock().map(|s| *s).unwrap_or(false)
    }
}

/// Spawn the grammar actor thread.
/// The checker is created inside the actor thread (important for SWI-Prolog thread safety).
pub fn spawn_grammar_actor(
    checker: AnyChecker,
    model: Option<Arc<Mutex<Model>>>,
) -> GrammarActorHandle {
    let (req_tx, req_rx) = mpsc::channel::<GrammarRequest>();
    let (scan_tx, scan_rx) = mpsc::channel::<ScanResult>();
    let scanning = Arc::new(Mutex::new(false));
    let scanning_clone = scanning.clone();

    thread::Builder::new()
        .name("grammar-actor".to_string())
        .spawn(move || {
            actor_loop(checker, model, req_rx, scan_tx, scanning_clone);
        })
        .expect("Failed to spawn grammar actor thread");

    GrammarActorHandle {
        sender: req_tx,
        scan_receiver: scan_rx,
        scanning,
    }
}

fn actor_loop(
    mut checker: AnyChecker,
    model: Option<Arc<Mutex<Model>>>,
    rx: mpsc::Receiver<GrammarRequest>,
    scan_tx: mpsc::Sender<ScanResult>,
    scanning: Arc<Mutex<bool>>,
) {
    eprintln!("Grammar actor: started on thread {:?}", thread::current().id());

    while let Ok(msg) = rx.recv() {
        match msg {
            GrammarRequest::CheckSentence { text, reply } => {
                let errors = checker.check_sentence(&text);
                let _ = reply.send(errors);
            }
            GrammarRequest::SplitByProlog { text, reply } => {
                let result = checker.split_by_prolog(&text);
                let _ = reply.send(result);
            }
            GrammarRequest::SuggestCompound { word, reply } => {
                let result = checker.suggest_compound(&word);
                let _ = reply.send(result);
            }
            GrammarRequest::ScanDocument {
                doc_text: _doc_text,
                sentences,
                clean_hashes,
                existing_errors,
            } => {
                let result = run_scan(
                    &mut checker,
                    &model,
                    sentences,
                    clean_hashes,
                    existing_errors,
                );
                if let Ok(mut s) = scanning.lock() {
                    *s = false;
                }
                let _ = scan_tx.send(result);
            }
            GrammarRequest::Shutdown => {
                eprintln!("Grammar actor: shutting down");
                break;
            }
        }
    }
}

/// Run a full grammar scan on the given sentences.
/// This is the heavy work that was previously done in update_grammar_errors().
fn run_scan(
    checker: &mut AnyChecker,
    _model: &Option<Arc<Mutex<Model>>>,
    sentences: Vec<(String, usize)>,
    mut clean_hashes: std::collections::HashSet<u64>,
    existing_errors: Vec<crate::WritingError>,
) -> ScanResult {
    use crate::{ErrorCategory, WritingError, hash_str, replace_word_at_position};
    let mut new_errors: Vec<WritingError> = Vec::new();

    for (trimmed, doc_offset) in &sentences {
        let sent_h = hash_str(trimmed);

        // Skip if already known clean
        if clean_hashes.contains(&sent_h) {
            continue;
        }

        // Skip if existing errors already cover this occurrence
        let has_errors = existing_errors.iter().any(|e| {
            e.sentence_context == *trimmed && e.doc_offset == *doc_offset && !e.ignored
        }) || new_errors.iter().any(|e| {
            e.sentence_context == *trimmed && e.doc_offset == *doc_offset && !e.ignored
        });
        if has_errors {
            continue;
        }

        // Skip if there are spelling errors for this sentence (fix spelling first)
        let has_spelling = existing_errors.iter().any(|e| {
            e.sentence_context == *trimmed
                && e.doc_offset == *doc_offset
                && !e.ignored
                && matches!(e.category, ErrorCategory::Spelling)
        });
        if has_spelling {
            continue;
        }

        let errors = checker.check_sentence(trimmed);
        if errors.is_empty() {
            clean_hashes.insert(sent_h);
            continue;
        }

        for ge in &errors {
            eprintln!("  Grammar actor: '{}' → '{}' ({})", ge.word, ge.suggestion, ge.rule_name);
        }

        // BERT scoring of corrections
        let corrections = if let Some(model_arc) = _model {
            if let Ok(mut model) = model_arc.lock() {
                best_sentence_corrections_standalone(&mut *model, trimmed, &errors)
            } else {
                Vec::new()
            }
        } else {
            Vec::new()
        };

        if corrections.is_empty() {
            // Direct grammar fix fallback
            let errors_with_suggestions: Vec<_> = errors.iter()
                .filter(|e| !e.suggestion.is_empty())
                .collect();
            if !errors_with_suggestions.is_empty() {
                for (i, ge) in errors_with_suggestions.iter().enumerate() {
                    let first_alt = ge.suggestion.split('|').next().unwrap_or(&ge.suggestion);
                    let corrected = crate::replace_word_at_position(trimmed, &ge.word, first_alt);
                    if corrected.trim() == trimmed.trim() {
                        continue;
                    }
                    new_errors.push(WritingError {
                        category: ErrorCategory::Grammar,
                        word: trimmed.to_string(),
                        suggestion: corrected,
                        explanation: format!("\u{ab}{}\u{bb} \u{2192} \u{ab}{}\u{bb}: {}", ge.word, first_alt, ge.explanation),
                        rule_name: ge.rule_name.clone(),
                        sentence_context: trimmed.to_string(),
                        doc_offset: *doc_offset,
                        position: i,
                        ignored: false,
                    });
                }
            } else {
                let first = &errors[0];
                new_errors.push(WritingError {
                    category: ErrorCategory::Grammar,
                    word: trimmed.to_string(),
                    suggestion: String::new(),
                    explanation: first.explanation.clone(),
                    rule_name: first.rule_name.clone(),
                    sentence_context: trimmed.to_string(),
                    doc_offset: *doc_offset,
                    position: 0,
                    ignored: false,
                });
            }
        }

        for (i, (corrected, explanation, rule_name, score)) in corrections.iter().enumerate() {
            if corrected.trim() == trimmed.trim() {
                continue;
            }
            new_errors.push(WritingError {
                category: ErrorCategory::Grammar,
                word: trimmed.to_string(),
                suggestion: corrected.clone(),
                explanation: explanation.clone(),
                rule_name: rule_name.clone(),
                sentence_context: trimmed.to_string(),
                doc_offset: *doc_offset,
                position: i,
                ignored: false,
            });
        }
    }

    ScanResult {
        errors: new_errors,
        new_clean_hashes: clean_hashes,
    }
}

/// Standalone version of best_sentence_corrections for the actor thread
fn best_sentence_corrections_standalone(
    model: &mut Model,
    sentence: &str,
    errors: &[GrammarError],
) -> Vec<(String, String, String, f32)> {
    let mut candidates: Vec<(String, String, String)> = Vec::new();

    for ge in errors {
        if ge.suggestion.is_empty() { continue; }
        for alt in ge.suggestion.split('|') {
            let corrected = crate::replace_word_at_position(sentence, &ge.word, alt);
            if corrected.trim() != sentence.trim() {
                candidates.push((
                    corrected,
                    format!("\u{ab}{}\u{bb} \u{2192} \u{ab}{}\u{bb}: {}", ge.word, alt, ge.explanation),
                    ge.rule_name.clone(),
                ));
            }
        }
    }

    if candidates.is_empty() { return Vec::new(); }

    // Score with BERT
    let mut scored: Vec<(String, String, String, f32)> = candidates.into_iter()
        .map(|(corrected, explanation, rule)| {
            let score = bert_sentence_score_standalone(model, &corrected);
            (corrected, explanation, rule, score)
        })
        .collect();

    scored.sort_by(|a, b| b.3.partial_cmp(&a.3).unwrap_or(std::cmp::Ordering::Equal));
    scored.truncate(3);
    scored
}

/// Standalone BERT sentence scoring for the actor thread
fn bert_sentence_score_standalone(model: &mut Model, sentence: &str) -> f32 {
    let words: Vec<&str> = sentence.split_whitespace().collect();
    if words.is_empty() { return f32::NEG_INFINITY; }

    let mut total_score: f32 = 0.0;
    for i in 0..words.len() {
        let masked: String = words.iter().enumerate()
            .map(|(j, w)| if j == i { "<mask>" } else { *w })
            .collect::<Vec<_>>()
            .join(" ");

        if let Ok((logits, _)) = model.single_forward(&masked) {
            let word_clean = words[i].trim_matches(|c: char| c.is_ascii_punctuation());
            let token_with_g = format!("\u{120}{}", word_clean.to_lowercase());
            let token_id = model.tokenizer.token_to_id(&token_with_g)
                .or_else(|| model.tokenizer.token_to_id(&word_clean.to_lowercase()));
            if let Some(tid) = token_id {
                total_score += logits[tid as usize];
            }
        }
    }
    total_score / words.len() as f32
}
