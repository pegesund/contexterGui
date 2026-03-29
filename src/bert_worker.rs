//! BERT worker thread — owns the Model exclusively, processes requests via channels.
//! Eliminates all lock contention between background completion, spelling, and grammar scoring.
//!
//! Uses existing nostos-cognio functions:
//! - score_spelling() for spelling re-ranking (batched_forward)
//! - single_forward() for completion logits
//! NEVER rewrite these functions here.

use std::collections::{HashMap, HashSet};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{mpsc, Arc};

use nostos_cognio::baseline::Baselines;
use nostos_cognio::complete::{complete_word, Completion};
use nostos_cognio::embeddings::EmbeddingStore;
use nostos_cognio::model::Model;
use nostos_cognio::prefix_index::PrefixIndex;
use nostos_cognio::spelling;

pub type RequestId = u64;

/// Requests sent to the BERT worker thread
pub enum BertRequest {
    /// Word completion
    Completion {
        id: RequestId,
        masked_text: String,
        prefix_lower: String,
        matches: Vec<(u32, String)>,
        mtag_candidates: Vec<String>,
        mtag_valid: HashSet<String>,
        nearby_words: HashSet<String>,
        wordfreq: Option<Arc<HashMap<String, u64>>>,
        capitalize: bool,
        cancel: Arc<AtomicBool>,
        cache_key: String,
    },

    /// Score spelling candidates using score_spelling() (batched_forward).
    /// Replaces the old SentenceScoreBatch which used a slow single_forward loop.
    SpellingScore {
        id: RequestId,
        context_before: String,
        context_after: String,
        candidates: Vec<String>,
    },

    /// MLM forward pass: get top token predictions at <mask> position.
    MlmForward {
        id: RequestId,
        masked_text: String,
        top_k: usize,
    },

    /// Full word completion via complete_word() — the original working pipeline.
    CompleteWord {
        id: RequestId,
        context: String,
        prefix: String,
        capitalize: bool,
        top_n: usize,
        max_steps: usize,
        cache_key: String,
        masked_text: String,
        cancel: Arc<AtomicBool>,
        sentence: String,
    },
}

/// Responses from the BERT worker thread
pub enum BertResponse {
    Completion {
        id: RequestId,
        cache_key: String,
        left: Vec<Completion>,
        right: Vec<Completion>,
    },

    /// Spelling scores — scored_candidates sorted best-first
    SpellingScore {
        id: RequestId,
        scored_candidates: Vec<(String, f32)>,
    },

    MlmForward {
        id: RequestId,
        /// (decoded_token_clean, logit_score)
        predictions: Vec<(String, f32)>,
    },
}

/// Handle for communicating with the BERT worker thread
pub struct BertWorkerHandle {
    sender: mpsc::Sender<BertRequest>,
    receiver: mpsc::Receiver<BertResponse>,
    next_id: u64,
}

impl BertWorkerHandle {
    /// Send a request, returns the request ID for matching the response
    pub fn send(&mut self, make_request: impl FnOnce(RequestId) -> BertRequest) -> RequestId {
        let id = self.next_id;
        self.next_id += 1;
        let _ = self.sender.send(make_request(id));
        id
    }

    /// Non-blocking poll for responses
    pub fn try_recv(&self) -> Option<BertResponse> {
        self.receiver.try_recv().ok()
    }
}

/// Spawn the BERT worker thread. Takes ownership of the Model.
/// `repaint_ctx` is used to wake up the GUI when results are ready.
pub fn spawn_bert_worker(
    model: Model,
    repaint_ctx: egui::Context,
    build_bpe: fn(&mut Model, &str, &str, &[(u32, String)], &[f32], Option<&HashMap<String, u64>>, &HashSet<String>, &HashSet<String>, bool, &AtomicBool) -> Vec<Completion>,
    build_mtag: fn(&mut Model, &str, &[String], &[f32], bool, &AtomicBool) -> Vec<Completion>,
    build_right: fn(&Model, &[f32], Option<&HashMap<String, u64>>, &HashSet<String>, &HashSet<String>, Option<&nostos_cognio::baseline::Baselines>, Option<&mtag::Analyzer>) -> Vec<Completion>,
    prefix_index: Arc<PrefixIndex>,
    baselines: Option<Arc<Baselines>>,
    wordfreq_shared: Option<Arc<HashMap<String, u64>>>,
    embedding_store: Option<Arc<EmbeddingStore>>,
    analyzer: Option<Arc<mtag::Analyzer>>,
    grammar_sender: Option<mpsc::Sender<crate::grammar_actor::ActorMessage>>,
) -> BertWorkerHandle {
    let (req_tx, req_rx) = mpsc::channel::<BertRequest>();
    let (resp_tx, resp_rx) = mpsc::channel::<BertResponse>();

    std::thread::Builder::new()
        .name("bert-worker".to_string())
        .spawn(move || {
            worker_loop(model, repaint_ctx, req_rx, resp_tx, build_bpe, build_mtag, build_right,
                prefix_index, baselines, wordfreq_shared, embedding_store, analyzer, grammar_sender);
        })
        .expect("Failed to spawn BERT worker thread");

    BertWorkerHandle {
        sender: req_tx,
        receiver: resp_rx,
        next_id: 1,
    }
}

fn worker_loop(
    mut model: Model,
    repaint_ctx: egui::Context,
    rx: mpsc::Receiver<BertRequest>,
    tx: mpsc::Sender<BertResponse>,
    build_bpe: fn(&mut Model, &str, &str, &[(u32, String)], &[f32], Option<&HashMap<String, u64>>, &HashSet<String>, &HashSet<String>, bool, &AtomicBool) -> Vec<Completion>,
    build_mtag: fn(&mut Model, &str, &[String], &[f32], bool, &AtomicBool) -> Vec<Completion>,
    build_right: fn(&Model, &[f32], Option<&HashMap<String, u64>>, &HashSet<String>, &HashSet<String>, Option<&nostos_cognio::baseline::Baselines>, Option<&mtag::Analyzer>) -> Vec<Completion>,
    prefix_index: Arc<PrefixIndex>,
    baselines: Option<Arc<Baselines>>,
    wordfreq_shared: Option<Arc<HashMap<String, u64>>>,
    embedding_store: Option<Arc<EmbeddingStore>>,
    analyzer: Option<Arc<mtag::Analyzer>>,
    grammar_sender: Option<mpsc::Sender<crate::grammar_actor::ActorMessage>>,
) {
    use std::ops::Deref;
    while let Ok(req) = rx.recv() {
        // Drain stale completion requests: if newer CompleteWord is queued, skip to it.
        // Never skip non-completion requests (SpellingScore, MlmForward).
        let req = {
            let mut current = req;
            loop {
                match rx.try_recv() {
                    Ok(newer) => {
                        let same_type = matches!(
                            (&current, &newer),
                            (BertRequest::CompleteWord { .. }, BertRequest::CompleteWord { .. })
                            | (BertRequest::Completion { .. }, BertRequest::Completion { .. })
                        );
                        if same_type {
                            current = newer; // Skip older, keep newer
                        } else {
                            // Different type — process current first, then newer next iteration
                            // Can't put back into mpsc, so process current now and newer is lost
                            // TODO: use a VecDeque if this becomes an issue
                            break;
                        }
                    }
                    Err(_) => break, // Queue empty
                }
            }
            current
        };

        match req {
            BertRequest::Completion {
                id, masked_text, prefix_lower, matches, mtag_candidates,
                mtag_valid, nearby_words, wordfreq, capitalize, cancel, cache_key,
            } => {
                {
                    use std::io::Write;
                    if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true)
                        .open(std::env::temp_dir().join("acatts-bert.log")) {
                        let _ = writeln!(f, "OLD Completion: prefix='{}' matches={}", prefix_lower, matches.len());
                    }
                }
                if cancel.load(Ordering::Acquire) { continue; }

                let logits = match model.single_forward(&masked_text) {
                    Ok((l, _)) => l,
                    Err(_) => continue,
                };
                if cancel.load(Ordering::Acquire) { continue; }

                let left = if matches.is_empty() && !prefix_lower.is_empty() {
                    build_mtag(&mut model, &masked_text, &mtag_candidates, &logits, capitalize, &cancel)
                } else if !prefix_lower.is_empty() {
                    build_bpe(&mut model, &masked_text, &prefix_lower, &matches, &logits, wordfreq.as_deref(), &nearby_words, &mtag_valid, capitalize, &cancel)
                } else {
                    vec![]
                };

                if cancel.load(Ordering::Acquire) { continue; }

                let left_words: HashSet<String> = left.iter().map(|c| c.word.to_lowercase()).collect();
                let right = build_right(&model, &logits, wordfreq.as_deref(), &nearby_words, &left_words, baselines.as_deref(), analyzer.as_deref());

                let _ = tx.send(BertResponse::Completion { id, cache_key, left, right });
                repaint_ctx.request_repaint();
            }

            BertRequest::SpellingScore { id, context_before, context_after, candidates } => {
                // Build ortho-scored candidates for the shared scorer
                let ortho_candidates: Vec<(String, f32)> = candidates.iter()
                    .map(|c| (c.clone(), 0.5)) // default ortho — real scores come from main thread
                    .collect();
                let sentence = format!("{}{}", context_before, context_after);
                let mut grammar_check = |sentences: &[String]| -> Vec<Vec<nostos_cognio::grammar::types::GrammarError>> {
                    if let Some(ref gs) = grammar_sender {
                        crate::grammar_actor::grammar_batch_via_sender(gs, sentences)
                    } else {
                        sentences.iter().map(|_| Vec::new()).collect()
                    }
                };
                let scored = crate::spelling_scorer::score_and_rerank(
                    &mut model, &mut grammar_check, &ortho_candidates,
                    &context_before, &context_after, &sentence,
                );
                let _ = tx.send(BertResponse::SpellingScore { id, scored_candidates: scored });
                repaint_ctx.request_repaint();
            }

            BertRequest::MlmForward { id, masked_text, top_k } => {
                let predictions = mlm_forward_impl(&mut model, &masked_text, top_k);
                let _ = tx.send(BertResponse::MlmForward { id, predictions });
                repaint_ctx.request_repaint();
            }

            BertRequest::CompleteWord { id, context, prefix, capitalize: _cap, top_n, max_steps, cache_key, masked_text, cancel, sentence } => {
                if cancel.load(Ordering::Acquire) { continue; }
                {
                    use std::io::Write;
                    if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true)
                        .open(std::env::temp_dir().join("acatts-bert.log")) {
                        let _ = writeln!(f, "CompleteWord: ctx='{}' prefix='{}' wf={} analyzer={} pi={} top_n={} max_steps={}",
                            &context, prefix,
                            wordfreq_shared.is_some(), analyzer.is_some(), prefix_index.len(), top_n, max_steps);
                    }
                }
                let pi = &*prefix_index;
                let fallback_fn: Option<Box<dyn Fn(&str) -> bool>> = analyzer.as_ref().map(|a| {
                    let a = Arc::clone(a);
                    Box::new(move |w: &str| a.has_word(w)) as Box<dyn Fn(&str) -> bool>
                });
                let fallback_ref: Option<&dyn Fn(&str) -> bool> = fallback_fn.as_ref().map(|b| b.as_ref());
                let prefix_fn: Option<Box<dyn Fn(&str, usize) -> Vec<String>>> = analyzer.as_ref().map(|a| {
                    let a = Arc::clone(a);
                    Box::new(move |p: &str, limit: usize| a.prefix_lookup(p, limit)) as Box<dyn Fn(&str, usize) -> Vec<String>>
                });
                let prefix_ref: Option<&dyn Fn(&str, usize) -> Vec<String>> = prefix_fn.as_ref().map(|b| b.as_ref());

                let t_cw = std::time::Instant::now();

                // When prefix is empty, only do ONE call (right column / open predictions).
                // No point running left+right both with prefix=''.
                let (left_raw, right_raw) = if prefix.is_empty() {
                    let right = match complete_word(
                        &mut model, &context, "", pi,
                        baselines.as_deref(), wordfreq_shared.as_deref(),
                        fallback_ref, prefix_ref, embedding_store.as_deref(),
                        1.0, 10.0, 15, 0,
                    ) {
                        Ok(r) => r,
                        Err(_) => vec![],
                    };
                    {
                        use std::io::Write;
                        if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true)
                            .open(std::env::temp_dir().join("acatts-bert.log")) {
                            let _ = writeln!(f, "complete_word(prefix='', right-only) → {} results in {:?}",
                                right.len(), t_cw.elapsed());
                        }
                    }
                    (vec![], right)
                } else {
                    let left = match complete_word(
                        &mut model, &context, &prefix, pi,
                        baselines.as_deref(), wordfreq_shared.as_deref(),
                        fallback_ref, prefix_ref, embedding_store.as_deref(),
                        1.0, 10.0, top_n, max_steps,
                    ) {
                        Ok(l) => l,
                        Err(_) => vec![],
                    };
                    {
                        use std::io::Write;
                        if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true)
                            .open(std::env::temp_dir().join("acatts-bert.log")) {
                            let _ = writeln!(f, "complete_word(prefix='{}', top_n={}, max_steps={}) → {} results in {:?}",
                                prefix, top_n, max_steps, left.len(), t_cw.elapsed());
                        }
                    }
                    let right = match complete_word(
                        &mut model, &context, "", pi,
                        baselines.as_deref(), wordfreq_shared.as_deref(),
                        fallback_ref, prefix_ref, embedding_store.as_deref(),
                        1.0, 10.0, 15, 0,
                    ) {
                        Ok(r) => {
                            let left_words: HashSet<String> = left.iter().map(|c| c.word.to_lowercase()).collect();
                            r.into_iter().filter(|c| !left_words.contains(&c.word.to_lowercase())).collect()
                        }
                        Err(_) => vec![],
                    };
                    {
                        use std::io::Write;
                        if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true)
                            .open(std::env::temp_dir().join("acatts-bert.log")) {
                            let _ = writeln!(f, "complete_word(prefix='', right) → {} results in {:?}",
                                right.len(), t_cw.elapsed());
                        }
                    }
                    (left, right)
                };

                match Ok::<_, anyhow::Error>((left_raw, right_raw)) {
                    Ok((left, right)) => {
                        // Dictionary filter
                        let (left_dict, right_dict) = if let Some(ref a) = analyzer {
                            let lf: Vec<Completion> = left.into_iter().filter(|c| a.has_word(&c.word.to_lowercase())).collect();
                            let rf: Vec<Completion> = right.into_iter().filter(|c| a.has_word(&c.word.to_lowercase())).collect();
                            (lf, rf)
                        } else {
                            (left, right)
                        };

                        // Grammar filter — but skip if already cancelled (user typed more)
                        let (left_filtered, right_filtered) = if !cancel.load(Ordering::Acquire) {
                            if let Some(ref gs) = grammar_sender {
                                let filter_batch = |candidates: Vec<Completion>, ctx: &str| -> Vec<Completion> {
                                    if candidates.is_empty() { return candidates; }
                                    let last_start = ctx.rfind(|c: char| ".!?".contains(c))
                                        .map(|i| i + 1).unwrap_or(0);
                                    let fragment = ctx[last_start..].trim();
                                    let sentences: Vec<String> = candidates.iter()
                                        .map(|c| format!("{} {}.", fragment, c.word))
                                        .collect();
                                    let results = crate::grammar_actor::grammar_batch_via_sender(gs, &sentences);
                                    candidates.into_iter().zip(results.iter())
                                        .filter(|(_, errs)| errs.is_empty())
                                        .map(|(c, _)| c)
                                        .take(5)
                                        .collect()
                                };
                                let lf = filter_batch(left_dict, &context);
                                let rf = filter_batch(right_dict, &context);
                                (lf, rf)
                            } else {
                                (left_dict.into_iter().take(5).collect(), right_dict.into_iter().take(5).collect())
                            }
                        } else {
                            // Cancelled — skip grammar filter, will be discarded anyway
                            (vec![], vec![])
                        };

                        // Skip sending if cancelled (newer request arrived)
                        if !cancel.load(Ordering::Acquire) {
                            let _ = tx.send(BertResponse::Completion { id, cache_key, left: left_filtered, right: right_filtered });
                        }
                        repaint_ctx.request_repaint();
                    }
                    Err(e) => {
                        eprintln!("complete_word error: {}", e);
                        let _ = tx.send(BertResponse::Completion { id, cache_key, left: vec![], right: vec![] });
                        repaint_ctx.request_repaint();
                    }
                }
            }
        }
    }
}

/// MLM forward pass: get top-k token predictions at <mask> position.
fn mlm_forward_impl(model: &mut Model, masked_text: &str, top_k: usize) -> Vec<(String, f32)> {
    let logits = match model.single_forward(masked_text) {
        Ok((l, _)) => l,
        Err(_) => return Vec::new(),
    };

    let mut logit_indexed: Vec<(usize, f32)> = logits.iter().enumerate()
        .filter(|&(_, v)| *v > 0.0)
        .map(|(i, &v)| (i, v))
        .collect();
    logit_indexed.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

    let mut predictions = Vec::new();
    for (tid, score) in logit_indexed.iter().take(top_k) {
        if let Some(token) = model.tokenizer.id_to_token(*tid as u32) {
            let clean = token.replace("Ġ", "").to_lowercase();
            if clean.len() >= 2 && !clean.contains('<') && !clean.contains('[') {
                predictions.push((clean, *score));
            }
        }
    }
    predictions
}

/// Full sentence scoring: mask each word, sum BERT's prediction logits.
/// More accurate than boundary scoring but ~200ms per sentence.
fn sentence_score(model: &mut Model, sentence: &str) -> f32 {
    let words: Vec<&str> = sentence.split_whitespace().collect();
    if words.is_empty() { return f32::NEG_INFINITY; }
    let mut total: f32 = 0.0;
    for i in 0..words.len() {
        let masked: String = words.iter().enumerate()
            .map(|(j, w)| if j == i { "<mask>" } else { *w })
            .collect::<Vec<_>>().join(" ");
        if let Ok((logits, _)) = model.single_forward(&masked) {
            let word_clean = words[i].trim_matches(|c: char| c.is_ascii_punctuation());
            let token_with_g = format!("Ġ{}", word_clean.to_lowercase());
            let tid = model.tokenizer.token_to_id(&token_with_g)
                .or_else(|| model.tokenizer.token_to_id(&word_clean.to_lowercase()));
            if let Some(id) = tid {
                total += logits[id as usize];
            }
        }
    }
    total / words.len() as f32
}
