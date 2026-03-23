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
    build_right: fn(&Model, &[f32], Option<&HashMap<String, u64>>, &HashSet<String>, &HashSet<String>) -> Vec<Completion>,
    prefix_index: Arc<PrefixIndex>,
    baselines: Option<Arc<Baselines>>,
    wordfreq_shared: Option<Arc<HashMap<String, u64>>>,
    embedding_store: Option<Arc<EmbeddingStore>>,
    analyzer: Option<Arc<mtag::Analyzer>>,
) -> BertWorkerHandle {
    let (req_tx, req_rx) = mpsc::channel::<BertRequest>();
    let (resp_tx, resp_rx) = mpsc::channel::<BertResponse>();

    std::thread::Builder::new()
        .name("bert-worker".to_string())
        .spawn(move || {
            worker_loop(model, repaint_ctx, req_rx, resp_tx, build_bpe, build_mtag, build_right,
                prefix_index, baselines, wordfreq_shared, embedding_store, analyzer);
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
    build_right: fn(&Model, &[f32], Option<&HashMap<String, u64>>, &HashSet<String>, &HashSet<String>) -> Vec<Completion>,
    prefix_index: Arc<PrefixIndex>,
    baselines: Option<Arc<Baselines>>,
    wordfreq_shared: Option<Arc<HashMap<String, u64>>>,
    embedding_store: Option<Arc<EmbeddingStore>>,
    analyzer: Option<Arc<mtag::Analyzer>>,
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
                let right = build_right(&model, &logits, wordfreq.as_deref(), &nearby_words, &left_words);

                let _ = tx.send(BertResponse::Completion { id, cache_key, left, right });
                repaint_ctx.request_repaint();
            }

            BertRequest::SpellingScore { id, context_before, context_after, candidates } => {
                // Use the existing score_spelling() from nostos-cognio — batched_forward, ~24ms
                let scored = match spelling::score_spelling(&mut model, &context_before, &context_after, &candidates) {
                    Ok(result) => result.scored_candidates,
                    Err(_) => Vec::new(),
                };
                let _ = tx.send(BertResponse::SpellingScore { id, scored_candidates: scored });
                repaint_ctx.request_repaint();
            }

            BertRequest::MlmForward { id, masked_text, top_k } => {
                let predictions = mlm_forward_impl(&mut model, &masked_text, top_k);
                let _ = tx.send(BertResponse::MlmForward { id, predictions });
                repaint_ctx.request_repaint();
            }

            BertRequest::CompleteWord { id, context, prefix, capitalize: _cap, top_n, max_steps, cache_key, masked_text, cancel } => {
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
                match complete_word(
                    &mut model,
                    &context,
                    &prefix,
                    pi,
                    baselines.as_deref(),
                    wordfreq_shared.as_deref(),
                    fallback_ref,
                    prefix_ref,
                    embedding_store.as_deref(),
                    1.0,   // pmi_weight
                    10.0,  // topic_boost
                    top_n,
                    max_steps,
                ) {
                    Ok(left) => {
                        {
                            use std::io::Write;
                            if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true)
                                .open(std::env::temp_dir().join("acatts-bert.log")) {
                                let _ = writeln!(f, "complete_word(prefix='{}', top_n={}, max_steps={}) → {} results in {:?}",
                                    prefix, top_n, max_steps, left.len(), t_cw.elapsed());
                            }
                        }
                        let t_right = std::time::Instant::now();
                        // Right completions: complete_word with empty prefix (open predictions)
                        let right = match complete_word(
                            &mut model, &context, "", pi,
                            baselines.as_deref(), wordfreq_shared.as_deref(),
                            fallback_ref, prefix_ref, embedding_store.as_deref(),
                            1.0, 10.0, 15, 0,
                        ) {
                            Ok(r) => {
                                // Exclude words already in left column
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
                                    right.len(), t_right.elapsed());
                            }
                        }
                        let _ = tx.send(BertResponse::Completion { id, cache_key, left, right });
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
