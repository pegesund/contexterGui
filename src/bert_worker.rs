//! BERT worker thread — owns the Model exclusively, processes requests via channels.
//! Eliminates all lock contention between background completion, spelling, and grammar scoring.

use std::collections::{HashMap, HashSet};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{mpsc, Arc};

use nostos_cognio::complete::Completion;
use nostos_cognio::model::Model;

pub type RequestId = u64;

/// Requests sent to the BERT worker thread
pub enum BertRequest {
    /// Word completion (replaces the spawned background thread)
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

    /// Score a batch of sentences using pseudo-log-likelihood.
    /// Returns one score per sentence. Used by spelling re-ranking,
    /// grammar correction scoring, and consonant confusion.
    SentenceScoreBatch {
        id: RequestId,
        sentences: Vec<String>,
    },

    /// MLM forward pass: get top token predictions at <mask> position.
    /// Used by spelling MLM fallback.
    MlmForward {
        id: RequestId,
        masked_text: String,
        top_k: usize,
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

    SentenceScoreBatch {
        id: RequestId,
        scores: Vec<f32>,
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
pub fn spawn_bert_worker(
    model: Model,
    build_bpe: fn(&mut Model, &str, &str, &[(u32, String)], &[f32], Option<&HashMap<String, u64>>, &HashSet<String>, &HashSet<String>, bool, &AtomicBool) -> Vec<Completion>,
    build_mtag: fn(&mut Model, &str, &[String], &[f32], bool, &AtomicBool) -> Vec<Completion>,
    build_right: fn(&Model, &[f32], Option<&HashMap<String, u64>>, &HashSet<String>, &HashSet<String>) -> Vec<Completion>,
) -> BertWorkerHandle {
    let (req_tx, req_rx) = mpsc::channel::<BertRequest>();
    let (resp_tx, resp_rx) = mpsc::channel::<BertResponse>();

    std::thread::Builder::new()
        .name("bert-worker".to_string())
        .spawn(move || {
            worker_loop(model, req_rx, resp_tx, build_bpe, build_mtag, build_right);
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
    rx: mpsc::Receiver<BertRequest>,
    tx: mpsc::Sender<BertResponse>,
    build_bpe: fn(&mut Model, &str, &str, &[(u32, String)], &[f32], Option<&HashMap<String, u64>>, &HashSet<String>, &HashSet<String>, bool, &AtomicBool) -> Vec<Completion>,
    build_mtag: fn(&mut Model, &str, &[String], &[f32], bool, &AtomicBool) -> Vec<Completion>,
    build_right: fn(&Model, &[f32], Option<&HashMap<String, u64>>, &HashSet<String>, &HashSet<String>) -> Vec<Completion>,
) {
    while let Ok(req) = rx.recv() {
        match req {
            BertRequest::Completion {
                id, masked_text, prefix_lower, matches, mtag_candidates,
                mtag_valid, nearby_words, wordfreq, capitalize, cancel, cache_key,
            } => {
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
            }

            BertRequest::SentenceScoreBatch { id, sentences } => {
                let scores: Vec<f32> = sentences.iter()
                    .map(|s| sentence_score_impl(&mut model, s))
                    .collect();
                let _ = tx.send(BertResponse::SentenceScoreBatch { id, scores });
            }

            BertRequest::MlmForward { id, masked_text, top_k } => {
                let predictions = mlm_forward_impl(&mut model, &masked_text, top_k);
                let _ = tx.send(BertResponse::MlmForward { id, predictions });
            }
        }
    }
}

/// Score a sentence using BERT pseudo-log-likelihood.
/// For each word, mask it and check how well BERT predicts the actual word.
fn sentence_score_impl(model: &mut Model, sentence: &str) -> f32 {
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
            let token_with_g = format!("Ġ{}", word_clean.to_lowercase());
            let token_id = model.tokenizer.token_to_id(&token_with_g)
                .or_else(|| model.tokenizer.token_to_id(&word_clean.to_lowercase()));
            if let Some(tid) = token_id {
                total_score += logits[tid as usize];
            }
        }
    }
    total_score / words.len() as f32
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
