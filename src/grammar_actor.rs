//! Grammar actor — dedicated thread for SWI-Prolog grammar checking.
//!
//! SWI-Prolog uses raw C pointers (!Send), so the checker must stay on
//! one thread. The main thread sends sentences via channel, the actor
//! checks them and sends results back. Never blocks the main thread.
//!
//! The actor also calls ctx.request_repaint() after sending results
//! so the GUI wakes up immediately.

use std::sync::mpsc;

use nostos_cognio::grammar::types::{GrammarError, UnknownWord, CompoundCandidate};

use crate::AnyChecker;
use crate::user_dict::UserDict;

/// Request sent to the grammar actor
pub struct GrammarCheckRequest {
    pub sentence: String,
    pub doc_offset: usize,
    pub paragraph_id: String,
    pub sentence_index: usize,
    pub doc_text: String,
    pub user_words: Vec<String>,
}

/// Synchronous check request — for completion grammar filtering
pub struct SyncCheckRequest {
    pub sentence: String,
    pub reply: mpsc::Sender<SyncCheckResponse>,
}

pub struct SyncCheckResponse {
    pub errors: Vec<GrammarError>,
}

/// Batch synchronous check — multiple sentences at once
pub struct SyncBatchRequest {
    pub sentences: Vec<String>,
    pub reply: mpsc::Sender<SyncBatchResponse>,
}

pub struct SyncBatchResponse {
    pub results: Vec<Vec<GrammarError>>,
}

/// Batch ASYNCHRONOUS check — results come back via the handle's batch
/// receiver channel rather than a per-request reply channel. Avoids the
/// caller having to block their thread waiting on the actor.
///
/// Used by the BERT spelling pipeline so the BERT worker can hand off
/// candidate-grammar filtering without blocking on SWI-Prolog.
pub struct AsyncBatchRequest {
    pub request_id: u64,
    pub sentences: Vec<String>,
}

pub struct AsyncBatchResponse {
    pub request_id: u64,
    pub results: Vec<Vec<GrammarError>>,
}

/// Result sent back from the grammar actor — full check_sentence_full output.
pub struct GrammarCheckResponse {
    pub sentence: String,
    pub doc_offset: usize,
    pub paragraph_id: String,
    pub sentence_index: usize,
    pub errors: Vec<GrammarError>,
    pub unknown_words: Vec<UnknownWord>,
    pub compound_candidates: Vec<CompoundCandidate>,
}

/// Messages the actor can receive
pub enum ActorMessage {
    Async(GrammarCheckRequest),
    Sync(SyncCheckRequest),
    SyncBatch(SyncBatchRequest),
    AsyncBatch(AsyncBatchRequest),
}

/// Handle to communicate with the grammar actor
pub struct GrammarActorHandle {
    sender: mpsc::Sender<ActorMessage>,
    receiver: mpsc::Receiver<GrammarCheckResponse>,
    batch_receiver: mpsc::Receiver<AsyncBatchResponse>,
    next_batch_id: std::cell::Cell<u64>,
}

impl GrammarActorHandle {
    /// Get a sender clone for use by other threads (e.g. BERT worker for grammar filtering)
    pub fn sender_clone(&self) -> mpsc::Sender<ActorMessage> {
        self.sender.clone()
    }

    /// Send a sentence for checking (non-blocking)
    pub fn check_sentence(&self, sentence: &str, doc_offset: usize, paragraph_id: &str, sentence_index: usize) {
        self.check_sentence_with_doc(sentence, doc_offset, paragraph_id, sentence_index, "", &[])
    }

    pub fn check_sentence_with_doc(&self, sentence: &str, doc_offset: usize, paragraph_id: &str, sentence_index: usize, doc_text: &str, user_words: &[String]) {
        crate::log!("ACTOR ASYNC SEND: '{}'", sentence);
        let _ = self.sender.send(ActorMessage::Async(GrammarCheckRequest {
            sentence: sentence.to_string(),
            doc_offset,
            paragraph_id: paragraph_id.to_string(),
            doc_text: doc_text.to_string(),
            sentence_index,
            user_words: user_words.to_vec(),
        }));
    }

    /// Synchronous grammar check — blocks until the actor responds.
    /// Used for completion grammar filtering.
    pub fn check_sentence_sync(&self, sentence: &str) -> Vec<GrammarError> {
        let (reply_tx, reply_rx) = mpsc::channel();
        let _ = self.sender.send(ActorMessage::Sync(SyncCheckRequest {
            sentence: sentence.to_string(),
            reply: reply_tx,
        }));
        match reply_rx.recv() {
            Ok(resp) => resp.errors,
            Err(_) => Vec::new(),
        }
    }

    /// Batch synchronous grammar check — one round-trip for all sentences.
    pub fn check_sentences_batch(&self, sentences: &[String]) -> Vec<Vec<GrammarError>> {
        let (reply_tx, reply_rx) = mpsc::channel();
        let _ = self.sender.send(ActorMessage::SyncBatch(SyncBatchRequest {
            sentences: sentences.to_vec(),
            reply: reply_tx,
        }));
        match reply_rx.recv() {
            Ok(resp) => resp.results,
            Err(_) => sentences.iter().map(|_| Vec::new()).collect(),
        }
    }

    /// Non-blocking poll for results
    pub fn try_recv(&self) -> Option<GrammarCheckResponse> {
        self.receiver.try_recv().ok()
    }

    /// Send a batch of sentences for asynchronous checking. Returns the
    /// `request_id` the response will carry. Caller polls `try_recv_batch`
    /// to retrieve results. Used by the spelling pipeline so the BERT
    /// worker doesn't block on SWI-Prolog (`feedback_*` memory: BERT
    /// worker must not call grammar synchronously).
    pub fn send_async_batch(&self, sentences: Vec<String>) -> u64 {
        let id = self.next_batch_id.get();
        self.next_batch_id.set(id.wrapping_add(1));
        crate::log!("ACTOR ASYNC BATCH SEND: id={} {} sentences", id, sentences.len());
        let _ = self.sender.send(ActorMessage::AsyncBatch(AsyncBatchRequest {
            request_id: id,
            sentences,
        }));
        id
    }

    /// Non-blocking poll for async-batch results.
    pub fn try_recv_batch(&self) -> Option<AsyncBatchResponse> {
        self.batch_receiver.try_recv().ok()
    }
}

/// Grammar batch check using a sender clone. Can be called from any thread.
pub fn grammar_batch_via_sender(sender: &mpsc::Sender<ActorMessage>, sentences: &[String]) -> Vec<Vec<GrammarError>> {
    {
        use std::io::Write;
        if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true)
            .open(std::env::temp_dir().join("spell.log")) {
            let _ = writeln!(f, "ACTOR BATCH SEND: {} sentences", sentences.len());
        }
    }
    let (reply_tx, reply_rx) = mpsc::channel();
    let _ = sender.send(ActorMessage::SyncBatch(SyncBatchRequest {
        sentences: sentences.to_vec(),
        reply: reply_tx,
    }));
    match reply_rx.recv() {
        Ok(resp) => resp.results,
        Err(_) => sentences.iter().map(|_| Vec::new()).collect(),
    }
}

/// Spawn grammar actor that loads SWI-Prolog on its own thread.
/// SWI-Prolog must be initialized and used on the same thread.
pub fn spawn_grammar_actor_with_loader(
    swipl_path: String,
    dict_path: String,
    grammar_rules_path: String,
    syntaxer_dir: String,
    compound_data: String,
    repaint_ctx: egui::Context,
    compound_fst: Option<std::sync::Arc<fst::raw::Fst<Vec<u8>>>>,
    language: std::sync::Arc<dyn language::LanguageBundle>,
    wordfreq: Option<std::sync::Arc<std::collections::HashMap<String, u64>>>,
) -> GrammarActorHandle {
    let (req_tx, req_rx) = mpsc::channel::<ActorMessage>();
    let (resp_tx, resp_rx) = mpsc::channel::<GrammarCheckResponse>();
    let (batch_tx, batch_rx) = mpsc::channel::<AsyncBatchResponse>();

    std::thread::Builder::new()
        .name("grammar-actor".into())
        .spawn(move || {
            // SWI-Prolog's libswipl depends on @rpath/libgmp.10.dylib.
            // libswipl's rpath includes @executable_path/../Frameworks,
            // so we symlink libgmp there (one level up from the binary).
            let exe_dir = std::env::current_exe()
                .ok()
                .and_then(|p| p.parent().map(|d| d.to_path_buf()))
                .unwrap_or_else(|| std::path::PathBuf::from("."));
            let frameworks_dir = exe_dir.join("../Frameworks").canonicalize()
                .unwrap_or_else(|_| {
                    let d = exe_dir.join("../Frameworks");
                    let _ = std::fs::create_dir_all(&d);
                    d
                });
            #[cfg(unix)]
            {
                let gmp_link = frameworks_dir.join("libgmp.10.dylib");
                let gmp_source = std::path::Path::new("/Applications/SWI-Prolog.app/Contents/Frameworks/libgmp.10.dylib");
                if !gmp_link.exists() && gmp_source.exists() {
                    let _ = std::fs::create_dir_all(&frameworks_dir);
                    let _ = std::os::unix::fs::symlink(gmp_source, &gmp_link);
                    eprintln!("Grammar actor: symlinked libgmp to {:?}", gmp_link);
                }
            }

            // Load checker on THIS thread so SWI-Prolog stays on one thread
            let checker: AnyChecker = match nostos_cognio::grammar::swipl_checker::SwiGrammarChecker::new(
                &swipl_path,
                &dict_path,
                &grammar_rules_path,
                &syntaxer_dir,
            ) {
                Ok(c) => {
                    eprintln!("Grammar actor: SWI-Prolog loaded on actor thread");
                    AnyChecker::Swi(c)
                }
                Err(e) => {
                    eprintln!("FATAL: Grammar actor: SWI-Prolog failed to load: {}", e);
                    eprintln!("SWI-Prolog is required. No fallback.");
                    panic!("SWI-Prolog failed to load: {}", e);
                }
            };

            let mut checker = checker;
            let compound_fst_ref2 = compound_fst.clone();
            let lang_ref2 = language.clone();
            let wf_ref2 = wordfreq.clone();
            while let Ok(msg) = req_rx.recv() {
                match msg {
                    ActorMessage::Async(req) => {
                        let compound_check2: Option<Box<dyn Fn(&str) -> bool>> = compound_fst_ref2.as_ref().map(|fst| {
                            let fst = fst.clone();
                            let lang = lang_ref2.clone();
                            let wf = wf_ref2.clone();
                            // Pass the same validators used in the spelling path so
                            // compound_walker applies the same strictness here as in main.rs —
                            // parts must be real dictionary words, last part must be a noun.
                            let analyzer = match &checker {
                                AnyChecker::Swi(c) => c.analyzer().clone(),
                            };
                            let _ = (&fst, &lang, &wf); // fuzzy walk no longer used here
                            Box::new(move |word: &str| -> bool {
                                nostos_cognio::grammar::is_novel_compound(&analyzer, word)
                            }) as Box<dyn Fn(&str) -> bool>
                        });
                        let compound_ref2: Option<&dyn Fn(&str) -> bool> = compound_check2.as_ref().map(|b| b.as_ref());
                        let mut result = checker.check_sentence_full_with_compound(&req.sentence, &req.doc_text, compound_ref2);
                        // Filter out user dict words from unknowns
                        if !req.user_words.is_empty() {
                            result.unknown_words.retain(|u| {
                                !req.user_words.iter().any(|uw| uw.eq_ignore_ascii_case(&u.word))
                            });
                        }
                        let _ = resp_tx.send(GrammarCheckResponse {
                            sentence: req.sentence,
                            doc_offset: req.doc_offset,
                            paragraph_id: req.paragraph_id,
                            sentence_index: req.sentence_index,
                            errors: result.errors,
                            unknown_words: result.unknown_words,
                            compound_candidates: result.compound_candidates,
                        });
                        repaint_ctx.request_repaint();
                    }
                    ActorMessage::Sync(req) => {
                        let errors = checker.check_sentence(&req.sentence);
                        let _ = req.reply.send(SyncCheckResponse { errors });
                    }
                    ActorMessage::SyncBatch(req) => {
                        let results: Vec<Vec<GrammarError>> = req.sentences.iter()
                            .map(|s| checker.check_sentence(s))
                            .collect();
                        let _ = req.reply.send(SyncBatchResponse { results });
                    }
                    ActorMessage::AsyncBatch(req) => {
                        let results: Vec<Vec<GrammarError>> = req.sentences.iter()
                            .map(|s| checker.check_sentence(s))
                            .collect();
                        let _ = batch_tx.send(AsyncBatchResponse {
                            request_id: req.request_id,
                            results,
                        });
                        repaint_ctx.request_repaint();
                    }
                }
            }
            std::mem::forget(checker);
            loop { std::thread::park(); }
        })
        .expect("Failed to spawn grammar actor");

    GrammarActorHandle {
        sender: req_tx,
        receiver: resp_rx,
        batch_receiver: batch_rx,
        next_batch_id: std::cell::Cell::new(1),
    }
}
