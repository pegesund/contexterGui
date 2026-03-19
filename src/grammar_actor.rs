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

/// Request sent to the grammar actor
pub struct GrammarCheckRequest {
    pub sentence: String,
    pub doc_offset: usize,
    pub paragraph_id: String,
    pub sentence_index: usize,
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

/// Handle to communicate with the grammar actor
pub struct GrammarActorHandle {
    sender: mpsc::Sender<GrammarCheckRequest>,
    receiver: mpsc::Receiver<GrammarCheckResponse>,
}

impl GrammarActorHandle {
    /// Send a sentence for checking (non-blocking)
    pub fn check_sentence(&self, sentence: &str, doc_offset: usize, paragraph_id: &str, sentence_index: usize) {
        let _ = self.sender.send(GrammarCheckRequest {
            sentence: sentence.to_string(),
            doc_offset,
            paragraph_id: paragraph_id.to_string(),
            sentence_index,
        });
    }

    /// Non-blocking poll for results
    pub fn try_recv(&self) -> Option<GrammarCheckResponse> {
        self.receiver.try_recv().ok()
    }
}

/// Spawn the grammar actor thread. Takes ownership of the checker.
/// The egui Context is used to trigger repaint when results are ready.
pub fn spawn_grammar_actor(
    checker: AnyChecker,
    repaint_ctx: egui::Context,
) -> GrammarActorHandle {
    let (req_tx, req_rx) = mpsc::channel::<GrammarCheckRequest>();
    let (resp_tx, resp_rx) = mpsc::channel::<GrammarCheckResponse>();

    std::thread::Builder::new()
        .name("grammar-actor".into())
        .spawn(move || {
            let mut checker = checker;
            while let Ok(req) = req_rx.recv() {
                let result = checker.check_sentence_full(&req.sentence);
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
        })
        .expect("Failed to spawn grammar actor");

    GrammarActorHandle {
        sender: req_tx,
        receiver: resp_rx,
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
) -> GrammarActorHandle {
    let (req_tx, req_rx) = mpsc::channel::<GrammarCheckRequest>();
    let (resp_tx, resp_rx) = mpsc::channel::<GrammarCheckResponse>();

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
            let gmp_link = frameworks_dir.join("libgmp.10.dylib");
            let gmp_source = std::path::Path::new("/Applications/SWI-Prolog.app/Contents/Frameworks/libgmp.10.dylib");
            if !gmp_link.exists() && gmp_source.exists() {
                let _ = std::fs::create_dir_all(&frameworks_dir);
                let _ = std::os::unix::fs::symlink(gmp_source, &gmp_link);
                eprintln!("Grammar actor: symlinked libgmp to {:?}", gmp_link);
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
            while let Ok(req) = req_rx.recv() {
                let result = checker.check_sentence_full(&req.sentence);
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
        })
        .expect("Failed to spawn grammar actor");

    GrammarActorHandle {
        sender: req_tx,
        receiver: resp_rx,
    }
}
