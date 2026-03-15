//! Grammar actor — dedicated thread for SWI-Prolog grammar checking.
//!
//! SWI-Prolog uses raw C pointers (!Send), so the checker must stay on
//! one thread. The main thread sends sentences via channel, the actor
//! checks them and sends results back. Never blocks the main thread.
//!
//! The actor also calls ctx.request_repaint() after sending results
//! so the GUI wakes up immediately.

use std::sync::mpsc;

use nostos_cognio::grammar::types::GrammarError;

use crate::AnyChecker;

/// Request sent to the grammar actor
pub struct GrammarCheckRequest {
    pub sentence: String,
    pub doc_offset: usize,
}

/// Result sent back from the grammar actor
pub struct GrammarCheckResponse {
    pub sentence: String,
    pub doc_offset: usize,
    pub errors: Vec<GrammarError>,
}

/// Handle to communicate with the grammar actor
pub struct GrammarActorHandle {
    sender: mpsc::Sender<GrammarCheckRequest>,
    receiver: mpsc::Receiver<GrammarCheckResponse>,
}

impl GrammarActorHandle {
    /// Send a sentence for checking (non-blocking)
    pub fn check_sentence(&self, sentence: &str, doc_offset: usize) {
        let _ = self.sender.send(GrammarCheckRequest {
            sentence: sentence.to_string(),
            doc_offset,
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
                let errors = checker.check_sentence(&req.sentence);
                let _ = resp_tx.send(GrammarCheckResponse {
                    sentence: req.sentence,
                    doc_offset: req.doc_offset,
                    errors,
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
                    eprintln!("Grammar actor: SWI-Prolog failed ({}), trying neorusticus", e);
                    match nostos_cognio::grammar::GrammarChecker::new(&dict_path, &compound_data) {
                        Ok(c) => {
                            eprintln!("Grammar actor: neorusticus loaded ({} clauses)", c.clause_count());
                            AnyChecker::Neo(c)
                        }
                        Err(e2) => {
                            eprintln!("Grammar actor: no checker available: {}", e2);
                            return;
                        }
                    }
                }
            };

            let mut checker = checker;
            while let Ok(req) = req_rx.recv() {
                let errors = checker.check_sentence(&req.sentence);
                let _ = resp_tx.send(GrammarCheckResponse {
                    sentence: req.sentence,
                    doc_offset: req.doc_offset,
                    errors,
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
