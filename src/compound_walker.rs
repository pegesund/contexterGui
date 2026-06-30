//! Re-export of `nostos_cognio::compound_walker`.
//!
//! The compound walker is the only fuzzy-lookup path for Bokmål and
//! Nynorsk (see `feedback_spelling_pipeline_duplicated.md`). It used to
//! live here, but `dyslexia_tests` in nostos-cognio also needs to run
//! the *exact same* candidate generation as the GUI — without that
//! shared location the tests would inevitably drift from the GUI's
//! behaviour again. The implementation moved to nostos-cognio in
//! 2026-06; this module exists only so existing Spell call sites
//! (main.rs, grammar_actor.rs, the test_compound_* bins) keep
//! compiling.

pub use nostos_cognio::compound_walker::{
    compound_fuzzy_walk,
    load_fst_from_mfst,
    CompoundPart,
    CompoundResult,
};
