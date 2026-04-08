/// Nynorsk spelling suggestion test using the shared pipeline.
///
/// This is the EXACT same code path as the Word add-in spelling check,
/// just with Nynorsk resources wired in:
///   - fullform_nn.mfst (mtag FST for NN)
///   - wordfreq_nn.tsv  (NN word frequencies)
///   - nynorsk/grammar_rules.pl  (NN SWI-Prolog grammar)
///   - norbert4_base_int8.onnx   (shared BM/NN BERT model)
///
/// Runs `generate_spelling_candidates` + `score_and_rerank` against
/// a small list of known-misspelled sentences and prints full scoring
/// telemetry (including the intermediate bert scores).
///
/// Usage: cargo run --release --bin test_spelling_nn

use nostos_cognio::model::Model;
use nostos_cognio::grammar::swipl_checker::SwiGrammarChecker;
use std::collections::HashMap;
use std::path::PathBuf;

use acatts_rust::spelling_scorer::{generate_spelling_candidates, score_and_rerank};

fn main() {
    // Set ORT_DYLIB_PATH if not already set — mirrors the idiom used by
    // the other test_bert_* bins.
    if std::env::var("ORT_DYLIB_PATH").is_err() {
        let ort_candidates = vec![
            PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../onnxruntime/lib/libonnxruntime.dylib"),
            PathBuf::from("/usr/local/lib/libonnxruntime.dylib"),
            PathBuf::from("/opt/homebrew/lib/libonnxruntime.dylib"),
        ];
        for p in &ort_candidates {
            if p.exists() {
                unsafe { std::env::set_var("ORT_DYLIB_PATH", p); }
                eprintln!("Using ORT dylib: {}", p.display());
                break;
            }
        }
    }

    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));

    // Paths — match the runtime resolve_language("nn") values. All paths
    // are resolved the same way as the production app so this test exercises
    // the identical code path.
    let training = {
        let mac = base.join("../contexter-repo/training-data");
        if mac.exists() { mac } else { base.join("../../contexter-repo/training-data") }
    };

    let onnx_path = training.join("onnx/norbert4_base_int8.onnx");
    let tok_path = training.join("onnx/tokenizer.json");
    println!("Loading NorBERT4 (shared BM/NN)...");
    let mut model = Model::load(onnx_path.to_str().unwrap(), tok_path.to_str().unwrap())
        .expect("Failed to load model");
    println!("Loaded. Vocab: {}", model.vocab_size());

    // Nynorsk mtag FST
    let dict_path = {
        let mac = base.join("../rustSpell/mtag-rs/data/fullform_nn.mfst");
        if mac.exists() { mac } else { base.join("../../rustSpell/mtag-rs/data/fullform_nn.mfst") }
    };

    // Nynorsk grammar rules (live in /dyslex/nynorsk/, NOT in /dyslex/syntaxer/)
    let grammar_rules_path = {
        let mac = base.join("../nynorsk/grammar_rules.pl");
        if mac.exists() { mac } else { base.join("../../nynorsk/grammar_rules.pl") }
    };
    let nynorsk_dir = {
        let mac = base.join("../nynorsk");
        if mac.exists() { mac } else { base.join("../../nynorsk") }
    };

    let swipl_dll = if cfg!(target_os = "macos") {
        "/Applications/SWI-Prolog.app/Contents/Frameworks/libswipl.dylib"
    } else {
        "C:/Program Files/swipl/bin/libswipl.dll"
    };

    println!("Loading NN analyzer from {}...", dict_path.display());
    let analyzer = mtag::Analyzer::new(dict_path.to_str().unwrap())
        .expect("Failed to load analyzer");

    println!("Loading NN SWI grammar checker from {}...", grammar_rules_path.display());
    let mut checker = SwiGrammarChecker::new(
        swipl_dll,
        dict_path.to_str().unwrap(),
        grammar_rules_path.to_str().unwrap(),
        nynorsk_dir.to_str().unwrap(),
    ).expect("Failed to load SWI grammar checker");
    println!("Ready.\n");

    // Nynorsk wordfreq
    let wf_path = training.join("wordfreq_nn.tsv");
    println!("Loading NN wordfreq from {}", wf_path.display());
    let wf = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);
    println!("Loaded {} entries", wf.len());

    let empty_doc: HashMap<String, u16> = HashMap::new();
    let empty_user: Vec<String> = Vec::new();

    // (sentence, misspelled_word, expected_top1_correction)
    //
    // The CRITICAL test case is the first one: `ikkkje` (4 k's) should
    // correct to `ikkje` (NN for "not"). This is the case the user
    // reported as failing in the GUI — every BERT score came back as -inf
    // and the scorer fell back to picking `nykkje` instead.
    let test_cases = vec![
        ("Eg er ikkkje frå Bergen.",     "ikkkje",  "ikkje"),
        ("Eg er ikkkje fra Bergen.",     "ikkkje",  "ikkje"),  // BM-flavoured "fra"
        // Mid-typing state (this is what the GUI sees when the user types
        // a space after "ikkkje" — the spelling check fires at that point
        // with context_after essentially empty or just trailing spaces).
        // Reproduces the user-reported bug: BERT scores all came back as
        // -inf in the GUI for this exact state.
        ("Eg er ikkkje",                 "ikkkje",  "ikkje"),  // no context after
        ("Eg er ikkkje ",                "ikkkje",  "ikkje"),  // trailing space
        ("Eg er ikkkje f",               "ikkkje",  "ikkje"),  // partial next word
        ("Ho likkjer ikkje musikk.",     "likkjer", "likar"),
        ("Han har ein stoor bil.",       "stoor",   "stor"),
        ("Vi spelar fottball.",          "fottball","fotball"),
        ("Barna fekk bøler og brus.",    "bøler",   "boller|bollar"),
    ];

    let mut pass = 0;
    let mut fail = 0;

    for (sentence, misspelled, expected) in &test_cases {
        println!("\n{}", "=".repeat(70));
        println!("Test: '{}' → expected '{}'", misspelled, expected);
        println!("Sentence: '{}'", sentence);

        // Phase 1: Generate candidates (same code as app)
        let candidates = generate_spelling_candidates(
            &analyzer,
            Some(&wf),
            &empty_user,
            &empty_doc,
            misspelled,
            sentence,
        );
        println!("  Phase 1: {} candidates", candidates.len());

        let expected_alts: Vec<&str> = expected.split('|').collect();
        let expected_in_pool = expected_alts.iter().any(|alt| {
            candidates.iter().any(|(c, _)| c == &alt.to_lowercase())
        });
        if expected_in_pool {
            println!("  ✓ expected word IS in candidate pool");
            for alt in &expected_alts {
                if let Some((pos, (_, score))) = candidates.iter().enumerate().find(|(_, (c, _))| c == &alt.to_lowercase()) {
                    println!("    '{}' at Phase1 rank #{} ortho={:.3}", alt, pos + 1, score);
                }
            }
        } else {
            println!("  ✗ expected word NOT in candidate pool!");
            println!("  (top 10 candidates in pool: {:?})",
                candidates.iter().take(10).map(|(c, s)| format!("{}={:.2}", c, s)).collect::<Vec<_>>());
        }

        // Phase 2: Score and rerank (same code as app's BERT worker)
        let sentence_lower = sentence.to_lowercase();
        let word_lower = misspelled.to_lowercase();
        let (context_before, context_after) = if let Some(pos) = sentence_lower.find(&word_lower) {
            (sentence_lower[..pos].trim_end().to_string(), sentence_lower[pos + word_lower.len()..].trim_start().to_string())
        } else {
            (sentence_lower.clone(), String::new())
        };

        // Mirror main.rs: if context_after is empty/whitespace, append sentinel "."
        // to BOTH context_after AND the sentence so score_and_rerank's word-position
        // extraction works. Without this, find("") returns Some(0) and word_lower
        // is extracted as "" → garbled corrected_sent → NEG_INFINITY for every candidate.
        let (context_after, sentence_for_scorer) = if context_after.trim().is_empty() {
            (".".to_string(), format!("{}.", sentence))
        } else {
            (context_after, sentence.to_string())
        };

        println!("  Context: before='{}' after='{}'", context_before, context_after);

        let mut grammar_check = |sentences: &[String]| -> Vec<Vec<nostos_cognio::grammar::types::GrammarError>> {
            sentences.iter().map(|s| checker.check_sentence(s)).collect()
        };

        let results = score_and_rerank(
            &mut model,
            &mut grammar_check,
            &candidates,
            &context_before,
            &context_after,
            &sentence_for_scorer,
        );

        println!("  Top 10 after BERT scoring:");
        for (i, (w, s)) in results.iter().take(10).enumerate() {
            let marker = if expected_alts.contains(&w.as_str()) { " ✓ EXPECTED" } else { "" };
            let inf_marker = if s.is_infinite() { " ⚠️ INFINITE" } else { "" };
            println!("    #{}: '{}' score={:.6}{}{}", i + 1, w, s, marker, inf_marker);
        }

        if let Some((top, _)) = results.first() {
            if expected_alts.contains(&top.as_str()) {
                println!("  PASS");
                pass += 1;
            } else {
                println!("  FAIL: got '{}', expected '{}'", top, expected);
                fail += 1;
            }
        } else {
            println!("  FAIL: no candidates");
            fail += 1;
        }
    }

    println!("\n{}", "=".repeat(70));
    println!("Results: {}/{} passed", pass, pass + fail);

    // -----------------------------------------------------------------------
    // Secondary probe: does the NN grammar checker flag the BM `fra`?
    // The user's original sentence, AFTER spelling correction, becomes
    // `Eg er ikkje fra Bergen.`. In NN, `frå` is the preposition "from".
    // `fra` is in the NN FST only as `prop unormert` (a rare proper noun,
    // non-standard). So the NN grammar rules are our only chance to catch
    // `fra` and suggest `frå`.
    // -----------------------------------------------------------------------
    println!("\n{}", "=".repeat(70));
    println!("Grammar check probe: does NN grammar flag 'fra'?");
    for sentence in &[
        "Eg er ikkje fra Bergen.",
        "Eg er ikkje frå Bergen.",
        "Eg kjem fra Oslo.",
        "Eg kjem frå Oslo.",
    ] {
        let errors = checker.check_sentence(sentence);
        println!("\n  Sentence: '{}'", sentence);
        if errors.is_empty() {
            println!("    (no grammar errors reported)");
        } else {
            for e in &errors {
                println!("    ERROR: word='{}' rule='{}' suggestion='{}'",
                    e.word, e.rule_name, e.suggestion);
                if !e.explanation.is_empty() {
                    println!("      explanation: {}", e.explanation);
                }
            }
        }
    }

    if fail > 0 {
        std::process::exit(1);
    }
}
