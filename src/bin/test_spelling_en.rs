/// English spelling suggestion test using ModernBERT.
///
/// Same code path as Norwegian spelling check, but with:
///   - fullform_en.mfst (mtag FST for English)
///   - ModernBERT-base INT8 ONNX (English BERT model)
///
/// Tests two scenarios:
///   1. Mid-sentence: misspelled word with context on both sides
///   2. End-of-sentence: misspelled word at the end (empty context_after)
///
/// Usage: cargo run --release --bin test_spelling_en

use nostos_cognio::model::Model;
use std::collections::HashMap;
use std::path::PathBuf;

use acatts_rust::spelling_scorer::{generate_spelling_candidates, score_and_rerank};

fn main() {
    // Set ORT_DYLIB_PATH
    if std::env::var("ORT_DYLIB_PATH").is_err() {
        let candidates = vec![
            PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../onnxruntime/lib/libonnxruntime.dylib"),
            PathBuf::from("/usr/local/lib/libonnxruntime.dylib"),
            PathBuf::from("/opt/homebrew/lib/libonnxruntime.dylib"),
        ];
        for p in &candidates {
            if p.exists() {
                unsafe { std::env::set_var("ORT_DYLIB_PATH", p); }
                eprintln!("Using ORT dylib: {}", p.display());
                break;
            }
        }
    }

    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));

    // ModernBERT English model
    let onnx_path = PathBuf::from("/tmp/modernbert-en/model_int8.onnx");
    let tok_path = PathBuf::from("/tmp/modernbert-en/tokenizer.json");
    println!("Loading ModernBERT-base (English)...");
    let mut model = Model::load(onnx_path.to_str().unwrap(), tok_path.to_str().unwrap())
        .expect("Failed to load model");
    println!("Loaded. Vocab: {}, Mask token ID: {}", model.vocab_size(), model.mask_token_id);

    // English mtag FST
    let dict_path = {
        let mac = base.join("../rustSpell/mtag-rs/data/fullform_en.mfst");
        if mac.exists() { mac } else { base.join("../../rustSpell/mtag-rs/data/fullform_en.mfst") }
    };
    println!("Loading EN dictionary from {}...", dict_path.display());
    let analyzer = mtag::Analyzer::new(dict_path.to_str().unwrap())
        .expect("Failed to load analyzer");

    let empty_doc: HashMap<String, u16> = HashMap::new();
    let empty_user: Vec<String> = Vec::new();

    // (sentence, misspelled_word, expected_top1_correction, description)
    let test_cases: Vec<(&str, &str, &str, &str)> = vec![
        // --- Mid-sentence tests (context on both sides) ---
        ("I left early becuase I was tired.", "becuase", "because|became",
            "Mid: common typo 'becuase' → 'because'"),
        ("She is a beautful person.", "beautful", "beautiful",
            "Mid: missing letter 'beautful' → 'beautiful'"),
        ("He went to the schoool yesterday.", "schoool", "school",
            "Mid: extra letter 'schoool' → 'school'"),
        ("The package was delivvered this morning.", "delivvered", "delivered",
            "Mid: doubled letter 'delivvered' → 'delivered'"),
        ("I like to reed books.", "reed", "read",
            "Mid: wrong vowel 'reed' → 'read'"),

        // --- End-of-sentence tests (no context after) ---
        ("I want to go to the libary", "libary", "library",
            "End: missing letter 'libary' → 'library'"),
        ("She went to the hosptial", "hosptial", "hospital",
            "End: transposed letters 'hosptial' → 'hospital'"),
        ("He is very inteligent", "inteligent", "intelligent",
            "End: missing letter 'inteligent' → 'intelligent'"),
        ("The children are playying", "playying", "playing",
            "End: extra letter 'playying' → 'playing'"),
        ("I need to buy some groseries", "groseries", "groceries",
            "End: wrong letter 'groseries' → 'groceries'"),
    ];

    let mut pass = 0;
    let mut fail = 0;
    let mut results_table: Vec<(String, String, String, String, bool)> = Vec::new();

    for (sentence, misspelled, expected, desc) in &test_cases {
        println!("\n{}", "=".repeat(70));
        println!("Test: '{}' → expected '{}'", misspelled, expected);
        println!("Sentence: '{}'", sentence);
        println!("({})", desc);

        // Phase 1: Generate candidates
        let candidates = generate_spelling_candidates(
            &analyzer,
            None, // no wordfreq for English yet
            &empty_user,
            &empty_doc,
            misspelled,
            sentence,
            &language::EnglishLanguage,
        );
        println!("  Phase 1: {} candidates", candidates.len());

        let expected_alts: Vec<&str> = expected.split('|').collect();
        let expected_in_pool = expected_alts.iter().any(|alt| {
            candidates.iter().any(|(c, _)| c == &alt.to_lowercase())
        });
        if expected_in_pool {
            println!("  ✓ expected word IS in candidate pool");
        } else {
            println!("  ✗ expected word NOT in candidate pool!");
            println!("  (top 10: {:?})",
                candidates.iter().take(10).map(|(c, s)| format!("{}={:.2}", c, s)).collect::<Vec<_>>());
        }

        // Phase 2: Score and rerank with BERT
        let sentence_lower = sentence.to_lowercase();
        let word_lower = misspelled.to_lowercase();
        let (context_before, context_after) = if let Some(pos) = sentence_lower.find(&word_lower) {
            (sentence_lower[..pos].trim_end().to_string(), sentence_lower[pos + word_lower.len()..].trim_start().to_string())
        } else {
            (sentence_lower.clone(), String::new())
        };

        // Sentinel for empty context_after (same as main.rs)
        let (context_after, sentence_for_scorer) = if context_after.trim().is_empty() {
            (".".to_string(), format!("{}.", sentence))
        } else {
            (context_after, sentence.to_string())
        };

        println!("  Context: before='{}' after='{}'", context_before, context_after);

        // No grammar checker for English yet — pass a no-op
        let mut grammar_check = |sentences: &[String]| -> Vec<Vec<nostos_cognio::grammar::types::GrammarError>> {
            sentences.iter().map(|_| vec![]).collect()
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
            let inf_marker = if s.is_infinite() { " ⚠ INFINITE" } else { "" };
            println!("    #{}: '{}' score={:.6}{}{}", i + 1, w, s, marker, inf_marker);
        }

        let ok = if let Some((top, _)) = results.first() {
            if expected_alts.contains(&top.as_str()) {
                println!("  PASS");
                true
            } else {
                println!("  FAIL: got '{}', expected '{}'", top, expected);
                false
            }
        } else {
            println!("  FAIL: no candidates");
            false
        };

        let got = results.first().map(|(w, _)| w.as_str()).unwrap_or("(none)");
        results_table.push((desc.to_string(), misspelled.to_string(), expected.to_string(), got.to_string(), ok));

        if ok { pass += 1; } else { fail += 1; }
    }

    // Summary table
    println!("\n{}", "=".repeat(70));
    println!("RESULTS: {}/{} passed\n", pass, pass + fail);
    println!("{:<45} {:<12} {:<12} {:<12} {}", "Test", "Input", "Expected", "Got", "Status");
    println!("{}", "-".repeat(95));
    for (desc, input, expected, got, ok) in &results_table {
        let status = if *ok { "PASS" } else { "FAIL" };
        println!("{:<45} {:<12} {:<12} {:<12} {}", desc, input, expected, got, status);
    }

    if fail > 0 {
        std::process::exit(1);
    }
}
