/// Spelling suggestion test using the shared pipeline.
/// Uses the EXACT same code as the production app.
/// Usage: cargo run --release --bin test_spelling

use nostos_cognio::model::Model;
use nostos_cognio::grammar::swipl_checker::SwiGrammarChecker;
use std::collections::HashMap;
use std::path::PathBuf;

use acatts_rust::spelling_scorer::{generate_spelling_candidates, score_and_rerank};
use language::LanguageSpelling as _;

fn main() {
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    // Try both Mac layout (../contexter-repo) and Windows layout (../../contexter-repo)
    let training = {
        let mac = base.join("../contexter-repo/training-data");
        if mac.exists() { mac } else { base.join("../../contexter-repo/training-data") }
    };

    let onnx_path = training.join("onnx/norbert4_base_int8.onnx");
    let tok_path = training.join("onnx/tokenizer.json");
    println!("Loading NorBERT4...");
    let mut model = Model::load(onnx_path.to_str().unwrap(), tok_path.to_str().unwrap())
        .expect("Failed to load model");
    println!("Loaded. Vocab: {}", model.vocab_size());

    let dict_path = {
        let mac = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
        if mac.exists() { mac } else { base.join("../../rustSpell/mtag-rs/data/fullform_bm.mfst") }
    };
    let grammar_rules_path = {
        let mac = base.join("../syntaxer/grammar_rules.pl");
        if mac.exists() { mac } else { base.join("../../syntaxer/grammar_rules.pl") }
    };
    let syntaxer_dir = {
        let mac = base.join("../syntaxer");
        if mac.exists() { mac } else { base.join("../../syntaxer") }
    };
    let swipl_dll = if cfg!(target_os = "macos") {
        "/Applications/SWI-Prolog.app/Contents/Frameworks/libswipl.dylib"
    } else {
        "C:/Program Files/swipl/bin/libswipl.dll"
    };

    println!("Loading analyzer...");
    let analyzer = mtag::Analyzer::new(dict_path.to_str().unwrap())
        .expect("Failed to load analyzer");

    println!("Loading SWI grammar checker...");
    let mut checker = SwiGrammarChecker::new(
        swipl_dll,
        dict_path.to_str().unwrap(),
        grammar_rules_path.to_str().unwrap(),
        syntaxer_dir.to_str().unwrap(),
    ).expect("Failed to load SWI grammar checker");
    println!("Ready.\n");

    let wf_path = training.join("wordfreq.tsv");
    let wf = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);

    let empty_doc: HashMap<String, u16> = HashMap::new();
    let empty_user: Vec<String> = Vec::new();

    // (sentence, misspelled_word, expected_top1_correction)
    let test_cases = vec![
        ("De skulle få bossller og brus.", "bossller", "boller"),
        ("Fisken hopper i vannetx.", "vannetx", "vannet"),
        ("Hun leser en bokk.", "bokk", "bok"),
        ("Vi skal reise til Bergern.", "bergern", "bergen"),
        ("Katten sitterr på stolen.", "sitterr", "sitter"),
        ("Han spiller fotballl.", "fotballl", "fotball"),
        ("Jeg skal skrierl en bok.", "skrierl", "skrive|skrives"),
        ("Barna fikk boller og bbrus.", "bbrus", "brus"),
        ("Barna fikk gåtterier.", "gåtterier", "godterier|godteri"),
        // First-char wrong + grammar inflection: sjøkken → kjøkkenet
        ("Jeg har mange gryter på sjøkken mitt.", "sjøkken", "kjøkken|kjøkkenet"),
        // Phonetic substitutions (dyslexic patterns)
        ("Det var et gott år.", "gott", "godt"),              // silent d: tt→dt
        ("Jeg vil lare meg norsk.", "lare", "lære"),          // a→æ
        ("Hun herte ikke.", "herte", "hørte"),                // e→ø
    ];

    let mut pass = 0;
    let mut fail = 0;

    for (sentence, misspelled, expected) in &test_cases {
        println!("\n{}", "=".repeat(60));
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
            &language::BokmalLanguage,
        );
        println!("  Phase 1: {} candidates", candidates.len());

        let expected_alts: Vec<&str> = expected.split('|').collect();
        let expected_in_pool = expected_alts.iter().any(|alt| {
            candidates.iter().any(|(c, _)| c == &alt.to_lowercase())
        });
        if expected_in_pool {
            println!("  ✓ expected word IS in candidate pool");
            // Show expected word's position and ortho score
            for alt in &expected_alts {
                if let Some((pos, (_, score))) = candidates.iter().enumerate().find(|(_, (c, _))| c == &alt.to_lowercase()) {
                    println!("    '{}' at Phase1 rank #{} ortho={:.3}", alt, pos + 1, score);
                }
            }
        } else {
            println!("  ✗ expected word NOT in candidate pool!");
        }

        // Phase 2: Score and rerank (same code as app's BERT worker)
        let sentence_lower = sentence.to_lowercase();
        let word_lower = misspelled.to_lowercase();
        let (context_before, context_after) = if let Some(pos) = sentence_lower.find(&word_lower) {
            (sentence_lower[..pos].to_string(), sentence_lower[pos + word_lower.len()..].to_string())
        } else {
            (sentence_lower.clone(), String::new())
        };

        let mut grammar_check = |sentences: &[String]| -> Vec<Vec<nostos_cognio::grammar::types::GrammarError>> {
            sentences.iter().map(|s| checker.check_sentence(s)).collect()
        };

        let results = score_and_rerank(
            &mut model,
            &mut grammar_check,
            &candidates,
            &context_before,
            &context_after,
            sentence,
        );

        println!("  Top 5:");
        for (i, (w, s)) in results.iter().take(5).enumerate() {
            let marker = if expected_alts.contains(&w.as_str()) { " ✓" } else { "" };
            println!("    #{}: '{}' score={:.3}{}", i + 1, w, s, marker);
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

    // Split detection tests
    println!("\n{}", "=".repeat(60));
    println!("=== Split detection tests ===");
    let split_tests = vec![
        ("tilbutikken", Some("til butikken")),
        ("imorgen", None),
        ("pågrunn", Some("på grunn")),
        ("medvilje", Some("med vilje")),
        ("avstand", None),
        ("tilstand", None),
        ("iform", Some("i form")),
        ("tilslutt", None),
        ("frastart", Some("fra start")),
        ("vedsiden", Some("ved siden")),
        ("løpsakte", Some("løp sakte")),
    ];

    for (word, expected_split) in &split_tests {
        let result = acatts_rust::spelling_scorer::try_split_function_word(word, &analyzer, &language::BokmalLanguage);
        let ok = match (expected_split, &result) {
            (Some(exp), Some(got)) => got == exp,
            (None, None) => true,
            _ => false,
        };
        if ok { pass += 1; } else { fail += 1; }
        println!("  {}  '{}' → '{}'",
            if ok { "PASS" } else { "FAIL" },
            word, result.as_deref().unwrap_or("(no split)"));
    }

    println!("\n{}", "=".repeat(60));
    println!("Results: {}/{} passed", pass, pass + fail);
    if fail > 0 {
        std::process::exit(1);
    }
}
