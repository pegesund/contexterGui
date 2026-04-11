/// Grammar test runner using SWI-Prolog checker — ENGLISH.
/// Runs all test sets in test_data_en and reports FP/FN.
/// Usage: cargo run --bin test_grammar_en [set_num ...]

use nostos_cognio::grammar::swipl_checker::SwiGrammarChecker;
use std::fs;
use std::time::Instant;

#[derive(serde::Deserialize)]
struct TestCase {
    sentence: String,
    is_correct: bool,
    #[allow(dead_code)]
    error_type: String,
    #[serde(default)]
    #[allow(dead_code)]
    explanation: String,
}

fn main() {
    let base = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let dict_path = base.join("../rustSpell/mtag-rs/data/fullform_en.mfst");
    let grammar_path = base.join("../english/grammar_rules.pl");
    let english_dir = base.join("../english");

    let swipl_path = find_swipl();
    eprintln!("SWI-Prolog: {}", swipl_path);
    eprintln!("Grammar rules: {}", grammar_path.display());
    eprintln!("Dictionary: {}", dict_path.display());

    let t = Instant::now();
    let mut checker = SwiGrammarChecker::new(
        &swipl_path,
        dict_path.to_str().unwrap(),
        grammar_path.to_str().unwrap(),
        english_dir.to_str().unwrap(),
    ).expect("Failed to create SWI checker");
    eprintln!("Checker created in {:?}", t.elapsed());

    let test_dir = base.join("test_data_en");
    let args: Vec<String> = std::env::args().collect();
    let sets: Vec<usize> = if args.len() > 1 {
        args[1..].iter().filter_map(|a| a.parse().ok()).collect()
    } else {
        (1..=39).collect()
    };

    let mut total_sentences = 0;
    let mut total_fp = 0;
    let mut total_fn = 0;
    let mut total_time = std::time::Duration::ZERO;
    let mut fn_by_type: std::collections::HashMap<String, usize> = std::collections::HashMap::new();

    for set_num in &sets {
        let path = test_dir.join(format!("test_set_{}.json", set_num));
        let data = match fs::read_to_string(&path) {
            Ok(d) => d,
            Err(_) => { continue; }
        };
        let tests: Vec<TestCase> = match serde_json::from_str(&data) {
            Ok(t) => t,
            Err(e) => { eprintln!("Failed to parse {}: {}", path.display(), e); continue; }
        };

        let mut set_fp = 0;
        let mut set_fn = 0;
        let mut set_time = std::time::Duration::ZERO;

        for test in &tests {
            let t = Instant::now();
            let errors = checker.check_sentence(&test.sentence);
            let elapsed = t.elapsed();
            set_time += elapsed;
            total_sentences += 1;

            let has_error = !errors.is_empty();
            if test.is_correct && has_error {
                set_fp += 1;
                if set_fp <= 5 {
                    eprintln!("  FP set {}: '{}' flagged: {:?}", set_num, test.sentence,
                        errors.iter().map(|e| &e.rule_name).collect::<Vec<_>>());
                }
            } else if !test.is_correct && !has_error {
                set_fn += 1;
                *fn_by_type.entry(test.error_type.clone()).or_insert(0) += 1;
                if set_fn <= 3 {
                    eprintln!("  FN set {}: '{}' (expected: {})", set_num, test.sentence, test.error_type);
                }
            }
        }

        total_fp += set_fp;
        total_fn += set_fn;
        total_time += set_time;
        let status = if set_fp == 0 && set_fn == 0 { "PASS" } else { "FAIL" };
        let avg_ms = if tests.is_empty() { 0.0 } else { set_time.as_secs_f64() * 1000.0 / tests.len() as f64 };
        println!("Set {:2}: {:5} sentences, FP={}, FN={}, {:.1}ms/sent  [{}]",
            set_num, tests.len(), set_fp, set_fn, avg_ms, status);
    }

    println!("\n=== TOTAL: {} sentences, FP={}, FN={}, {:.1}s ===",
        total_sentences, total_fp, total_fn, total_time.as_secs_f64());
    if total_fp == 0 && total_fn == 0 {
        println!("ALL TESTS PASSED");
    } else {
        println!("\nFN by error type:");
        let mut fn_vec: Vec<_> = fn_by_type.iter().collect();
        fn_vec.sort_by(|a, b| b.1.cmp(a.1));
        for (et, count) in &fn_vec {
            println!("  {:5} {}", count, et);
        }
        let precision = if total_sentences > 0 { 100.0 * (1.0 - total_fp as f64 / total_sentences as f64) } else { 0.0 };
        let total_errors = total_sentences - (total_sentences - total_fp - total_fn); // approximate
        println!("Precision: {:.1}%, FP rate: {:.2}%, FN rate: {:.2}%",
            precision,
            100.0 * total_fp as f64 / total_sentences as f64,
            100.0 * total_fn as f64 / total_sentences as f64);
    }
}

fn find_swipl() -> String {
    for path in &[
        "/opt/homebrew/lib/swipl/lib/arm64-darwin/libswipl.dylib",
        "/opt/homebrew/lib/libswipl.dylib",
        "/usr/local/lib/swipl/lib/x86_64-darwin/libswipl.dylib",
        "/usr/local/lib/libswipl.dylib",
        "/Applications/SWI-Prolog.app/Contents/Frameworks/libswipl.dylib",
    ] {
        if std::path::Path::new(path).exists() {
            return path.to_string();
        }
    }
    if let Ok(output) = std::process::Command::new("mdfind").arg("libswipl.dylib").output() {
        let paths = String::from_utf8_lossy(&output.stdout);
        if let Some(p) = paths.lines().next() {
            if !p.is_empty() { return p.to_string(); }
        }
    }
    panic!("Could not find libswipl.dylib — install SWI-Prolog");
}
