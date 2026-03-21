use std::path::PathBuf;
use std::time::Instant;

fn main() {
    let swipl_path = "/Applications/SWI-Prolog.app/Contents/Frameworks/libswipl.dylib";
    let dict_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let grammar_rules = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../syntaxer/grammar_rules.pl");
    let syntaxer_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../syntaxer");

    eprintln!("Loading SWI-Prolog...");
    let mut checker = nostos_cognio::grammar::swipl_checker::SwiGrammarChecker::new(
        swipl_path,
        dict_path.to_str().unwrap(),
        grammar_rules.to_str().unwrap(),
        syntaxer_dir.to_str().unwrap(),
    ).expect("Failed to load SWI-Prolog");
    eprintln!("Ready.\n");

    let sentences = vec![
        "Fotball er en morsom sport.",
        "Fotball er en morsom spor.",
        "Fotball er en morsom spill.",
        "Fotball er en morsom sportsgren.",
        "Fotball er en morsom idrett.",
        "Fotball er en morsom aktivitet.",
        "Fotball er en morsom hobby.",
        "Fotball er en morsom sak.",
        "Fotball er en morsom serie.",
        "Fotball er en morsom sporte.",
        "Fotball er en morsom spilletid.",
        "Fotball er en morsom spiller.",
        "Fotball er en morsom si.",
        "Fotball er en morsom sel.",
        "Fotball er en morsom som.",
    ];

    // Warmup
    let _ = checker.check_sentence("Hei dette er en test.");

    // Test: is it "spor" specifically or any word with many readings?
    let timing_tests = vec![
        "Morsom spor.",      // slow
        "Morsom sport.",     // fast
        "Morsom ball.",      // ?
        "Morsom lag.",       // ? (neuter, like spor)
        "Morsom bord.",      // ? (neuter, but why fast?)
        "Morsomt bord.",     // correct neuter
        "Et morsomt bord.",  // correct with article
        "Morsom hus.",       // ? (neuter)
        "God spor.",         // ? (different adj)
        "Fin spor.",         // ?
        "Stor spor.",        // ?
    ];
    println!("=== Is it 'spor' or any gender mismatch? ===");
    for s in &timing_tests {
        let t = Instant::now();
        let errors = checker.check_sentence(s);
        let ms = t.elapsed().as_micros() as f64 / 1000.0;
        println!("{:>8.1}ms  {} errors  '{}'", ms, errors.len(), s);
    }
    println!();

    // Test more sentences to find the slow pattern
    let slow_tests = vec![
        "Fotball er en morsom spor.",
        "Fotball er en morsom spil.",
        "Han er en god spor.",
        "Det er et morsomt spor.",
        "Jeg liker å spor.",
        "Fotball er morsom spor.",
        "En morsom spor.",
        "Morsom spor.",
    ];
    println!("\n=== Slow pattern analysis ===");
    for s in &slow_tests {
        let t = Instant::now();
        let errors = checker.check_sentence(s);
        let ms = t.elapsed().as_micros() as f64 / 1000.0;
        let err_desc: Vec<String> = errors.iter().map(|e| format!("{}:{}", e.rule_name, e.word)).collect();
        println!("{:>8.1}ms  {} errors  '{}'  rules: [{}]", ms, errors.len(), s, err_desc.join(", "));
    }

    // Test has_error (quick check — stops at first error)
    println!("\n=== has_error (should be faster) ===");
    for s in &timing_tests {
        let t = Instant::now();
        let has_err = checker.has_error(s);
        let ms = t.elapsed().as_micros() as f64 / 1000.0;
        println!("{:>8.1}ms  has_error={}  '{}'", ms, has_err, s);
    }

    println!("\n=== Main test ===");

    // Batch timing
    let t_total = Instant::now();
    for s in &sentences {
        let t = Instant::now();
        let errors = checker.check_sentence(s);
        let ms = t.elapsed().as_micros() as f64 / 1000.0;
        println!("{:>8.1}ms  {} errors  '{}'", ms, errors.len(), s);
    }
    let total_ms = t_total.elapsed().as_millis();
    println!("\nTotal: {}ms for {} sentences ({:.1}ms avg)",
        total_ms, sentences.len(), total_ms as f64 / sentences.len() as f64);
}
