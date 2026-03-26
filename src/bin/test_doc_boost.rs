/// Test document-frequency and user-dictionary boosting for suggestions.
/// Self-contained: reimplements compute_boost logic for testing.

use std::collections::HashMap;

/// Same logic as compute_boost() in main.rs — kept in sync.
fn compute_boost(
    word: &str,
    doc_word_counts: &HashMap<String, u16>,
    user_words: &[String],
    wordfreq: Option<&HashMap<String, u64>>,
) -> f32 {
    let lower = word.to_lowercase();
    const COMMON_THRESHOLD: u64 = 40_000;
    if wordfreq.and_then(|wf| wf.get(&lower)).map_or(false, |&f| f >= COMMON_THRESHOLD) {
        return 1.0;
    }
    let in_doc = doc_word_counts.get(&lower).copied().unwrap_or(0) >= 2;
    let in_user = user_words.iter().any(|uw| uw.eq_ignore_ascii_case(&lower));
    match (in_doc, in_user) {
        (true, true)   => 1.6,
        (false, true)  => 1.3,
        (true, false)  => 1.25,
        (false, false) => 1.0,
    }
}

fn build_doc_word_counts(text: &str) -> HashMap<String, u16> {
    let mut counts = HashMap::new();
    for word in text.split(|c: char| !c.is_alphanumeric() && c != '-') {
        if word.len() < 2 { continue; }
        *counts.entry(word.to_lowercase()).or_insert(0u16) += 1;
    }
    counts
}

fn main() {
    let mut passed = 0u32;
    let mut failed = 0u32;

    // --- Build test data ---
    let mut doc_counts: HashMap<String, u16> = HashMap::new();
    doc_counts.insert("fotball".to_string(), 5);
    doc_counts.insert("teknologi".to_string(), 3);
    doc_counts.insert("cognio".to_string(), 4);
    doc_counts.insert("og".to_string(), 25);
    doc_counts.insert("er".to_string(), 30);
    doc_counts.insert("skole".to_string(), 2);
    doc_counts.insert("engang".to_string(), 1);

    let mut wf: HashMap<String, u64> = HashMap::new();
    wf.insert("og".to_string(), 422_137);
    wf.insert("er".to_string(), 210_535);
    wf.insert("for".to_string(), 267_874);
    wf.insert("fotball".to_string(), 153);
    wf.insert("teknologi".to_string(), 1107);
    wf.insert("skole".to_string(), 2147);
    wf.insert("engang".to_string(), 500);

    let user_words: Vec<String> = vec!["nevrale".into(), "fotball".into(), "og".into()];

    // === compute_boost tests ===
    println!("=== compute_boost tests ===");

    // Common words: no boost
    check("common 'og' (freq 422K, 25x in doc) → 1.0",
        compute_boost("og", &doc_counts, &[], Some(&wf)), 1.0, &mut passed, &mut failed);
    check("common 'er' (freq 210K, 30x in doc) → 1.0",
        compute_boost("er", &doc_counts, &[], Some(&wf)), 1.0, &mut passed, &mut failed);

    // Domain words in doc
    check("'fotball' (freq 153, 5x in doc) → 1.25",
        compute_boost("fotball", &doc_counts, &[], Some(&wf)), 1.25, &mut passed, &mut failed);
    check("'teknologi' (freq 1107, 3x in doc) → 1.25",
        compute_boost("teknologi", &doc_counts, &[], Some(&wf)), 1.25, &mut passed, &mut failed);
    check("'skole' (freq 2147, 2x in doc) → 1.25",
        compute_boost("skole", &doc_counts, &[], Some(&wf)), 1.25, &mut passed, &mut failed);

    // Below threshold (1x)
    check("'engang' (1x in doc) → 1.0 (below threshold)",
        compute_boost("engang", &doc_counts, &[], Some(&wf)), 1.0, &mut passed, &mut failed);

    // Not in corpus at all
    check("'cognio' (not in corpus, 4x in doc) → 1.25",
        compute_boost("cognio", &doc_counts, &[], Some(&wf)), 1.25, &mut passed, &mut failed);

    // Not in doc, not in user dict
    check("'hest' (not in doc) → 1.0",
        compute_boost("hest", &doc_counts, &[], Some(&wf)), 1.0, &mut passed, &mut failed);

    // User dict only
    check("'nevrale' (user dict, not in doc) → 1.3",
        compute_boost("nevrale", &doc_counts, &user_words, Some(&wf)), 1.3, &mut passed, &mut failed);

    // Both doc + user dict
    check("'fotball' (doc + user dict) → 1.6",
        compute_boost("fotball", &doc_counts, &user_words, Some(&wf)), 1.6, &mut passed, &mut failed);

    // Case insensitive
    check("'Fotball' (uppercase) → 1.25",
        compute_boost("Fotball", &doc_counts, &[], Some(&wf)), 1.25, &mut passed, &mut failed);

    // Common word + user dict → still no boost
    check("'og' (common, in user dict) → 1.0",
        compute_boost("og", &doc_counts, &user_words, Some(&wf)), 1.0, &mut passed, &mut failed);

    // No wordfreq available → doc words still boosted
    check("'fotball' (no wordfreq, 5x in doc) → 1.25",
        compute_boost("fotball", &doc_counts, &[], None), 1.25, &mut passed, &mut failed);

    // === build_doc_word_counts tests ===
    println!("\n=== build_doc_word_counts tests ===");

    let text = "Fotball er en spennende idrett. Fotball er veldig populært. Jeg liker fotball og fotball-VM.";
    let counts = build_doc_word_counts(text);

    // "fotball" appears 3 times standalone; "fotball-vm" is kept as hyphenated token
    check_eq("'fotball' count", counts.get("fotball").copied().unwrap_or(0), 3, &mut passed, &mut failed);
    check_eq("'er' count", counts.get("er").copied().unwrap_or(0), 2, &mut passed, &mut failed);
    check_eq("'fotball-vm' count (hyphenated)", counts.get("fotball-vm").copied().unwrap_or(0), 1, &mut passed, &mut failed);
    check_eq("'en' count", counts.get("en").copied().unwrap_or(0), 1, &mut passed, &mut failed);

    // === Boost applied to ranking ===
    println!("\n=== Ranking simulation ===");

    // Simulate: 3 candidates with similar BERT scores, one is in document
    let candidates = vec![
        ("fotografi".to_string(), 15.0f32),
        ("fotball".to_string(), 14.5),
        ("foto".to_string(), 14.8),
    ];
    let doc = build_doc_word_counts("Jeg liker fotball. Fotball er gøy. Vi spiller fotball.");
    let mut boosted: Vec<(String, f32)> = candidates.iter().map(|(w, s)| {
        let b = compute_boost(w, &doc, &[], Some(&wf));
        (w.clone(), s * b)
    }).collect();
    boosted.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

    let top = &boosted[0].0;
    if top == "fotball" {
        println!("  PASS: 'fotball' ranked #1 after boost (score {:.2})", boosted[0].1);
        passed += 1;
    } else {
        println!("  FAIL: expected 'fotball' #1 but got '{}' ({:.2})", top, boosted[0].1);
        failed += 1;
    }

    // Without boost, fotografi would be #1
    let mut unboosted = candidates.clone();
    unboosted.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
    if unboosted[0].0 == "fotografi" {
        println!("  PASS: without boost, 'fotografi' is #1 (control)");
        passed += 1;
    } else {
        println!("  FAIL: control — expected 'fotografi' #1 without boost");
        failed += 1;
    }

    println!("\n=== {}/{} tests passed ===", passed, passed + failed);
    if failed > 0 {
        std::process::exit(1);
    }
}

fn check(name: &str, got: f32, expected: f32, passed: &mut u32, failed: &mut u32) {
    if (got - expected).abs() < 0.01 {
        println!("  PASS: {}", name);
        *passed += 1;
    } else {
        println!("  FAIL: {} — got {:.2}, expected {:.2}", name, got, expected);
        *failed += 1;
    }
}

fn check_eq(name: &str, got: u16, expected: u16, passed: &mut u32, failed: &mut u32) {
    if got == expected {
        println!("  PASS: {} = {}", name, got);
        *passed += 1;
    } else {
        println!("  FAIL: {} — got {}, expected {}", name, got, expected);
        *failed += 1;
    }
}
