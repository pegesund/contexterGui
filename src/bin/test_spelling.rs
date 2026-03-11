/// Console test for spelling suggestions with BERT context.
/// Tests that the same scoring logic used in the app picks the right correction.
/// Usage: cargo run --release --bin test_spelling

use nostos_cognio::model::Model;
use std::collections::{HashMap, HashSet};
use std::path::PathBuf;

fn trigrams(word: &str) -> Vec<String> {
    let chars: Vec<char> = word.chars().collect();
    if chars.len() < 3 {
        return vec![word.to_string()];
    }
    (0..chars.len() - 2)
        .map(|i| chars[i..i + 3].iter().collect())
        .collect()
}

/// Score candidates using the same algorithm as main.rs trigram_suggestions()
fn score_candidates(
    model: &mut Model,
    checker: &mut nostos_cognio::grammar::GrammarChecker,
    sentence: &str,
    word: &str,
) -> Vec<(String, f32)> {
    let word_lower = word.to_lowercase();
    let word_trigrams = trigrams(&word_lower);
    let word_first = word_lower.chars().next().unwrap_or(' ');

    // Collect candidates + edit distances
    let mut candidates: Vec<String> = Vec::new();
    let mut seen = HashSet::new();
    let mut edit_distances: HashMap<String, u32> = HashMap::new();
    let mut prefix_matches: HashSet<String> = HashSet::new();

    // Source 1: Fuzzy (distance 2)
    let fuzzy = checker.fuzzy_lookup(&word_lower, 2);
    for (w, dist) in fuzzy {
        let wl = w.to_lowercase();
        edit_distances.insert(wl.clone(), dist);
        if wl != word_lower && wl.len() >= 2 && seen.insert(wl.clone()) {
            let w_tri = trigrams(&wl);
            let common = word_trigrams.iter().filter(|t| w_tri.contains(t)).count();
            if common > 0 || wl.chars().next().unwrap_or(' ') == word_first {
                candidates.push(wl);
            }
        }
    }

    // Source 2: Prefix lookup
    for w in checker.prefix_lookup(&word_lower, 20) {
        let wl = w.to_lowercase();
        let extra = wl.len() as i32 - word_lower.len() as i32;
        if extra >= 1 && extra <= 3 {
            prefix_matches.insert(wl.clone());
            edit_distances.entry(wl.clone()).or_insert(extra as u32);
            if wl != word_lower && wl.len() >= 2 && seen.insert(wl.clone()) {
                candidates.push(wl);
            }
        }
    }

    // Source 3: Prefix with last char removed
    if word_lower.len() >= 3 {
        let shorter = &word_lower[..word_lower.len() - 1];
        for w in checker.prefix_lookup(shorter, 20) {
            let wl = w.to_lowercase();
            let diff = (wl.len() as i32 - word_lower.len() as i32).unsigned_abs() + 1;
            edit_distances.entry(wl.clone()).or_insert(diff);
            if wl != word_lower && wl.len() >= 2 && seen.insert(wl.clone()) {
                let w_tri = trigrams(&wl);
                let common = word_trigrams.iter().filter(|t| w_tri.contains(t)).count();
                if common > 0 || wl.chars().next().unwrap_or(' ') == word_first {
                    candidates.push(wl);
                }
            }
        }
    }

    // Build masked context (same as app: glued, trim_end before)
    let sentence_lower = sentence.to_lowercase();
    let masked = if let Some(pos) = sentence_lower.find(&word_lower) {
        let before = &sentence[..pos];
        let after = &sentence[pos + word_lower.len()..];
        format!("{}<mask>{}", before.trim_end(), after)
    } else {
        format!("{} <mask>", sentence)
    };
    println!("  Masked: '{}'", masked);
    println!("  Candidates: {}", candidates.len());

    // Score with BERT
    let mut scored: Vec<(String, f32)> = Vec::new();
    if let Ok((logits, _ms)) = model.single_forward(&masked) {
        for w in &candidates {
            let bert_score = if let Ok(enc) = model.tokenizer.encode(w.as_str(), false) {
                let ids = enc.get_ids();
                if ids.is_empty() {
                    0.0
                } else {
                    let raw = logits.get(ids[0] as usize).copied().unwrap_or(0.0);
                    // Discount multi-token candidates (same as app)
                    if ids.len() > 1 { raw * 0.9 } else { raw }
                }
            } else {
                0.0
            };

            let w_trigrams = trigrams(w);
            let common = word_trigrams.iter().filter(|t| w_trigrams.contains(t)).count();
            let max_t = word_trigrams.len().max(w_trigrams.len()).max(1);
            let trigram_sim = common as f32 / max_t as f32;

            let edit_sim = match edit_distances.get(w) {
                Some(1) => 0.85,
                Some(2) => 0.65,
                _ => 0.0,
            };
            let ortho_sim = trigram_sim.max(edit_sim);

            let prefix_bonus = if prefix_matches.contains(w) {
                let extra = w.len() as f32 - word_lower.len() as f32;
                if extra <= 1.0 { 1.5 } else if extra <= 2.0 { 1.2 } else { 1.1 }
            } else {
                1.0
            };

            let score = bert_score.max(0.0) * ortho_sim * prefix_bonus;
            scored.push((w.clone(), score));
        }
    }

    scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap());

    // Grammar filter (same as app: pick first that passes)
    let mut passing: Vec<(String, f32)> = Vec::new();
    for (candidate, score) in scored.iter().take(8) {
        if !checker.has_word(candidate) { continue; }
        let corrected = sentence.to_lowercase().replacen(&word_lower, candidate, 1);
        let errors = checker.check_sentence(&corrected);
        if errors.is_empty() {
            passing.push((candidate.clone(), *score));
        }
    }

    if !passing.is_empty() { passing } else { scored }
}

fn main() {
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let training = base.join("../../contexter-repo/training-data");

    let onnx_path = training.join("onnx/norbert4_base_int8.onnx");
    let tok_path = training.join("onnx/tokenizer.json");
    println!("Loading NorBERT4...");
    let mut model = Model::load(onnx_path.to_str().unwrap(), tok_path.to_str().unwrap())
        .expect("Failed to load model");
    println!("Loaded. Vocab: {}", model.vocab_size());

    let dict_path = base.join("../../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let mut checker = nostos_cognio::grammar::GrammarChecker::new(
        dict_path.to_str().unwrap(),
        "",
    )
    .expect("Failed to load dictionary");

    // (sentence, misspelled_word, expected_top1_correction)
    // Tests verify the BERT contextual scoring algorithm picks the right word.
    // Focus: typos (extra/wrong chars) where correct word is edit distance 1-2.
    let test_cases = vec![
        ("De skulle få bossller og brus.", "bossller", "boller"),
        ("Fisken hopper i vannetx.", "vannetx", "vannet"),
        ("Hun leser en bokk.", "bokk", "bok"),
        ("Vi skal reise til Bergern.", "bergern", "bergen"),
        ("Katten sitterr på stolen.", "sitterr", "sitter"),
        ("Han spiller fotballl.", "fotballl", "fotball"),
    ];

    let mut pass = 0;
    let mut fail = 0;

    for (sentence, misspelled, expected) in &test_cases {
        println!("\n{}", "=".repeat(60));
        println!("Test: '{}' → expected '{}'", misspelled, expected);
        println!("Sentence: '{}'", sentence);

        let results = score_candidates(&mut model, &mut checker, sentence, misspelled);

        println!("  Top 5:");
        for (i, (w, s)) in results.iter().take(5).enumerate() {
            let marker = if w == expected { " ✓" } else { "" };
            println!("    #{}: '{}' score={:.3}{}", i + 1, w, s, marker);
        }

        if let Some((top, _)) = results.first() {
            if top == expected {
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

    println!("\n{}", "=".repeat(60));
    println!("Results: {}/{} passed", pass, pass + fail);
    if fail > 0 {
        std::process::exit(1);
    }
}
