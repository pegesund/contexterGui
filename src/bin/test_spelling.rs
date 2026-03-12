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
    checker: &mut nostos_cognio::grammar::swipl_checker::SwiGrammarChecker,
    sentence: &str,
    word: &str,
    expected: &str,
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

    // Source 4: Fuzzy on truncated word (strip 1-2 trailing chars then fuzzy)
    // Catches "skrierl" → strip 'l' → fuzzy("skrier",2) → finds "skrive", "skriver"
    for strip in 1..=2u32 {
        let chars: Vec<char> = word_lower.chars().collect();
        if chars.len() <= 3 + strip as usize { continue; }
        let truncated: String = chars[..chars.len() - strip as usize].iter().collect();
        let fuzzy = checker.fuzzy_lookup(&truncated, 2);
        for (w, dist) in fuzzy {
            let wl = w.to_lowercase();
            edit_distances.entry(wl.clone()).or_insert(dist + strip);
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
    // Step 1: Score all candidates by orthographic similarity, take top-N
    let mut ortho_scored: Vec<(String, f32)> = Vec::new();
    for w in &candidates {
        let w_tri = trigrams(w);
        let common = word_trigrams.iter().filter(|t| w_tri.contains(t)).count();
        let max_t = word_trigrams.len().max(w_tri.len()).max(1);
        let trigram_sim = common as f32 / max_t as f32;

        let prefix_len = word_lower.chars().zip(w.chars())
            .take_while(|(a, b)| a == b).count();
        let max_len = word_lower.chars().count().max(w.chars().count()).max(1);
        let prefix_sim = prefix_len as f32 / max_len as f32;

        let edit_sim = match edit_distances.get(w.as_str()) {
            Some(1) => 0.85,
            Some(2) => 0.65,
            Some(3) => 0.45,
            _ => 0.0,
        };
        let ortho_sim = trigram_sim.max(edit_sim).max(prefix_sim);
        ortho_scored.push((w.clone(), ortho_sim));
    }
    ortho_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap());
    println!("  Masked: '{}'", masked);
    println!("  Candidates: {}", candidates.len());
    let expected_lower = expected.to_lowercase();
    if candidates.contains(&expected_lower) {
        println!("  ✓ '{}' IS in candidate pool", expected);
    } else {
        println!("  ✗ '{}' NOT in candidate pool!", expected);
    }

    // Step 2: Score with BERT × ortho_sim
    let mut scored: Vec<(String, f32)> = Vec::new();
    if let Ok((logits, _ms)) = model.single_forward(&masked) {
        for w in &candidates {
            let bert_score = if let Ok(enc) = model.tokenizer.encode(w.as_str(), false) {
                let ids = enc.get_ids();
                if ids.is_empty() {
                    0.0
                } else {
                    let raw = logits.get(ids[0] as usize).copied().unwrap_or(0.0);
                    if ids.len() > 1 { raw * 0.9 } else { raw }
                }
            } else {
                0.0
            };

            let w_tri = trigrams(w);
            let common = word_trigrams.iter().filter(|t| w_tri.contains(t)).count();
            let max_t = word_trigrams.len().max(w_tri.len()).max(1);
            let trigram_sim = common as f32 / max_t as f32;

            let prefix_len = word_lower.chars().zip(w.chars())
                .take_while(|(a, b)| a == b).count();
            let max_len = word_lower.chars().count().max(w.chars().count()).max(1);
            let prefix_sim = prefix_len as f32 / max_len as f32;

            let edit_sim = match edit_distances.get(w.as_str()) {
                Some(1) => 0.85,
                Some(2) => 0.65,
                Some(3) => 0.45,
                _ => 0.0,
            };
            let ortho_sim = trigram_sim.max(edit_sim).max(prefix_sim);

            let score = bert_score.max(0.0) * ortho_sim;
            scored.push((w.clone(), score));
        }
    }

    scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap());

    // Debug: show where expected word ranks
    if let Some(pos) = scored.iter().position(|(w, _)| *w == expected.to_lowercase()) {
        let (_, s) = &scored[pos];
        println!("  DEBUG: '{}' at rank #{} score={:.3}", expected, pos + 1, s);
    } else {
        println!("  DEBUG: '{}' not in scored list!", expected);
    }

    // Grammar filter (same as app: pick first that passes)
    let mut passing: Vec<(String, f32)> = Vec::new();
    for (candidate, score) in scored.iter().take(25) {
        if !checker.has_word(candidate) { continue; }
        let corrected = sentence.to_lowercase().replacen(&word_lower, candidate, 1);
        let errors = checker.check_sentence(&corrected);
        println!("  grammar: '{}' score={:.3} → {} errors", candidate, score, errors.len());
        if errors.is_empty() {
            passing.push((candidate.clone(), *score));
        }
    }

    if !passing.is_empty() { passing } else { scored }
}

/// Same logic as main.rs try_split_function_word
fn try_split_function_word(word: &str, checker: &nostos_cognio::grammar::swipl_checker::SwiGrammarChecker) -> Option<String> {
    const FUNCTION_WORDS: &[&str] = &[
        "gjennom", "mellom", "under", "etter", "langs", "rundt",
        "foran", "bortover", "innover", "utover",
        "forbi", "siden", "etter", "blant",
        "over", "inne", "borte",
        "uten", "utenfor", "innenfor",
        "med", "mot", "ved", "hos", "fra",
        "for", "som", "men",
        "til", "per", "via",
        "på", "av", "om",
        "en", "et", "ei",
        "og", "at",
        "i",
    ];
    let lower = word.to_lowercase();
    for prefix in FUNCTION_WORDS {
        if lower.len() <= prefix.len() + 1 { continue; }
        if !lower.starts_with(prefix) { continue; }
        let remainder = &lower[prefix.len()..];
        if remainder.len() < 2 { continue; }
        if checker.has_word(remainder) {
            return Some(format!("{} {}", prefix, remainder));
        }
    }
    // General split: both parts ≥3 chars, both in dictionary
    let chars: Vec<char> = lower.chars().collect();
    let mut best_split: Option<(String, usize)> = None;
    for split_at in 3..=(chars.len().saturating_sub(3)) {
        let left: String = chars[..split_at].iter().collect();
        let right: String = chars[split_at..].iter().collect();
        if checker.has_word(&left) && checker.has_word(&right) {
            let balance = left.len().min(right.len());
            if best_split.as_ref().map(|(_, b)| balance > *b).unwrap_or(true) {
                best_split = Some((format!("{} {}", left, right), balance));
            }
        }
    }
    best_split.map(|(s, _)| s)
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
    let grammar_rules_path = base.join("../../syntaxer/grammar_rules.pl");
    let syntaxer_dir = base.join("../../syntaxer");
    let swipl_dll = "C:/Program Files/swipl/bin/libswipl.dll";
    let mut checker = nostos_cognio::grammar::swipl_checker::SwiGrammarChecker::new(
        swipl_dll,
        dict_path.to_str().unwrap(),
        grammar_rules_path.to_str().unwrap(),
        syntaxer_dir.to_str().unwrap(),
    )
    .expect("Failed to load SWI grammar checker");

    // Debug: check readings for problematic words
    for w in &["skrier", "skrive", "skriver", "skrie", "strier"] {
        let token = checker.analyze_word(w);
        println!("  readings('{}'): {:?}", w, token);
    }

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
        ("Jeg skal skrierl.", "skrierl", "skrive|skrives"),
    ];

    let mut pass = 0;
    let mut fail = 0;

    for (sentence, misspelled, expected) in &test_cases {
        println!("\n{}", "=".repeat(60));
        println!("Test: '{}' → expected '{}'", misspelled, expected);
        println!("Sentence: '{}'", sentence);

        let results = score_candidates(&mut model, &mut checker, sentence, misspelled, expected);

        let expected_alts: Vec<&str> = expected.split('|').collect();
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

    // === Split detection tests (function word + remainder) ===
    println!("\n{}", "=".repeat(60));
    println!("=== Split detection tests ===");

    // (word, sentence_context, expected_split)
    // sentence_context is used for grammar validation of the split
    let split_tests: Vec<(&str, &str, &str)> = vec![
        ("tilbutikken", "Han gikk tilbutikken.", "til butikken"),
        ("imorgen", "Vi reiser imorgen.", ""),         // in dictionary
        ("pågrunn", "Det skjedde pågrunn av regnet.", "på grunn"),
        ("medvilje", "Han gjorde det medvilje.", "med vilje"),
        ("avstand", "Hold avstand.", ""),       // legitimate compound
        ("tilstand", "En god tilstand.", ""),   // legitimate compound
        ("iform", "Han er iform.", "i form"),
        ("tilslutt", "Vi kom tilslutt.", ""),   // in dictionary
        ("frastart", "Vi var med frastart.", "fra start"),
        ("vedsiden", "Hun stod vedsiden.", "ved siden"),
        ("løpsakte", "Hun løpsakte gjennom parken.", "løp sakte"),
    ];

    for (word, sentence, expected_split) in &split_tests {
        // In the real app, split is only tried for unknown words
        let result = if checker.has_word(word) {
            None // word exists in dictionary — no split needed
        } else {
            let split = try_split_function_word(word, &checker);
            // Grammar-validate: check sentence with split applied
            if let Some(ref s) = split {
                let corrected = sentence.to_lowercase().replacen(&word.to_lowercase(), s, 1);
                let errors = checker.check_sentence(&corrected);
                if !errors.is_empty() {
                    println!("  (grammar rejected: '{}' → {} errors)", corrected, errors.len());
                    None
                } else {
                    split
                }
            } else {
                split
            }
        };
        let result_str = result.as_deref().unwrap_or("");
        let ok = result_str == *expected_split;
        if ok {
            println!("  PASS  '{}' → '{}'", word, if expected_split.is_empty() { "(no split)" } else { expected_split });
            pass += 1;
        } else {
            println!("  FAIL  '{}' → got '{}', expected '{}'", word,
                if result_str.is_empty() { "(no split)" } else { result_str },
                if expected_split.is_empty() { "(no split)" } else { expected_split });
            fail += 1;
        }
    }

    println!("\n{}", "=".repeat(60));
    println!("Results: {}/{} passed", pass, pass + fail);
    if fail > 0 {
        std::process::exit(1);
    }
}
