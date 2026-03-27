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
    let char_count = word_lower.chars().count();
    if char_count >= 3 {
        let end_byte = word_lower.char_indices().rev().next().map(|(i, _)| i).unwrap_or(0);
        let shorter = &word_lower[..end_byte];
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

    // Score all candidates by orthographic similarity only
    // First-token BERT is unreliable — sentence-level BERT re-ranking handles context
    let sentence_lower = sentence.to_lowercase();
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
        let mut ortho_sim = trigram_sim.max(edit_sim).max(prefix_sim);
        // First-char bonus: misspellings usually preserve the initial letter
        if w.chars().next() == Some(word_first) {
            ortho_sim += 0.15;
        }
        ortho_scored.push((w.clone(), ortho_sim));
    }
    ortho_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap());
    println!("  Candidates: {}", candidates.len());
    let expected_lower = expected.to_lowercase();
    let expected_alts: Vec<&str> = expected.split('|').collect();
    let expected_in_pool = expected_alts.iter().any(|alt| candidates.contains(&alt.to_lowercase()));
    if expected_in_pool {
        println!("  ✓ '{}' IS in candidate pool", expected);
    } else {
        println!("  ✗ '{}' NOT in candidate pool!", expected);
    }

    // Debug: show where expected word ranks in ortho
    for alt in &expected_alts {
        if let Some(pos) = ortho_scored.iter().position(|(w, _)| *w == alt.to_lowercase()) {
            let (_, s) = &ortho_scored[pos];
            println!("  DEBUG ortho: '{}' at rank #{} score={:.3}", alt, pos + 1, s);
        }
    }

    // Grammar filter: check candidates, use grammar suggestions when available
    // "kjøkken" → grammar says error with suggestion "kjøkkenet" → use that
    let mut passing: Vec<(String, f32)> = Vec::new();
    let mut seen_passing = HashSet::new();
    for (candidate, score) in ortho_scored.iter().take(100) {
        if !word_lower.contains('-') && candidate.contains('-') { continue; }
        if !checker.has_word(candidate) { continue; }
        let corrected = sentence.to_lowercase().replacen(&word_lower, candidate, 1);
        let errors = checker.check_sentence(&corrected);
        println!("  grammar: '{}' score={:.3} → {} errors", candidate, score, errors.len());
        if errors.is_empty() {
            if seen_passing.insert(candidate.clone()) {
                passing.push((candidate.clone(), *score));
            }
        } else {
            // Use grammar suggestion — it's the grammatically correct form of a valid candidate
            // Give it a high ortho score (it came from a proper correction chain)
            for err in &errors {
                if !err.suggestion.is_empty() && err.suggestion != *candidate {
                    let sug = err.suggestion.to_lowercase();
                    println!("  grammar suggests: '{}' → '{}'", candidate, sug);
                    if seen_passing.insert(sug.clone()) {
                        edit_distances.insert(sug.clone(), 1); // treat as distance 1 (came from valid chain)
                        passing.push((sug, 1.0)); // high ortho score
                    }
                }
            }
        }
    }

    // Re-rank using score_spelling — the exact same function the BERT worker uses.
    // Uses boundary scoring: "context_before<mask> context_after" + first-token logit.
    let grammar_suggested: HashSet<String> = passing.iter()
        .filter(|(_, s)| (*s - 1.0).abs() < 0.01)
        .map(|(c, _)| c.clone())
        .collect();
    println!("  Grammar-valid: {} candidates", passing.len());
    if passing.len() > 1 {
        let sentence_lower = sentence.to_lowercase();
        let (context_before, context_after) = if let Some(pos) = sentence_lower.find(&word_lower) {
            (sentence_lower[..pos].to_string(), sentence_lower[pos + word_lower.len()..].to_string())
        } else {
            (sentence_lower.clone(), String::new())
        };

        // Step 1: Fast boundary scoring (score_spelling) — rank all candidates (~24ms)
        let all_candidates: Vec<String> = passing.iter().take(30).map(|(c, _)| c.clone()).collect();
        let ortho_map: HashMap<String, f32> = passing.iter().cloned().collect();
        let boundary_ranked = match nostos_cognio::spelling::score_spelling(model, &context_before, &context_after, &all_candidates) {
            Ok(result) => {
                let mut ranked: Vec<(String, f32)> = result.scored_candidates.iter().map(|(c, bs)| {
                    let ortho = ortho_map.get(c.as_str()).copied().unwrap_or(0.5);
                    let eff = if grammar_suggested.contains(c) { 1.0 } else { ortho };
                    (c.clone(), bs * eff.sqrt())
                }).collect();
                ranked.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap());
                ranked
            }
            Err(_) => passing.clone(),
        };

        // Step 2: Full sentence scoring on top 5 (~200ms × 5 = ~1s)
        // Merge boundary top + ortho top to avoid missing good candidates
        let mut top_set: Vec<String> = Vec::new();
        let mut top_seen = HashSet::new();
        for (c, _) in boundary_ranked.iter().take(3) {
            if top_seen.insert(c.clone()) { top_set.push(c.clone()); }
        }
        // Also add top ortho candidates (they may have high sentence scores)
        for (c, _) in passing.iter().take(3) {
            if top_seen.insert(c.clone()) { top_set.push(c.clone()); }
        }
        let top3 = top_set;
        println!("  Boundary top 3: {:?}", top3.iter().zip(boundary_ranked.iter().take(3)).map(|(c, (_, s))| format!("{}({:.1})", c, s)).collect::<Vec<_>>());
        let mut final_ranked: Vec<(String, f32)> = Vec::new();
        for candidate in &top3 {
            let corrected = sentence_lower.replacen(&word_lower, candidate, 1);
            let sent_score = bert_sentence_score(model, &corrected);
            let ortho = ortho_map.get(candidate.as_str()).copied().unwrap_or(0.5);
            let eff = if grammar_suggested.contains(candidate) { 1.0 } else { ortho };
            let final_score = sent_score * eff.sqrt();
            println!("  full-sentence: '{}' sent={:.3} × sqrt(ortho {:.2}{}) = {:.3}",
                candidate, sent_score, eff,
                if grammar_suggested.contains(candidate) { " grammar" } else { "" },
                final_score);
            final_ranked.push((candidate.clone(), final_score));
        }
        final_ranked.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap());
        passing = final_ranked;
    }

    if !passing.is_empty() { passing } else { ortho_scored }
}

/// Score a sentence using BERT pseudo-log-likelihood (same as main.rs bert_sentence_score).
/// Masks each word, checks how well BERT predicts the actual word.
fn bert_sentence_score(model: &mut Model, sentence: &str) -> f32 {
    let words: Vec<&str> = sentence.split_whitespace().collect();
    if words.is_empty() { return f32::NEG_INFINITY; }

    let mut total_score: f32 = 0.0;
    for i in 0..words.len() {
        let masked: String = words.iter().enumerate()
            .map(|(j, w)| if j == i { "<mask>" } else { *w })
            .collect::<Vec<_>>()
            .join(" ");

        if let Ok((logits, _)) = model.single_forward(&masked) {
            let word_clean = words[i].trim_matches(|c: char| c.is_ascii_punctuation());
            let token_with_g = format!("Ġ{}", word_clean.to_lowercase());
            let token_id = model.tokenizer.token_to_id(&token_with_g)
                .or_else(|| model.tokenizer.token_to_id(&word_clean.to_lowercase()));
            if let Some(tid) = token_id {
                total_score += logits[tid as usize];
            }
        }
    }
    total_score / words.len() as f32
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
    let training = base.join("../contexter-repo/training-data");

    let onnx_path = training.join("onnx/norbert4_base_int8.onnx");
    let tok_path = training.join("onnx/tokenizer.json");
    println!("Loading NorBERT4...");
    let mut model = Model::load(onnx_path.to_str().unwrap(), tok_path.to_str().unwrap())
        .expect("Failed to load model");
    println!("Loaded. Vocab: {}", model.vocab_size());

    let dict_path = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let grammar_rules_path = base.join("../syntaxer/grammar_rules.pl");
    let syntaxer_dir = base.join("../syntaxer");
    let swipl_dll = if cfg!(target_os = "macos") {
        "/Applications/SWI-Prolog.app/Contents/Frameworks/libswipl.dylib"
    } else {
        "C:/Program Files/swipl/bin/libswipl.dll"
    };
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
        // BERT must rank "brus" over "bakrus" in context
        ("Barna fikk boller og bbrus.", "bbrus", "brus"),
        // BERT must rank "godterier" over "lotterier" in context
        ("Barna fikk gåtterier.", "gåtterier", "godterier|godteri"),
        // First-char wrong + needs inflection: sjøkken → kjøkkenet (via kjøkken + grammar)
        ("Jeg har mange gryter på sjøkken mitt.", "sjøkken", "kjøkken|kjøkkenet"),
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
