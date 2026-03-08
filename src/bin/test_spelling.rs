/// Console test for spelling suggestions with BERT context.
/// Usage: cargo run --release --bin test_spelling

use nostos_cognio::model::Model;
use std::collections::HashMap;
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

fn main() {
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let training = base.join("../../contexter-repo/training-data");

    let onnx_path = training.join("onnx/norbert4_base_int8.onnx");
    let tok_path = training.join("onnx/tokenizer.json");
    println!("Loading NorBERT4...");
    let mut model = Model::load(onnx_path.to_str().unwrap(), tok_path.to_str().unwrap())
        .expect("Failed to load model");
    println!("Loaded. Vocab: {}", model.vocab_size());

    // Load dictionary for fuzzy lookup
    let dict_path = base.join("../../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let checker = nostos_cognio::grammar::GrammarChecker::new(
        dict_path.to_str().unwrap(),
        "",
    ).expect("Failed to load dictionary");

    let test_cases = vec![
        ("De skullendex få bossller og brus.", "skullendex", "Expected: skulle"),
        ("De skullendex få bossller og brus.", "bossller", "Expected: boller"),
        ("Jeg liker å spile fotball.", "spile", "Expected: spille"),
    ];

    for (sentence, misspelled, expected) in &test_cases {
        println!("\n{}", "=".repeat(60));
        println!("Sentence: '{}'", sentence);
        println!("Misspelled: '{}' ({})", misspelled, expected);

        let word_lower = misspelled.to_lowercase();
        let word_trigrams = trigrams(&word_lower);
        let word_first = word_lower.chars().next().unwrap_or(' ');

        // Build masked context
        let sentence_lower = sentence.to_lowercase();
        let masked = if let Some(pos) = sentence_lower.find(&word_lower) {
            let before = &sentence[..pos];
            let after = &sentence[pos + word_lower.len()..];
            format!("{}<mask>{}", before.trim_end(), after)
        } else {
            format!("{} <mask>", sentence)
        };
        println!("Masked: '{}'", masked);

        // Get BERT logits
        let (logits, ms) = model.single_forward(&masked).expect("Forward failed");
        println!("Forward pass: {:.0}ms", ms);

        // Extract top-200 BERT token-words
        let mut indexed: Vec<(usize, f32)> = logits.iter().cloned().enumerate().collect();
        indexed.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap());

        let mut bert_words: Vec<(String, f32)> = Vec::new();
        for &(token_id, score) in indexed.iter().take(200) {
            if let Some(token) = model.tokenizer.id_to_token(token_id as u32) {
                let clean = token.replace('Ġ', "").to_lowercase();
                if clean.len() < 2 || clean == word_lower { continue; }
                if clean.chars().any(|c| !c.is_alphanumeric() && c != '-') { continue; }
                if !bert_words.iter().any(|(w, _)| *w == clean) {
                    bert_words.push((clean, score));
                }
            }
        }
        println!("BERT token-words: {}", bert_words.len());

        // Build BERT score lookup
        let bert_scores: HashMap<String, f32> = bert_words.iter().cloned().collect();

        // Score BERT words by trigram similarity
        let mut scored: Vec<(String, f32)> = bert_words.into_iter()
            .filter_map(|(w, bert_score)| {
                let w_trigrams = trigrams(&w);
                let common = word_trigrams.iter()
                    .filter(|t| w_trigrams.contains(t))
                    .count();
                if common == 0 && w.chars().next().unwrap_or(' ') != word_first {
                    return None;
                }
                let max_trigrams = word_trigrams.len().max(w_trigrams.len()).max(1);
                let trigram_score = common as f32 / max_trigrams as f32;
                let mut score = bert_score / 10.0;
                score += trigram_score;
                if w.chars().next().unwrap_or(' ') == word_first {
                    score += 0.3;
                }
                Some((w, score))
            })
            .collect();

        // Add fuzzy matches (distance 2) scored with BERT
        let fuzzy = checker.fuzzy_lookup(&word_lower, 2);
        println!("Fuzzy(2): {} matches", fuzzy.len());
        for (w, dist) in fuzzy {
            if w == word_lower { continue; }
            if scored.iter().any(|(s, _)| s == &w) { continue; }
            let mut score = 1.0 - (dist as f32 * 0.2);
            if w.chars().next().unwrap_or(' ') == word_first {
                score += 0.3;
            }
            if let Some(&bs) = bert_scores.get(&w) {
                score += bs / 10.0;
            }
            scored.push((w, score));
        }

        scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap());
        scored.dedup_by(|a, b| a.0 == b.0);
        scored.truncate(15);

        println!("\nTop-15 candidates:");
        for (i, (word, score)) in scored.iter().enumerate() {
            let bert = bert_scores.get(word).copied().unwrap_or(0.0);
            println!("  #{}: {} (score={:.3}, bert={:.2})", i + 1, word, score, bert);
        }
    }
}
