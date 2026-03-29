/// Test word completion suggestions — standalone reproduction of the pipeline.
/// Usage: ORT_DYLIB_PATH=/opt/homebrew/lib/libonnxruntime.dylib cargo run --release --bin test-suggest

use nostos_cognio::model::Model;
use std::collections::{HashMap, HashSet};
use std::path::PathBuf;

fn main() {
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let training = base.join("../contexter-repo/training-data");
    let onnx_path = training.join("onnx/norbert4_base_int8.onnx");
    let tok_path = training.join("onnx/tokenizer.json");

    println!("Loading NorBERT4...");
    let mut model = Model::load(onnx_path.to_str().unwrap(), tok_path.to_str().unwrap())
        .expect("Failed to load model");

    let wf_path = training.join("wordfreq.tsv");
    let wordfreq = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);
    println!("Loaded. Vocab: {}, wordfreq: {} words\n", model.vocab_size(), wordfreq.len());

    let tests = vec![
        ("Fotball er en spennende <mask>.", "i"),
        ("Han liker å spise <mask>.", "m"),
        ("Vi skal på <mask> i morgen.", "t"),
        ("Hun er en flink <mask>.", "s"),
    ];

    for (masked, prefix) in &tests {
        println!("=== '{}' prefix='{}' ===", masked, prefix);

        // Step 1: single forward pass
        let logits = match model.single_forward(masked) {
            Ok((l, ms)) => { println!("  Forward: {:.1}ms", ms); l }
            Err(e) => { println!("  ERROR: {}", e); continue; }
        };

        // Step 2: find all tokens matching prefix
        let prefix_lower = prefix.to_lowercase();
        let mut matches: Vec<(u32, String)> = Vec::new();
        for tid in 0..model.vocab_size() as u32 {
            if let Some(token_str) = model.tokenizer.id_to_token(tid) {
                // Word-initial tokens start with Ġ (space prefix)
                if let Some(word) = token_str.strip_prefix("Ġ") {
                    let word_lower = word.to_lowercase();
                    if word_lower.starts_with(&prefix_lower) && word.len() >= 2 {
                        matches.push((tid, word_lower));
                    }
                }
            }
        }

        // Step 3: score and sort
        let mut scored: Vec<(String, f32, usize)> = matches.iter()
            .map(|(tid, word)| {
                let score = logits[*tid as usize];
                let n_toks = model.tokenizer.encode(format!(" {}", word), false)
                    .map(|e| e.get_ids().len()).unwrap_or(1);
                (word.clone(), score, n_toks)
            })
            .collect();
        scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

        // Step 4: filter by wordfreq + minimum length
        let min_len = prefix_lower.len() + 2;
        println!("  Raw top 10 (no filter):");
        for (w, s, t) in scored.iter().take(10) {
            let in_wf = wordfreq.contains_key(w.as_str());
            println!("    {:<25} logit={:.1} toks={} wf={} len={}", w, s, t, in_wf, w.chars().count());
        }

        let filtered: Vec<&(String, f32, usize)> = scored.iter()
            .filter(|(w, _, _)| {
                w.chars().count() >= min_len && wordfreq.contains_key(w.as_str())
            })
            .collect();
        println!("  Filtered top 10 (len>={}, in wordfreq):", min_len);
        for (w, s, t) in filtered.iter().take(10) {
            println!("    {:<25} logit={:.1} toks={}", w, s, t);
        }
        println!();
    }
}
