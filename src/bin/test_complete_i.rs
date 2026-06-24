/// Diagnostic + regression tests for complete_word.
/// Run via `cargo run --release --bin test-complete-i`.

use nostos_cognio::model::Model;
use nostos_cognio::baseline::compute_baseline;
use nostos_cognio::complete::complete_word;
use nostos_cognio::prefix_index::build_prefix_index;
use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::Arc;

fn diag(model: &mut Model, baselines: &nostos_cognio::baseline::Baselines, analyzer: &mtag::Analyzer, ctx: &str) {
    let mask = model.mask_token_str().to_string();
    let masked = format!("{ctx}{mask} .");
    println!("\n--- DIAG masked='{masked}' ---");
    let (raw, _) = model.single_forward(&masked).expect("forward");
    let pmi_w = 1.0 * 0.7; // mid-sentence weight from complete.rs

    // Word-initial tokens only
    let mut entries: Vec<(usize, String, f32, f32, f32)> = Vec::new();
    for (i, tok) in model.id_to_token.iter().enumerate() {
        if !tok.starts_with('\u{0120}') { continue; }
        let word = model.tokenizer.decode(&[i as u32], false)
            .unwrap_or_default().trim().to_lowercase();
        if word.len() < 2 { continue; }
        if !analyzer.has_word(&word) { continue; }
        let bert = raw[i];
        let bl = baselines.sentence[i];
        let pmi = bert + pmi_w * (bert - bl);
        entries.push((i, word, bert, bl, pmi));
    }
    let mut by_raw = entries.clone();
    by_raw.sort_by(|a, b| b.2.partial_cmp(&a.2).unwrap_or(std::cmp::Ordering::Equal));
    let mut by_pmi = entries.clone();
    by_pmi.sort_by(|a, b| b.4.partial_cmp(&a.4).unwrap_or(std::cmp::Ordering::Equal));

    println!("  Top 10 by RAW BERT logit:");
    for (_, w, r, bl, pmi) in by_raw.iter().take(10) {
        println!("    {:20} raw={:.2} baseline={:.2} pmi-adj={:.2}", w, r, bl, pmi);
    }
    println!("  Top 10 by PMI-ADJUSTED:");
    for (_, w, r, bl, pmi) in by_pmi.iter().take(10) {
        println!("    {:20} raw={:.2} baseline={:.2} pmi-adj={:.2}", w, r, bl, pmi);
    }
    // Specific words to track
    println!("  Tracked candidates:");
    for target in &["spillere", "barn", "også", "fordi", "og", "men", "som", "fans", "viktig"] {
        if let Some(e) = entries.iter().find(|(_,w,_,_,_)| w == target) {
            println!("    {:20} raw={:.2} baseline={:.2} pmi-adj={:.2}", e.1, e.2, e.3, e.4);
        }
    }
}

fn main() {
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let training = base.join("../contexter-repo/training-data");
    let onnx_path = training.join("onnx/norbert4_base_int8.onnx");
    let tok_path = training.join("onnx/tokenizer.json");
    let dict_path = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let wf_path = training.join("wordfreq.tsv");

    println!("Loading...");
    let mut model = Model::load(onnx_path.to_str().unwrap(), tok_path.to_str().unwrap())
        .expect("Failed to load model");
    let analyzer = mtag::Analyzer::new(dict_path.to_str().unwrap())
        .expect("Failed to load analyzer");
    let wordfreq = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);
    let baselines = compute_baseline(&mut model).expect("baselines");
    let pi = build_prefix_index(&model.tokenizer);
    println!("Loaded.\n");

    diag(&mut model, &baselines, &analyzer, "Fotball er en interessant sport for nybegynnere");

    let has_word = |w: &str| -> bool { analyzer.has_word(w) };
    let prefix_lookup = |p: &str, limit: usize| -> Vec<String> { analyzer.prefix_lookup(p, limit) };

    let tests = vec![
        // Regression: PMI used to bury connectors ("og", "som", "også") under
        // context-specific nouns ("barn", "spillere", "nybegynnere") for
        // empty-prefix next-word prediction. After the PMI suppression we
        // should see connectors and high-baseline naturals re-surface.
        ("Fotball er en interessant sport for nybegynnere", ""),
        ("Fotball er en spennende", "i"),
        ("Fotball er en spennende", "id"),
        ("Fotball er en spennende", "is"),
        ("Jeg spiser", "is"),
        ("Han liker å spise", "m"),
        ("Vi skal på", "t"),
        ("Det var en kald", "v"),
    ];

    for (context, prefix) in &tests {
        println!("=== ctx='{}' prefix='{}' ===", context, prefix);
        match complete_word(
            &mut model, context, prefix, &pi,
            Some(&baselines), Some(&wordfreq),
            Some(&has_word), Some(&prefix_lookup), None,
            1.0, 10.0, 10, 1,
        ) {
            Ok(results) => {
                for (i, r) in results.iter().enumerate() {
                    println!("  #{}: {:20} score={:.1}", i+1, r.word, r.score);
                }
            }
            Err(e) => println!("  ERROR: {}", e),
        }
        println!();
    }
}
