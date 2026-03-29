/// Reproduce exact complete_word call for prefix "i" after "Fotball er en spennende"
/// This is what the app calls — find where "id" and "is" come from and fix them.

use nostos_cognio::model::Model;
use nostos_cognio::baseline::compute_baseline;
use nostos_cognio::complete::complete_word;
use nostos_cognio::prefix_index::build_prefix_index;
use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::Arc;

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

    let has_word = |w: &str| -> bool { analyzer.has_word(w) };
    let prefix_lookup = |p: &str, limit: usize| -> Vec<String> { analyzer.prefix_lookup(p, limit) };

    let tests = vec![
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
