/// Test BERT predictions and complete_word for a given context.
/// Usage: cargo run --release --bin test_bert_predict

use nostos_cognio::model::Model;
use nostos_cognio::complete::complete_word;
use nostos_cognio::prefix_index;
use std::collections::HashMap;
use std::path::PathBuf;

fn main() {
    // Set ORT path
    let ort_candidates = vec![
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../onnxruntime/lib/libonnxruntime.dylib"),
        PathBuf::from("/usr/local/lib/libonnxruntime.dylib"),
        PathBuf::from("/opt/homebrew/lib/libonnxruntime.dylib"),
    ];
    for p in &ort_candidates {
        if p.exists() {
            unsafe { std::env::set_var("ORT_DYLIB_PATH", p); }
            break;
        }
    }

    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../contexter-repo/training-data/onnx");
    let onnx = base.join("norbert4_base_int8.onnx");
    let tok = base.join("tokenizer.json");

    eprintln!("Loading NorBERT4...");
    let mut model = Model::load(onnx.to_str().unwrap(), tok.to_str().unwrap())
        .expect("Failed to load model");
    eprintln!("Loaded. Vocab: {}", model.vocab_size());

    // Build prefix index
    let pi = prefix_index::build_prefix_index(&model.tokenizer);

    // Raw logit check for Ġen
    let masked = "Fotball er et morsomt<mask> .";
    let (logits, _) = model.single_forward(masked).expect("forward failed");
    if let Some(tid) = model.tokenizer.token_to_id("Ġen") {
        println!("Raw BERT logit for Ġen at '{}': {:.1}", masked, logits[tid as usize]);
    }

    // Load wordfreq (same as app)
    let wf_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../contexter-repo/training-data/wordfreq.tsv");
    let wf = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);
    let wf_map: HashMap<String, u64> = wf.into_iter().collect();
    println!("WordFreq: {} words", wf_map.len());

    // Test complete_word with prefix='' and max_steps=0 (what the app does for right/open column)
    let context = "Fotball er et morsomt";
    println!("\n=== complete_word(ctx='{}', prefix='', max_steps=0) ===", context);
    match complete_word(&mut model, context, "", &pi, None, Some(&wf_map), None, None, None, 1.0, 10.0, 15, 0) {
        Ok(results) => {
            for (i, c) in results.iter().enumerate() {
                let marker = if c.word.to_lowercase() == "en" { " <--- BUG" } else { "" };
                println!("  {:>2}. {:>20}  score={:.1}{}", i+1, c.word, c.score, marker);
            }
        }
        Err(e) => println!("Error: {}", e),
    }

    // Also test "en morsom" context
    let context2 = "Fotball er en morsom";
    println!("\n=== complete_word(ctx='{}', prefix='', max_steps=0) ===", context2);
    match complete_word(&mut model, context2, "", &pi, None, Some(&wf_map), None, None, None, 1.0, 10.0, 15, 0) {
        Ok(results) => {
            for (i, c) in results.iter().enumerate() {
                let marker = if ["og", "en", "er", "et", "men", "som"].contains(&c.word.to_lowercase().as_str()) { " <--- FUNCTION WORD" } else { "" };
                println!("  {:>2}. {:>20}  score={:.1}{}", i+1, c.word, c.score, marker);
            }
        }
        Err(e) => println!("Error: {}", e),
    }
}
