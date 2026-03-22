use std::path::PathBuf;
use nostos_cognio::model::Model;
use nostos_cognio::spelling;

fn main() {
    let model_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../contexter-repo/training-data/onnx/norbert4_base_int8.onnx");
    let tokenizer_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../contexter-repo/training-data/onnx/tokenizer.json");

    eprintln!("Loading NorBERT4...");
    let mut model = Model::load(model_path.to_str().unwrap(), tokenizer_path.to_str().unwrap())
        .expect("Failed to load model");
    eprintln!("Ready.\n");

    // Test: "Han spiller fotbollx godt." → candidates
    let context_before = "Han spiller";
    let context_after = "godt.";
    let candidates: Vec<String> = vec![
        "fotball", "foto", "fotboll", "fotballe", "fotballspiller", "fotografi",
    ].into_iter().map(String::from).collect();

    println!("=== BERT scoring: '{} <mask> {}' ===", context_before, context_after);
    let result = spelling::score_spelling(&mut model, context_before, context_after, &candidates)
        .expect("scoring failed");
    for (cand, score) in &result.scored_candidates {
        println!("  {:20} score={:.4}", cand, score);
    }
    println!("\n  Best: {} (score={:.4})", result.best_candidate, result.best_score);

    // Also test the tokenization to see what tokens the model sees
    println!("\n=== Token check ===");
    let masked = format!("{} <mask> {}", context_before, context_after);
    let encoding = model.tokenizer.encode(&*masked, true).unwrap();
    let ids = encoding.get_ids();
    let tokens: Vec<String> = encoding.get_tokens().iter().map(|s: &String| s.to_string()).collect();
    println!("  Input: '{}'", masked);
    println!("  Tokens: {:?}", tokens);
    println!("  IDs: {:?}", ids);

    // Check which position is mask
    let mask_pos = ids.iter().position(|&id| id == 4); // mask_token_id = 4
    println!("  Mask position: {:?}", mask_pos);

    // Show top 20 predictions at mask position
    let (logits, _) = model.single_forward(&masked).unwrap();
    let mut scored: Vec<(usize, f32)> = logits.iter().enumerate().map(|(i, s)| (i, *s)).collect();
    scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap());
    println!("\n=== Top 20 BERT predictions at <mask> ===");
    for (tid, score) in scored.iter().take(20) {
        let token = model.tokenizer.id_to_token(*tid as u32).unwrap_or_default();
        println!("  {:5} {:20} score={:.4}", tid, token, score);
    }

    // Check specific candidates
    println!("\n=== Specific candidate logits ===");
    for word in &["fotball", "foto", "fotboll", "fotballe"] {
        let gw = format!("Ġ{}", word);
        let tid = model.tokenizer.token_to_id(&gw);
        let score = tid.map(|id| logits[id as usize]).unwrap_or(f32::NEG_INFINITY);
        println!("  {:15} token='{}' id={:?} score={:.4}", word, gw, tid, score);
    }
}
