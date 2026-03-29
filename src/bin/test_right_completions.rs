/// Test right-column completions with PMI + mtag filter.
/// Reproduces the exact build_right_completions path from the app.
/// Usage: ORT_DYLIB_PATH=/opt/homebrew/lib/libonnxruntime.dylib cargo run --release --bin test-right-completions

use nostos_cognio::model::Model;
use nostos_cognio::baseline::{compute_baseline, Baselines};
use nostos_cognio::complete::Completion;
use std::collections::{HashMap, HashSet};
use std::path::PathBuf;

fn build_right_completions_test(
    model: &Model,
    logits: &[f32],
    wordfreq: Option<&HashMap<String, u64>>,
    nearby_words: &HashSet<String>,
    left_words: &HashSet<String>,
    baselines: Option<&Baselines>,
    analyzer: Option<&mtag::Analyzer>,
) -> Vec<(String, f32)> {
    let is_valid = |w: &str| -> bool {
        let key = w.to_lowercase();
        if nearby_words.contains(&key) { return false; }
        if !wordfreq.map_or(true, |wf| wf.contains_key(&key)) { return false; }
        if let Some(az) = analyzer {
            if !az.has_word(&key) { return false; }
        }
        true
    };

    // PMI: subtract baseline to demote generically common words
    let pmi_logits: Vec<f32> = if let Some(bl) = baselines {
        logits.iter().enumerate().map(|(i, &raw)| {
            let base = if i < bl.sentence.len() { bl.sentence[i] } else { 0.0 };
            raw + 1.0 * (raw - base)
        }).collect()
    } else {
        logits.to_vec()
    };

    let mut seen: HashSet<String> = HashSet::new();
    let mut all_scored: Vec<(String, f32)> = model.id_to_token.iter()
        .enumerate()
        .filter(|(_, tok)| tok.starts_with('Ġ'))
        .filter_map(|(i, _)| {
            let decoded = model.tokenizer
                .decode(&[i as u32], false)
                .unwrap_or_default().trim().to_lowercase();
            if decoded.is_empty() || decoded.len() <= 1 { return None; }
            if !is_valid(&decoded) || left_words.contains(&decoded) { return None; }
            if !seen.insert(decoded.clone()) { return None; }
            Some((decoded, pmi_logits[i]))
        })
        .collect();
    all_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
    all_scored.into_iter().take(10).collect()
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
    let baselines = compute_baseline(&mut model).expect("Failed to compute baselines");
    println!("Loaded. Vocab: {}, wordfreq: {}\n", model.vocab_size(), wordfreq.len());

    let empty_nearby: HashSet<String> = HashSet::new();
    let empty_left: HashSet<String> = HashSet::new();

    let tests = vec![
        "Fotball er en spennende <mask> .",
        "Han liker å spise <mask> .",
        "Vi skal på <mask> i morgen.",
        "Hun er en flink <mask> .",
        "Jeg liker å lese <mask> .",
        "Det var en kald <mask> .",
        "Barna leker i <mask> .",
        "Katten sover på <mask> .",
        "Om våren blomstrer <mask> .",
        "Norge er et <mask> .",
        "Hun kjøpte en <mask> .",
        "Vi spiste <mask> til middag.",
        "Elevene satt i <mask> .",
        "Det regner mye om <mask> .",
        "Han spiller <mask> hver dag.",
    ];

    for sent in &tests {
        let logits = match model.single_forward(sent) {
            Ok((l, _)) => l,
            Err(e) => { println!("  ERROR: {}", e); continue; }
        };

        let with_pmi = build_right_completions_test(
            &model, &logits, Some(&wordfreq), &empty_nearby, &empty_left,
            Some(&baselines), Some(&analyzer),
        );
        let without_pmi = build_right_completions_test(
            &model, &logits, Some(&wordfreq), &empty_nearby, &empty_left,
            None, Some(&analyzer),
        );

        let pmi_words: Vec<&str> = with_pmi.iter().map(|(w, _)| w.as_str()).collect();
        let raw_words: Vec<&str> = without_pmi.iter().map(|(w, _)| w.as_str()).collect();

        println!("  {}", sent);
        println!("    RAW: {}", raw_words[..6.min(raw_words.len())].join(", "));
        println!("    PMI: {}", pmi_words[..6.min(pmi_words.len())].join(", "));
        println!();
    }
}
