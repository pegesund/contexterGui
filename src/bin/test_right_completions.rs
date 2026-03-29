/// Test right-column completions — uses complete_word with empty prefix.
/// Usage: ORT_DYLIB_PATH=/opt/homebrew/lib/libonnxruntime.dylib cargo run --release --bin test-right-completions

use nostos_cognio::model::Model;
use nostos_cognio::baseline::compute_baseline;
use nostos_cognio::complete::complete_word;
use nostos_cognio::prefix_index::build_prefix_index;
use std::path::PathBuf;

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

    let pi = build_prefix_index(&model.tokenizer);
    let fallback = |w: &str| -> bool { analyzer.has_word(w) };

    println!("Loaded. Vocab: {}, wordfreq: {}\n", model.vocab_size(), wordfreq.len());

    // (context, expected_words_any_in_top6)
    let tests: Vec<(&str, Vec<&str>)> = vec![
        ("Fotball er en norsk idrett", vec!["som", "og", "med", "for", "der"]),
        ("Fotball er en spennende", vec!["sport", "idrett", "sjanger"]),
        ("Han liker å spise", vec!["mat", "middag", "maten"]),
        ("Vi skal på", vec!["tur", "ferie", "skolen", "jobb"]),
        ("Hun er en flink", vec!["jente", "elev", "student", "pike"]),
        ("Jeg liker å lese", vec!["bøker", "aviser", "romaner"]),
        ("Det var en kald", vec!["vinter", "natt", "dag", "kveld"]),
        ("Barna leker i", vec!["hagen", "parken", "sanden"]),
        ("Katten sover på", vec!["sofaen", "gulvet", "sengen"]),
        ("Norge er et", vec!["land", "kongerike", "demokrati"]),
        ("Hun kjøpte en", vec!["bil", "bok", "jakke", "leilighet"]),
        ("Vi spiste", vec!["fisk", "laks", "pizza", "pasta", "middag"]),
        ("Elevene satt i", vec!["klasserommet", "ringen", "salen"]),
        ("Det regner mye om", vec!["høsten", "sommeren", "vinteren"]),
        ("Han spiller", vec!["fotball", "piano", "gitar"]),
    ];

    let mut hits = 0;
    let total = tests.len();

    for (context, expected) in &tests {
        let results = match complete_word(
            &mut model, context, "", &pi,
            Some(&baselines), Some(&wordfreq),
            Some(&fallback), None, None,
            1.0, 10.0, 15, 0,
        ) {
            Ok(r) => r,
            Err(e) => { println!("  [ERR] {} — {}", context, e); continue; }
        };

        let words: Vec<&str> = results.iter().map(|c| c.word.as_str()).collect();
        let top6 = &words[..6.min(words.len())];
        let hit = expected.iter().any(|e| top6.iter().any(|w| w.to_lowercase() == e.to_lowercase()));
        if hit { hits += 1; }
        let status = if hit { "OK" } else { "MISS" };

        println!("  [{}] \"{}\"", status, context);
        println!("    TOP6: {}", top6.join(", "));
        println!("    EXPECT: {}", expected.join(", "));
        println!();
    }

    println!("Score: {}/{}", hits, total);
}
