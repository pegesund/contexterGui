use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::Arc;

fn main() -> anyhow::Result<()> {
    // Set ORT path for macOS
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

    let data = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../contexter-repo/training-data");
    let onnx = data.join("onnx/norbert4_base_int8.onnx");
    let tok = data.join("onnx/tokenizer.json");
    let wf_path = data.join("wordfreq.tsv");
    let dict_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../rustSpell/mtag-rs/data/fullform_bm.mfst");

    eprintln!("Loading model...");
    let mut model = nostos_cognio::model::Model::load(onnx.to_str().unwrap(), tok.to_str().unwrap())?;
    eprintln!("Building prefix index...");
    let pi = nostos_cognio::prefix_index::build_prefix_index(&model.tokenizer);
    let wf = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);
    eprintln!("Loading dictionary...");
    let analyzer = mtag::Analyzer::new(dict_path.to_str().unwrap())
        .map_err(|e| anyhow::anyhow!("{}", e))?;
    eprintln!("Ready.\n");

    let fallback_dict = |w: &str| -> bool { analyzer.has_word(w) };
    let fallback_prefix = |p: &str, limit: usize| -> Vec<String> { analyzer.prefix_lookup(p, limit) };

    let tests = vec![
        ("Fotball er en morsom", "s"),
        ("Jeg liker å spille", "fotba"),
        ("Hun er veldig", "f"),
        ("Det var en fin", "d"),
    ];

    // Test with None wordfreq (simulating broken worker)
    println!("=== TEST WITH NONE WORDFREQ ===");
    let results_no_wf = nostos_cognio::complete::complete_word(
        &mut model, "Fotball er en morsom", "s", &pi,
        None, None, Some(&fallback_dict), Some(&fallback_prefix), None,
        1.0, 10.0, 10, 3,
    )?;
    println!("Results WITHOUT wordfreq:");
    for (i, c) in results_no_wf.iter().enumerate() {
        println!("  {:>2}. {} ({:.1})", i + 1, c.word, c.score);
    }
    println!();

    // Test with None analyzer (simulating broken worker)
    println!("=== TEST WITH NONE ANALYZER ===");
    let results_no_an = nostos_cognio::complete::complete_word(
        &mut model, "Fotball er en morsom", "s", &pi,
        None, Some(&wf), None, None, None,
        1.0, 10.0, 10, 3,
    )?;
    println!("Results WITHOUT analyzer:");
    for (i, c) in results_no_an.iter().enumerate() {
        println!("  {:>2}. {} ({:.1})", i + 1, c.word, c.score);
    }
    println!();

    for (context, prefix) in &tests {
        println!("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
        println!("Context: '{}'  prefix: '{}'", context, prefix);

        let t0 = std::time::Instant::now();
        let results = nostos_cognio::complete::complete_word(
            &mut model,
            context,
            prefix,
            &pi,
            None,           // baselines
            Some(&wf),
            Some(&fallback_dict),
            Some(&fallback_prefix),
            None,           // embedding_store
            1.0,            // pmi_weight
            10.0,           // topic_boost
            10,             // top_n
            3,              // max_steps
        )?;
        let ms = t0.elapsed().as_millis();

        println!("Results ({} ms):", ms);
        for (i, c) in results.iter().enumerate() {
            println!("  {:>2}. {} ({:.1})", i + 1, c.word, c.score);
        }
        println!();
    }

    Ok(())
}
