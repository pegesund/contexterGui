use std::path::PathBuf;

fn data_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../contexter-repo/training-data")
}

fn main() -> anyhow::Result<()> {
    // Set ORT
    let ort_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../onnxruntime/onnxruntime-win-x64-1.23.0/lib/onnxruntime.dll");
    unsafe { std::env::set_var("ORT_DYLIB_PATH", &ort_path); }

    let data = data_dir();
    let onnx = data.join("onnx/norbert4_base_int8.onnx");
    let tok = data.join("onnx/tokenizer.json");
    let wf_path = data.join("wordfreq.tsv");
    let minilm_onnx = data.join("minilm-onnx/model_optimized.onnx");
    let minilm_tok = data.join("minilm-onnx/tokenizer.json");
    let embed_cache = data.join("word_embeddings.bin");

    eprintln!("Loading NorBERT4...");
    let mut model = nostos_cognio::model::Model::load(onnx.to_str().unwrap(), tok.to_str().unwrap())?;
    eprintln!("Building prefix index...");
    let pi = nostos_cognio::prefix_index::build_prefix_index(&model.tokenizer);
    eprintln!("Computing baselines...");
    let baselines = nostos_cognio::baseline::compute_baseline(&mut model)?;
    let wf = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);

    eprintln!("Loading MiniLM...");
    let embedder = nostos_cognio::embeddings::Embedder::load(
        minilm_onnx.to_str().unwrap(), minilm_tok.to_str().unwrap())?;
    let mut store = nostos_cognio::embeddings::EmbeddingStore::new(
        embedder, &wf, Some(embed_cache.as_path()))?;
    eprintln!("Ready.\n");

    // Test 1: Without topic context
    println!("=== Test 1: No topic context ===");
    println!("Context: 'Om ferien går jeg ofte til '  prefix: 'v'");
    let results = nostos_cognio::complete::complete_word(
        &mut model, "Om ferien går jeg ofte til", "v",
        &pi, Some(&baselines), Some(&wf), None, None,
        1.0, 10.0, 5, 3,
    )?;
    for c in &results {
        println!("  {} ({:.1})", c.word, c.score);
    }

    // Test 2: With fishing/bathing topic sentences embedded
    println!("\n=== Test 2: With topic context (fiske/bade) ===");
    let sentences = vec![
        "I fritiden liker jeg å fiske og bade.".to_string(),
    ];
    let n = store.sync_sentences(&sentences)?;
    println!("Embedded {} sentences", n);

    // Check topic words for 'v'
    let topics = store.topic_words("v", 10);
    let mut sorted: Vec<_> = topics.iter().collect();
    sorted.sort_by(|a, b| b.1.partial_cmp(a.1).unwrap());
    println!("Topic words for 'v': {:?}", &sorted[..sorted.len().min(10)]);

    println!("\nContext: 'Om ferien går jeg ofte til '  prefix: 'v'");
    let results = nostos_cognio::complete::complete_word(
        &mut model, "Om ferien går jeg ofte til", "v",
        &pi, Some(&baselines), Some(&wf), None, Some(&store),
        1.0, 10.0, 5, 3,
    )?;
    for c in &results {
        println!("  {} ({:.1})", c.word, c.score);
    }

    // Test 3: Context ends with punct (cross-sentence mode)
    println!("\n=== Test 3: Context with trailing punct (cross-sentence) ===");
    let sentences = vec![
        "I fritiden liker jeg å fiske og bade.".to_string(),
        "Vi dro til vannet for å bade.".to_string(),
        "Vannet var varmt og deilig.".to_string(),
    ];
    let n = store.sync_sentences(&sentences)?;
    println!("Embedded {} new sentences ({} total)", n, sentences.len());

    // NOTE: embeddings only activate when context ends with punct!
    println!("\nContext: 'I fritiden liker jeg å fiske og bade. Om ferien går jeg ofte til '  prefix: 'v'");
    let results = nostos_cognio::complete::complete_word(
        &mut model, "I fritiden liker jeg å fiske og bade. Om ferien går jeg ofte til", "v",
        &pi, Some(&baselines), Some(&wf), None, Some(&store),
        1.0, 10.0, 5, 3,
    )?;
    for c in &results {
        println!("  {} ({:.1})", c.word, c.score);
    }

    // Test 4: New sentence after period
    println!("\n=== Test 4: Starting new sentence after period ===");
    println!("Context: 'I fritiden liker jeg å fiske og bade. '  prefix: 'V'");
    let results = nostos_cognio::complete::complete_word(
        &mut model, "I fritiden liker jeg å fiske og bade.", "V",
        &pi, Some(&baselines), Some(&wf), None, Some(&store),
        1.0, 10.0, 5, 3,
    )?;
    for c in &results {
        println!("  {} ({:.1})", c.word, c.score);
    }

    Ok(())
}
