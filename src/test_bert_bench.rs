use std::path::PathBuf;
use std::time::Instant;
use ort::{session::{Session, builder::GraphOptimizationLevel}, inputs, value::TensorRef};

fn main() {
    let model_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../contexter-repo/training-data/onnx/norbert4_base_int8.onnx");
    let tokenizer_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../contexter-repo/training-data/onnx/tokenizer.json");

    let sentences: Vec<&str> = vec![
        "Han spiller<mask> godt.",
        "Fotball er en morsom<mask>.",
        "Jeg liker å spise<mask> og drikke melk.",
        "Det nevrale nettverket lærte seg å<mask> bilder.",
        "Vi jobber med maskinlæring og kunstig<mask> i Norge.",
        "Løsningen ble veldig godt mottatt og vi fikk også godt med<mask>.",
        "Prosjektet innbefattet trening av flere små algoritmer som var trent på<mask>.",
        "Et stort nevralt nettverk er<mask>.",
    ];

    let tokenizer = tokenizers::Tokenizer::from_file(tokenizer_path.to_str().unwrap()).unwrap();

    // How many cores?
    let num_cpus = std::thread::available_parallelism().map(|n| n.get()).unwrap_or(4);
    println!("CPU cores available: {}\n", num_cpus);

    // === Thread count sweep (single inference) ===
    println!("=== SINGLE INFERENCE: thread count sweep ===");
    for threads in [1, 2, 4, 6, 8, 10, 12] {
        if threads > num_cpus { break; }
        let mut session = Session::builder().unwrap()
            .with_optimization_level(GraphOptimizationLevel::Level3).unwrap()
            .with_intra_threads(threads).unwrap()
            .with_inter_threads(1).unwrap()
            .commit_from_file(model_path.to_str().unwrap()).unwrap();

        // Warmup
        run_inference(&mut session, &tokenizer, &sentences[0]);
        run_inference(&mut session, &tokenizer, &sentences[1]);

        let t = Instant::now();
        for s in &sentences {
            run_inference(&mut session, &tokenizer, s);
        }
        let total = t.elapsed();
        println!("  {:>2} threads: {} inferences in {:>4.0}ms ({:.1}ms avg)",
            threads, sentences.len(), total.as_millis(), total.as_millis() as f64 / sentences.len() as f64);
    }

    // === Thread count sweep (batched) ===
    println!("\n=== BATCHED (8 sentences): thread count sweep ===");
    for threads in [1, 2, 4, 6, 8, 10, 12] {
        if threads > num_cpus { break; }
        let mut session = Session::builder().unwrap()
            .with_optimization_level(GraphOptimizationLevel::Level3).unwrap()
            .with_intra_threads(threads).unwrap()
            .with_inter_threads(1).unwrap()
            .commit_from_file(model_path.to_str().unwrap()).unwrap();

        // Warmup
        run_batch_inference(&mut session, &tokenizer, &sentences);

        let t = Instant::now();
        let rounds = 5;
        for _ in 0..rounds {
            run_batch_inference(&mut session, &tokenizer, &sentences);
        }
        let total = t.elapsed();
        let per_batch = total.as_millis() as f64 / rounds as f64;
        let per_sent = per_batch / sentences.len() as f64;
        println!("  {:>2} threads: {:.1}ms/batch ({:.1}ms/sentence)",
            threads, per_batch, per_sent);
    }

    // === Check XNNPACK availability ===
    println!("\n=== XNNPACK EP ===");
    {
        use ort::ep;
        let xnn = ep::XNNPACK::default().build();
        match Session::builder().unwrap()
            .with_execution_providers([xnn]).unwrap()
            .commit_from_file(model_path.to_str().unwrap())
        {
            Ok(mut session) => {
                // Warmup
                run_inference(&mut session, &tokenizer, &sentences[0]);
                let t = Instant::now();
                for s in &sentences {
                    run_inference(&mut session, &tokenizer, s);
                }
                let total = t.elapsed();
                println!("  Single: {} inferences in {:.0}ms ({:.1}ms avg)",
                    sentences.len(), total.as_millis(), total.as_millis() as f64 / sentences.len() as f64);

                // Batched
                run_batch_inference(&mut session, &tokenizer, &sentences);
                let t = Instant::now();
                for _ in 0..5 {
                    run_batch_inference(&mut session, &tokenizer, &sentences);
                }
                let total = t.elapsed();
                println!("  Batch:  {:.1}ms/batch ({:.1}ms/sentence)",
                    total.as_millis() as f64 / 5.0,
                    total.as_millis() as f64 / (5.0 * sentences.len() as f64));
            }
            Err(e) => println!("  Not available: {}", e),
        }
    }
}

fn run_batch_inference(session: &mut Session, tokenizer: &tokenizers::Tokenizer, texts: &[&str]) {
    let encodings: Vec<_> = texts.iter()
        .map(|t| tokenizer.encode(*t, true).unwrap())
        .collect();
    let max_len = encodings.iter().map(|e| e.get_ids().len()).max().unwrap_or(0);
    let batch_size = texts.len();

    let mut ids_flat: Vec<i64> = Vec::with_capacity(batch_size * max_len);
    let mut mask_flat: Vec<i64> = Vec::with_capacity(batch_size * max_len);
    for enc in &encodings {
        let ids = enc.get_ids();
        for &id in ids { ids_flat.push(id as i64); }
        for _ in 0..ids.len() { mask_flat.push(1); }
        for _ in ids.len()..max_len { ids_flat.push(0); mask_flat.push(0); }
    }

    let input_ids = TensorRef::from_array_view(
        (vec![batch_size as i64, max_len as i64], &ids_flat[..]),
    ).unwrap();
    let attention_mask = TensorRef::from_array_view(
        (vec![batch_size as i64, max_len as i64], &mask_flat[..]),
    ).unwrap();

    let _outputs = session.run(inputs![input_ids, attention_mask]).unwrap();
}

fn run_inference(session: &mut Session, tokenizer: &tokenizers::Tokenizer, text: &str) {
    let encoding = tokenizer.encode(text, true).unwrap();
    let ids: Vec<i64> = encoding.get_ids().iter().map(|&id| id as i64).collect();
    let mask: Vec<i64> = vec![1i64; ids.len()];
    let seq_len = ids.len();

    let input_ids = TensorRef::from_array_view(
        (vec![1i64, seq_len as i64], &ids[..]),
    ).unwrap();
    let attention_mask = TensorRef::from_array_view(
        (vec![1i64, seq_len as i64], &mask[..]),
    ).unwrap();

    let _outputs = session.run(inputs![input_ids, attention_mask]).unwrap();
}
