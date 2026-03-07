use std::path::PathBuf;

fn data_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../contexter-repo/training-data")
}

fn main() -> anyhow::Result<()> {
    let ort_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../onnxruntime/onnxruntime-win-x64-1.23.0/lib/onnxruntime.dll");
    unsafe { std::env::set_var("ORT_DYLIB_PATH", &ort_path); }

    let data = data_dir();
    let onnx = data.join("onnx/norbert4_base_int8.onnx");
    let tok = data.join("onnx/tokenizer.json");

    eprintln!("Loading NorBERT4...");
    let mut model = nostos_cognio::model::Model::load(onnx.to_str().unwrap(), tok.to_str().unwrap())?;
    let pi = nostos_cognio::prefix_index::build_prefix_index(&model.tokenizer);
    eprintln!("Ready.\n");

    let prefix = "m";
    let prefix_lower = prefix.to_lowercase();

    // ============================================================
    // TEST A: Python-style short context (what cognio_demo.py does)
    // ============================================================
    println!("=== TEST A: Python-style short context ===");
    let ctx = "Jeg liker å jobbe med andre";
    let masked_text = format!("{}<mask> .", ctx);
    println!("Masked: {}", masked_text);

    let (logits, ms) = model.single_forward(&masked_text)?;
    println!("Forward pass: {:.0}ms", ms);

    // Get all tokens matching prefix "m"
    let matches: Vec<(u32, String)> = pi.get(&prefix_lower).cloned().unwrap_or_default();
    println!("Tokens matching '{}': {}", prefix, matches.len());

    // Score and sort
    let mut token_scored: Vec<(u32, String, f32)> = matches.iter()
        .map(|(tid, word)| (*tid, word.clone(), logits[*tid as usize]))
        .collect();
    token_scored.sort_by(|a, b| b.2.partial_cmp(&a.2).unwrap_or(std::cmp::Ordering::Equal));

    // Show top 20
    println!("\nTop 20 tokens by logit:");
    for (tid, word, score) in token_scored.iter().take(20) {
        println!("  {:20} tid={:5} logit={:.2}", word, tid, score);
    }

    // Find "menneske" specifically
    if let Some((tid, word, score)) = token_scored.iter().find(|(_, w, _)| w.to_lowercase().contains("menneske")) {
        println!("\n'menneske' found: {} tid={} logit={:.2}", word, tid, score);
        let rank = token_scored.iter().position(|(_, w, _)| w == word).unwrap_or(999);
        println!("  Rank: #{}", rank + 1);
    } else {
        println!("\n'menneske' NOT FOUND in matching tokens!");
    }

    // Now do BPE extension like Python
    println!("\n--- BPE Extension (Python-style) ---");
    let n_candidates = token_scored.len().min(30);
    struct Candidate {
        token_ids: Vec<u32>,
        word: String,
        score: f32,
        done: bool,
    }
    let mut candidates: Vec<Candidate> = token_scored.iter()
        .take(n_candidates)
        .map(|(tid, word, score)| Candidate {
            token_ids: vec![*tid],
            word: word.clone(),
            score: *score,
            done: false,
        })
        .collect();

    for step in 0..3 {
        let to_extend: Vec<usize> = candidates.iter().enumerate()
            .filter(|(_, c)| !c.done)
            .map(|(i, _)| i)
            .collect();
        if to_extend.is_empty() { break; }

        // Python: f"{ctx} {accumulated}{mask} ."
        let batch_texts: Vec<String> = to_extend.iter()
            .map(|&i| {
                let accumulated = model.tokenizer
                    .decode(&candidates[i].token_ids, false)
                    .unwrap_or_default();
                let accumulated = accumulated.trim();
                format!("{} {}<mask> .", ctx, accumulated)
            })
            .collect();

        let (argmaxes, _) = model.batched_forward_argmax(&batch_texts)?;
        for (k, &i) in to_extend.iter().enumerate() {
            let best_id = argmaxes[k];
            let best_token = &model.id_to_token[best_id as usize];
            let is_continuation = !best_token.starts_with('Ġ')
                && !matches!(best_token.as_str(), "." | "," | "!" | "?" | ";" | ":");

            if is_continuation {
                candidates[i].token_ids.push(best_id);
                candidates[i].word = model.tokenizer
                    .decode(&candidates[i].token_ids, false)
                    .unwrap_or_default().trim().to_string();
            } else {
                candidates[i].done = true;
            }
        }
        println!("Step {}: {} candidates extended", step, to_extend.len());
    }

    println!("\nExtended candidates (top 30):");
    for c in &candidates {
        println!("  {:20} score={:.2} done={}", c.word, c.score, c.done);
    }

    // ============================================================
    // TEST B-F: Does adding context make it worse?
    // ============================================================
    let test_cases = vec![
        ("B: 1 sentence, glued mask",
         "Jeg liker å jobbe med andre<mask> ."),
        ("C: 2 sentences, glued mask",
         "Jeg utvikler meg innenfor idrett. Jeg liker å jobbe med andre<mask> ."),
        ("D: Full doc, glued mask",
         "I fritiden liker jeg å bade og fiske. Når jeg har tid, går jeg ned til vannet. Dette er feil. Jeg liker å spille fotball. Det er et morsomt spill. Jeg utvikler meg innenfor idrett. Jeg liker å jobbe med andre<mask> ."),
        ("E: Full doc, SPACE before mask",
         "I fritiden liker jeg å bade og fiske. Når jeg har tid, går jeg ned til vannet. Dette er feil. Jeg liker å spille fotball. Det er et morsomt spill. Jeg utvikler meg innenfor idrett. Jeg liker å jobbe med andre <mask> ."),
        ("F: Full doc, space+after text (GUI-style)",
         "I fritiden liker jeg å bade og fiske. Når jeg har tid, går jeg ned til vannet. Dette er feil. Jeg liker å spille fotball. Det er et morsomt spill. Jeg utvikler meg innenfor idrett. Jeg liker å jobbe med andre <mask> Det kan å skrives mye om."),
    ];

    for (label, masked_text) in &test_cases {
        println!("\n=== {} ===", label);
        println!("Masked: {}", masked_text);

        let (logits_t, ms_t) = model.single_forward(masked_text)?;
        println!("Forward: {:.0}ms", ms_t);

        let mut scored: Vec<(u32, String, f32)> = matches.iter()
            .map(|(tid, word)| (*tid, word.clone(), logits_t[*tid as usize]))
            .collect();
        scored.sort_by(|a, b| b.2.partial_cmp(&a.2).unwrap_or(std::cmp::Ordering::Equal));

        println!("Top 10:");
        for (_, word, score) in scored.iter().take(10) {
            println!("  {:20} logit={:.2}", word, score);
        }

        if let Some((_, word, score)) = scored.iter().find(|(_, w, _)| w == "mennesker") {
            let rank = scored.iter().position(|(_, w, _)| w == word).unwrap_or(999);
            println!("'mennesker': logit={:.2}, rank=#{}", score, rank + 1);
        } else {
            println!("'mennesker' NOT FOUND");
        }
    }

    Ok(())
}
