/// Benchmark: sentence-only vs full-paragraph context for completions.
/// Measures timing and result quality differences.

use std::path::PathBuf;
use std::time::Instant;

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

    // Test cases: (paragraph, sentence_only, prefix, description)
    let tests: Vec<(&str, &str, &str, &str)> = vec![
        (
            "Hei, dette er Petter som skriver. Petter er en glad gutt som liker å lese bøker og tidsskrifter. Han liker også å",
            "Han liker også å",
            "s",
            "short prefix 's' after long paragraph",
        ),
        (
            "Den hete poteten i skole-Norge om dagen er jo knyttet til nettilgang og bruk av LLM-er. Det viser seg at nettverkene i skolene er helt åpne. Vi har en fleksibel teknologi der lærerne kan åpne og lukke for programmer i sanntid. Hver gang vi presenterer for lærere får vi tilbakemelding på at det er akkurat dette skolen",
            "Hver gang vi presenterer for lærere får vi tilbakemelding på at det er akkurat dette skolen",
            "tr",
            "prefix 'tr' in long document about school tech",
        ),
        (
            "Fotball er en morsom idrett. Jeg spiller fotball hver dag. Laget vårt vant kampen i går. Nå skal vi",
            "Nå skal vi",
            "tre",
            "prefix 'tre' with football context",
        ),
        (
            "Vi har startet et forskningsprosjekt sammen med UiO. Tilbakemeldingen fra forskerne er at dette er hyper-relevant. Men vår utfordring er at vi når ikke gjennom til",
            "Men vår utfordring er at vi når ikke gjennom til",
            "by",
            "prefix 'by' in formal text",
        ),
        (
            "Petter liker å lese.",
            "Petter liker å lese.",
            "b",
            "short paragraph, prefix 'b'",
        ),
    ];

    let n_runs = 5; // average over N runs

    println!("{:<45} {:>10} {:>10} {:>8} {}", "Test", "Sentence", "Paragraph", "Diff", "Top results (sentence → paragraph)");
    println!("{}", "─".repeat(120));

    for (paragraph, sentence, prefix, desc) in &tests {
        // Warm up
        let _ = nostos_cognio::complete::complete_word(
            &mut model, sentence, prefix, &pi,
            None, Some(&wf), Some(&fallback_dict), Some(&fallback_prefix), None,
            1.0, 10.0, 5, 3,
        )?;

        // Benchmark sentence-only
        let mut sentence_ms = 0u128;
        let mut sentence_results = vec![];
        for _ in 0..n_runs {
            let t = Instant::now();
            let r = nostos_cognio::complete::complete_word(
                &mut model, sentence, prefix, &pi,
                None, Some(&wf), Some(&fallback_dict), Some(&fallback_prefix), None,
                1.0, 10.0, 5, 3,
            )?;
            sentence_ms += t.elapsed().as_millis();
            sentence_results = r;
        }
        let avg_sentence = sentence_ms / n_runs as u128;

        // Benchmark full paragraph
        let mut para_ms = 0u128;
        let mut para_results = vec![];
        for _ in 0..n_runs {
            let t = Instant::now();
            let r = nostos_cognio::complete::complete_word(
                &mut model, paragraph, prefix, &pi,
                None, Some(&wf), Some(&fallback_dict), Some(&fallback_prefix), None,
                1.0, 10.0, 5, 3,
            )?;
            para_ms += t.elapsed().as_millis();
            para_results = r;
        }
        let avg_para = para_ms / n_runs as u128;

        let diff = avg_para as i128 - avg_sentence as i128;
        let diff_str = if diff > 0 { format!("+{}ms", diff) } else { format!("{}ms", diff) };

        let sent_top: Vec<String> = sentence_results.iter().take(3).map(|c| format!("{}({:.0})", c.word, c.score)).collect();
        let para_top: Vec<String> = para_results.iter().take(3).map(|c| format!("{}({:.0})", c.word, c.score)).collect();

        println!("{:<45} {:>7}ms {:>7}ms {:>8} {} → {}",
            desc, avg_sentence, avg_para, diff_str,
            sent_top.join(", "), para_top.join(", "));
    }

    // Token count comparison
    println!("\n{}", "─".repeat(120));
    println!("Token counts (context → truncated to 32):\n");
    for (paragraph, sentence, _, desc) in &tests {
        let sent_tokens = model.tokenizer.encode(*sentence, false)
            .map(|e| e.get_ids().len()).unwrap_or(0);
        let para_tokens = model.tokenizer.encode(*paragraph, false)
            .map(|e| e.get_ids().len()).unwrap_or(0);
        println!("  {:<45} sentence: {:>3} tokens, paragraph: {:>3} tokens (truncated to 32)",
            desc, sent_tokens, para_tokens);
    }

    Ok(())
}
