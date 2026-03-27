/// Test phonetic prefix matching in the completion pipeline.
/// Calls complete_word() directly — the EXACT same function the app uses.
/// Usage: cargo run --release --bin test_completion_phonetic

use std::collections::HashMap;
use std::path::PathBuf;

fn main() -> anyhow::Result<()> {
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let data = base.join("../contexter-repo/training-data");
    let onnx = data.join("onnx/norbert4_base_int8.onnx");
    let tok = data.join("onnx/tokenizer.json");
    let wf_path = data.join("wordfreq.tsv");
    let dict_path = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");

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

    // (context, prefix, expected_words, description)
    // expected_words: at least ONE must appear in top 10 results
    let tests: Vec<(&str, &str, Vec<&str>, &str)> = vec![
        // Phonetic: å → o
        ("Maten er ", "gå", vec!["godt", "god", "gode"], "å→o: gå should find godt/god"),
        ("Det var veldig ", "gå", vec!["godt", "god"], "å→o: gå in adj context"),
        // Phonetic: ø → e
        ("Han kom hit ", "fer", vec!["før", "først", "første"], "ø→e: fer should find før"),
        // Phonetic: æ → a
        ("Jeg vil ", "lare", vec!["lære", "læreren"], "æ→a: lare should find lære"),
        // Phonetic: sj → kj (via prefix first-char already works, but test it)
        ("Vi har et fint ", "kjø", vec!["kjøkken", "kjøkkenet"], "kjø finds kjøkken"),
        // Baseline: exact prefix still works (no regression)
        ("Jeg har ", "go", vec!["godt", "god", "gode"], "baseline: go finds godt"),
        ("Fotball er en ", "mo", vec!["morsom", "morsomt"], "baseline: mo finds morsom"),
        ("Hun liker å ", "le", vec!["lese", "leke"], "baseline: le finds lese"),
        // Phonetic: o → å (reverse direction)
        ("Han gjør det ", "go", vec!["godt", "god"], "o→å: go still finds godt (baseline + phonetic)"),
    ];

    let mut pass = 0;
    let mut fail = 0;

    for (context, prefix, expected, desc) in &tests {
        let results = nostos_cognio::complete::complete_word(
            &mut model, context, prefix, &pi,
            None, Some(&wf), Some(&fallback_dict), Some(&fallback_prefix), None,
            1.0, 10.0, 10, 3,
        )?;

        let result_words: Vec<String> = results.iter().map(|c| c.word.to_lowercase()).collect();
        let found = expected.iter().any(|exp| result_words.iter().any(|r| r == &exp.to_lowercase()));

        let top5: Vec<String> = results.iter().take(5).map(|c| format!("{}({:.1})", c.word, c.score)).collect();

        if found {
            println!("  PASS: {} → [{}]", desc, top5.join(", "));
            pass += 1;
        } else {
            println!("  FAIL: {} → [{}] (expected one of {:?})", desc, top5.join(", "), expected);
            fail += 1;
        }
    }

    println!("\nResults: {}/{} passed", pass, pass + fail);
    if fail > 0 { std::process::exit(1); }
    Ok(())
}
