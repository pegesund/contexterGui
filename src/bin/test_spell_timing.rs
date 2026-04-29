use std::time::Instant;
use acatts_rust::spelling_scorer::generate_spelling_candidates;
use std::collections::HashMap;

fn main() {
    let base = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let dict = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let analyzer = mtag::Analyzer::new(dict.to_str().unwrap()).expect("load");

    let wf_path = base.join("../contexter-repo/training-data/wordfreq.tsv");
    let wf = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);

    let empty_user: Vec<String> = vec![];
    let empty_doc: HashMap<String, u16> = HashMap::new();

    let cases = vec![
        ("Jeg spilflfler fotcaball.", "spilflfler"),
        ("Jeg spilflfler fotcaball.", "fotcaball"),
        ("Hun herte ikke.", "herte"),
        ("Vi lære norsk.", "lære"),
        ("Han spisle middag.", "spisle"),
    ];

    for (sentence, word) in &cases {
        let t = Instant::now();
        let candidates = generate_spelling_candidates(
            &analyzer, Some(&wf), &empty_user, &empty_doc,
            word, sentence, &language::BokmalLanguage, None,
        );
        let elapsed = t.elapsed();
        let top = candidates.first().map(|(w, s)| format!("{} ({:.2})", w, s)).unwrap_or("none".into());
        println!("{:20} → {:20} candidates={:3}  {:?}", word, top, candidates.len(), elapsed);
    }
}
