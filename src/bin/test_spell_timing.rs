use std::time::Instant;
use acatts_rust::spelling_scorer::find_candidates_pipeline;
use acatts_rust::compound_walker::load_fst_from_mfst;
use std::collections::HashMap;

fn main() {
    let base = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let dict = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let analyzer = mtag::Analyzer::new(dict.to_str().unwrap()).expect("load");
    let compound_fst = load_fst_from_mfst(dict.to_str().unwrap()).expect("load fst");

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
        let candidates = find_candidates_pipeline(
            &analyzer, Some(&compound_fst), Some(&wf), &empty_user, &empty_doc,
            word, sentence, &language::BokmalLanguage,
        );
        let elapsed = t.elapsed();
        let top = candidates.first().map(|(w, s)| format!("{} ({:.2})", w, s)).unwrap_or("none".into());
        println!("{:20} → {:20} candidates={:3}  {:?}", word, top, candidates.len(), elapsed);
    }
}
