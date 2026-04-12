fn main() {
    let base = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let dict = base.join("../rustSpell/mtag-rs/data/fullform_en.mfst");
    let analyzer = mtag::Analyzer::new(dict.to_str().unwrap()).expect("load");

    let words = vec!["came", "saw", "chose", "drank", "drove", "wore", "spoke",
                     "went", "got", "gone", "done", "seen", "run", "read",
                     "lost", "lay"];
    for w in &words {
        let tokens = analyzer.analyze(w);
        if let Some(t) = tokens.first() {
            println!("{}:", w);
            for r in &t.readings {
                println!("  {} {:?} {:?}", r.lemma, r.pos, r.tags);
            }
        } else {
            println!("{}: no readings", w);
        }
    }
}
