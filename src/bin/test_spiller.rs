fn main() {
    let base = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let dict = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let analyzer = mtag::Analyzer::new(dict.to_str().unwrap()).expect("load");
    for word in &["spiller", "spille", "liker", "å"] {
        println!("--- {} ---", word);
        let tokens = analyzer.analyze(word);
        for t in &tokens {
            for r in &t.readings {
                println!("  {} {:?} {:?}", r.lemma, r.pos, r.tags);
            }
        }
    }
}
