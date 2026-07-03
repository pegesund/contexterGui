fn main() {
    let base = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let dict = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let analyzer = mtag::Analyzer::new(dict.to_str().unwrap()).expect("load");
    let fst = acatts_rust::compound_walker::load_fst_from_mfst(dict.to_str().unwrap()).unwrap();

    let words = vec!["spilflfler", "fotcaball", "fotaaball", "karrierekompasset",
                     "maskinlæringsalgoritme", "fotball", "spiller", "hus",
                     "sneøen", "vintrern", "skii", "liikkker"];
    for w in &words {
        let t = std::time::Instant::now();
        let results = acatts_rust::compound_walker::compound_fuzzy_walk(
            &fst, w, &language::BokmalLanguage, None, None, None);
        let elapsed = t.elapsed();
        let exact = results.iter().any(|r| r.total_edits == 0);
        let has = analyzer.has_word(w);
        println!("{:30} has_word={:5} compound={:5} results={:3} {:?}", w, has, exact, results.len(), elapsed);
    }
}
