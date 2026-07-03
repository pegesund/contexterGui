// Scratch test: show how compound_fuzzy_walk decomposes words that the
// grammar actor's is_compound hook accepts (total_edits == 0), using the
// exact same word_check/noun_check callbacks as grammar_actor.rs.
fn main() {
    let base = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let dict = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let analyzer = mtag::Analyzer::new(dict.to_str().unwrap()).expect("load");
    let fst = acatts_rust::compound_walker::load_fst_from_mfst(dict.to_str().unwrap()).unwrap();

    let word_check = |w: &str| -> bool { analyzer.has_word(w) };
    let noun_check = |w: &str| -> bool {
        analyzer.dict_lookup(w).map_or(false, |rs| {
            let n = rs.iter().filter(|r| r.pos == mtag::types::Pos::Subst).count();
            let a = rs.iter().filter(|r| r.pos == mtag::types::Pos::Adj).count();
            n > 0 && n >= a
        })
    };

    for w in ["sneøen", "vintrern", "karrierekompasset", "fotballbane"] {
        let results = acatts_rust::compound_walker::compound_fuzzy_walk(
            &fst, w, &language::BokmalLanguage, None, Some(&word_check), Some(&noun_check));
        println!("\n'{}':", w);
        for r in results.iter().filter(|r| r.total_edits == 0).take(8) {
            let parts: Vec<&str> = r.parts.iter().map(|p| p.matched_word.as_str()).collect();
            println!("  edits=0 parts={:?} joined='{}'", parts, r.compound_word);
        }
    }
}
