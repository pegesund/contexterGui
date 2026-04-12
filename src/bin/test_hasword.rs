fn main() {
    let base = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let dict = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let analyzer = mtag::Analyzer::new(dict.to_str().unwrap()).expect("load");
    
    let words = vec!["spilflfler", "fotcaball", "fotaaball", "spilller", "spiiller",
                     "fotiball", "tets", "skrivver", "fotball", "spiller",
                     "spillflir", "spillflør", "fotaball", "fotoaball"];
    for w in &words {
        let has = analyzer.has_word(w);
        let readings = analyzer.dict_lookup(w);
        let count = readings.as_ref().map_or(0, |r| r.len());
        let normert = readings.as_ref().map_or(false, |rs| 
            rs.iter().any(|r| r.tags.contains(&mtag::types::Tag::Normert)));
        println!("{:15} has_word={:5} readings={} normert={}", w, has, count, normert);
    }
}
