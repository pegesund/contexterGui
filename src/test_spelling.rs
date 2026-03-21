use std::path::PathBuf;

fn main() {
    let swipl_path = "/Applications/SWI-Prolog.app/Contents/Frameworks/libswipl.dylib";
    let dict_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let grammar_rules = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../syntaxer/grammar_rules.pl");
    let syntaxer_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../syntaxer");

    eprintln!("Loading SWI-Prolog + dictionary...");
    let mut checker = nostos_cognio::grammar::swipl_checker::SwiGrammarChecker::new(
        swipl_path,
        dict_path.to_str().unwrap(),
        grammar_rules.to_str().unwrap(),
        syntaxer_dir.to_str().unwrap(),
    ).expect("Failed to load");
    eprintln!("Ready.\n");

    let tests = vec![
        // (sentence, expected_unknown_words)
        ("Fotball er en morsom sport somx er veldig morsson.", vec!["somx", "morsson"]),
        ("Dettex er en test.", vec!["Dettex"]),
        ("Morsson er et fint ord.", vec!["Morsson"]),
        ("Hei dette er en test.", vec![]),
        ("Fotball er en morsom sport.", vec![]),
        ("Han gjørre noe galt.", vec!["gjørre"]),
        ("Jeg liker å spise matx og drikkx.", vec!["matx", "drikkx"]),
    ];

    // Debug: analyze full sentence tokens
    println!("=== Full sentence tokenization ===");
    let tokens = checker.analyzer().analyze("Fotball er en morsom sport somx er veldig morsson.");
    for (i, t) in tokens.iter().enumerate() {
        let has_normert = t.readings.iter().any(|r| r.tags.iter().any(|t| *t == mtag::types::Tag::Normert));
        let is_prop_only = t.readings.len() == 1 && t.readings[0].pos == mtag::types::Pos::Prop;
        println!("  [{}] '{}' len={} readings={} normert={} prop_only={} tags={:?}",
            i, t.wordform, t.wordform.len(), t.readings.len(), has_normert, is_prop_only,
            t.readings.iter().map(|r| format!("{:?}:{:?}", r.pos, r.tags)).collect::<Vec<_>>());
    }
    println!();

    // Direct call to find_unknown_words
    println!("=== Direct find_unknown_words ===");
    let unk = nostos_cognio::grammar::find_unknown_words_with_source(
        checker.analyzer(),
        &tokens,
        "Fotball er en morsom sport somx er veldig morsson."
    );
    println!("Without freq:  {:?}", unk.iter().map(|u| &u.word).collect::<Vec<_>>());

    // With wordfreq
    let wf_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../contexter-repo/training-data/wordfreq.tsv");
    let wf = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);
    let unk3 = nostos_cognio::grammar::find_unknown_words_with_freq(
        checker.analyzer(),
        &tokens,
        "Fotball er en morsom sport somx er veldig morsson.",
        Some(&wf)
    );
    println!("With freq(50): {:?}", unk3.iter().map(|u| &u.word).collect::<Vec<_>>());

    // Check compound splitting
    let splits = nostos_cognio::grammar::try_splits(checker.analyzer(), "morsson");
    println!("try_splits('morsson'): {:?}", splits);
    let splits2 = nostos_cognio::grammar::try_splits(checker.analyzer(), "somx");
    println!("try_splits('somx'): {:?}", splits2);
    let splits3 = nostos_cognio::grammar::try_splits(checker.analyzer(), "dettex");
    println!("try_splits('dettex'): {:?}", splits3);

    // Check frequencies of split parts
    let wf_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../contexter-repo/training-data/wordfreq.tsv");
    let wf = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);
    // Check specific words
    for word in &["osv", "osv.", "nevrale", "nevral", "nevralt", "neural", "nevrale"] {
        let token = checker.analyze_word(word);
        let has = checker.has_word(word);
        println!("  '{}': has_word={} readings={}", word, has, token.readings.len());
        for r in &token.readings {
            println!("    lemma='{}' pos={:?} tags={:?}", r.lemma, r.pos, r.tags);
        }
    }
    for word in &["osv", "nevrale", "mor", "son", "mors", "det", "tex", "ex", "sport", "fotball", "som"] {
        println!("  freq('{}') = {:?}", word, wf.get(*word));
    }
    println!();

    let mut pass = 0;
    let mut fail = 0;

    for (sentence, expected) in &tests {
        let result = checker.check_sentence_full(sentence);
        let found: Vec<String> = result.unknown_words.iter().map(|u| u.word.clone()).collect();

        let mut ok = true;
        for exp in expected {
            if !found.iter().any(|f| f.eq_ignore_ascii_case(exp)) {
                ok = false;
            }
        }

        if ok && found.len() >= expected.len() {
            pass += 1;
            println!("  PASS  '{}'", sentence);
            println!("        found: {:?}", found);
        } else {
            fail += 1;
            println!("  FAIL  '{}'", sentence);
            println!("        expected: {:?}", expected);
            println!("        found:    {:?}", found);

            // Also check: is the word in the dictionary?
            for exp in expected {
                let has = checker.has_word(&exp.to_lowercase());
                println!("        has_word('{}') = {}", exp.to_lowercase(), has);
                let token = checker.analyze_word(&exp.to_lowercase());
                for r in &token.readings {
                    println!("        reading: lemma='{}' pos={:?} tags={:?}", r.lemma, r.pos, r.tags);
                }
            }
        }
    }

    println!("\n{} passed, {} failed", pass, fail);
}
