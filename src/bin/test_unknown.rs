use nostos_cognio::grammar::swipl_checker::SwiGrammarChecker;
use std::path::PathBuf;

fn main() {
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let dict = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let grammar = base.join("../syntaxer/grammar_rules.pl");
    let dir = base.join("../syntaxer");
    let swipl = if cfg!(target_os = "windows") {
        "C:\\Program Files\\swipl\\bin\\libswipl.dll"
    } else {
        "/Applications/SWI-Prolog.app/Contents/Frameworks/libswipl.dylib"
    };

    let mut checker = SwiGrammarChecker::new(
        swipl, dict.to_str().unwrap(), grammar.to_str().unwrap(), dir.to_str().unwrap(),
    ).expect("Failed to create checker");

    let sentences = vec![
        "Jeg spilflfler fotcaball.",
        "Jeg spiller fotaaball.",
        "Jeg spilller fotball.",
        "Jeg spiiller fotiball.",
        "Dette er en tets.",
        "Han skrivver bra.",
        "Nordmennene liikkker å gå på skii i sneøen om vintrern.",
        "Hun kjøpte et karrierekompasset til jobben.",
        "Fotballbanen ligger ved skolen.",
    ];

    // Same compound gate as the grammar actor uses.
    let analyzer = checker.analyzer().clone();
    let compound_check = |word: &str| -> bool {
        nostos_cognio::grammar::is_novel_compound(&analyzer, word)
    };

    for sentence in &sentences {
        let result = checker.check_sentence_full_with_compound(sentence, "", Some(&compound_check));
        println!("\n'{}'", sentence);
        println!("  Grammar errors: {}", result.errors.len());
        for e in &result.errors {
            println!("    rule='{}' word='{}' suggestion='{}'", e.rule_name, e.word, e.suggestion);
        }
        println!("  Unknown words: {}", result.unknown_words.len());
        for u in &result.unknown_words {
            println!("    '{}' suggestions={:?}", u.word, u.spelling_suggestions.iter().take(3).collect::<Vec<_>>());
        }
    }
}
