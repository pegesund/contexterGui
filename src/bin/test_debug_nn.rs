use nostos_cognio::grammar::swipl_checker::SwiGrammarChecker;
use std::path::PathBuf;

fn main() {
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let dict_path = base.join("../rustSpell/mtag-rs/data/fullform_nn.mfst");
    let grammar_path = base.join("../nynorsk/grammar_rules.pl");
    let syntaxer_dir = base.join("../nynorsk");
    let swipl_path = "/Applications/SWI-Prolog.app/Contents/Frameworks/libswipl.dylib";
    let mut checker = SwiGrammarChecker::new(swipl_path,
        dict_path.to_str().unwrap(),
        grammar_path.to_str().unwrap(),
        syntaxer_dir.to_str().unwrap()).expect("create");
    let analyzer = checker.analyzer();
    for w in &["symjer", "et", "kome", "går"] {
        let token = analyzer.analyze_word(w);
        println!("=== {} ===", w);
        for r in &token.readings {
            println!("  lemma={} pos={:?} tags={:?}", r.lemma, r.pos, r.tags);
        }
    }
    let sents = [
        "Vi kjøpte ein godt ven",
        "Det er ein dyr minutt",
        "Vi kjøpte ein kvitt bil",
    ];
    for s in &sents {
        let errs = checker.check_sentence(s);
        println!("'{}' -> {:?}", s, errs.iter().map(|e| &e.rule_name).collect::<Vec<_>>());
    }
}
