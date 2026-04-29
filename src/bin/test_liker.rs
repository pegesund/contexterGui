use nostos_cognio::grammar::swipl_checker::SwiGrammarChecker;
use std::path::PathBuf;

fn main() {
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let dict = base.join("../rustSpell/mtag-rs/data/fullform_nn.mfst");
    let grammar = base.join("../nynorsk/grammar_rules.pl");
    let dir = base.join("../nynorsk");
    let swipl = "/Applications/SWI-Prolog.app/Contents/Frameworks/libswipl.dylib";

    let mut checker = SwiGrammarChecker::new(
        swipl, dict.to_str().unwrap(), grammar.to_str().unwrap(), dir.to_str().unwrap(),
    ).expect("Failed to create checker");

    for sentence in &[
        "Eg liker å spille fotball når det er sol.",
        "Eg likar å spele fotball når det er sol.",
        "Eg liker sjokolade.",
        "Eg likar sjokolade.",
    ] {
        let errors = checker.check_sentence(sentence);
        println!("\n'{}'", sentence);
        if errors.is_empty() {
            println!("  (no errors)");
        } else {
            for e in &errors {
                println!("  ERROR: word='{}' rule='{}' suggestion='{}'", e.word, e.rule_name, e.suggestion);
                println!("    {}", e.explanation);
            }
        }
    }
}
