/// Test the compound FST walker.
/// Loads the Norwegian FST dictionary and tests compound word decomposition
/// with fuzzy matching per part.

use acatts_rust::compound_walker::{compound_fuzzy_walk, load_fst_from_mfst};
use std::path::PathBuf;
use std::time::Instant;

fn main() {
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let mfst_path = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let dict_path = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");

    println!("Loading FST from {}...", mfst_path.display());
    let t = Instant::now();
    let fst = load_fst_from_mfst(mfst_path.to_str().unwrap())
        .expect("Failed to load FST");
    println!("Loaded in {:?}", t.elapsed());

    let analyzer = mtag::Analyzer::new(dict_path.to_str().unwrap())
        .expect("Failed to load analyzer");

    // Check which compound words are already in the dictionary
    println!("\n=== Dictionary check ===");
    let check_words = vec![
        "kjøkkenbord", "fotballkamp", "arbeidsplass", "skolegård",
        "frokostbord", "middagspris", "vårtilbud", "sommervarme",
        "skrivebord", "soverom", "togstasjon", "busstopp",
        "sykkelvei", "barneskole", "matbutikk", "bilverksted",
    ];
    for w in &check_words {
        let in_fst = analyzer.has_word(w);
        println!("  {:20} {}", w, if in_fst { "✓ IN dictionary" } else { "✗ NOT in dictionary" });
    }
    println!();

    // (input, expected_in_results, description)
    let tests: Vec<(&str, Vec<&str>, &str)> = vec![
        // === Single words (baseline) ===
        ("kjøkken", vec!["kjøkken"], "exact single word"),
        ("sjøkken", vec!["kjøkken"], "fuzzy single: s→k"),
        ("bord", vec!["bord"], "short word, exact"),
        ("bort", vec!["bord", "bort"], "short word, fuzzy"),

        // === Two-part compounds, exact ===
        ("kjøkkenbord", vec!["kjøkkenbord"], "exact: kjøkken+bord"),
        ("fotballkamp", vec!["fotballkamp"], "exact: fotball+kamp"),
        ("arbeidsplass", vec!["arbeidsplass"], "exact: arbeid+s+plass"),
        ("skolegård", vec!["skolegård"], "exact: skole+gård"),

        // === Error in first part ===
        ("sjøkkenbord", vec!["kjøkkenbord"], "part1: sj→kj in kjøkkenbord"),
        ("fotbalskamp", vec!["fotballkamp"], "part1: missing l in fotball"),
        ("skollegård", vec!["skolegård"], "part1: ll→l in skole"),

        // === Error in second part ===
        ("kjøkkenbort", vec!["kjøkkenbord"], "part2: t→d in bord"),
        ("fotballkamb", vec!["fotballkamp"], "part2: b→p in kamp"),
        ("skolegårt", vec!["skolegård"], "part2: t→d in gård"),

        // === Errors in BOTH parts ===
        ("sjøkkenbort", vec!["kjøkkenbord"], "both: sj→kj + t→d"),

        // === Binding letter 's' ===
        ("arbeidsplasss", vec!["arbeidsplass"], "binding s: extra s"),

        // === Phonetic å↔o/u in compounds ===
        ("gåttebord", vec!["guttebord"], "phonetic: å→u in gutte"),
        ("lekeplaas", vec!["lekeplass"], "missing s in plass"),
        ("barnehagge", vec!["barnehage"], "double consonant: gg→g"),

        // === Productive compounds NOT in dictionary ===
        ("frokostbort", vec!["frokostbord"], "productive: frokost+bord t→d"),
        ("middagspris", vec!["middagspris"], "productive: middag+s+pris"),
        ("vårtilbud", vec!["vårtilbud"], "productive: vår+tilbud"),
        ("sommervarme", vec!["sommervarme"], "productive: sommer+varme"),
        ("skrivebort", vec!["skrivebord"], "productive: skrive+bord t→d"),
        ("togstassjon", vec!["togstasjon"], "productive: tog+stasjon ss→s"),
        ("busstop", vec!["busstopp"], "productive: buss+topp missing p"),
        ("sykkelveien", vec!["sykkelveien"], "productive: sykkel+veien"),
        ("barneskole", vec!["barneskole"], "productive: barne+skole"),
        ("matbuttikk", vec!["matbutikk"], "productive: mat+butikk tt→t"),

        // === Dyslexic-style errors ===
        ("sjokkolade", vec!["sjokolade"], "double k: kk→k"),
        ("informassjon", vec!["informasjon"], "double s: ss→s"),
        ("datamaskinn", vec!["datamaskin"], "extra n at end"),

        // === Three-part compounds ===
        ("barnehageplass", vec!["barnehageplass"], "three-part: barne+hage+plass"),
    ];

    let mut pass = 0;
    let mut fail = 0;

    for (input, expected, desc) in &tests {
        let t = Instant::now();
        let results = compound_fuzzy_walk(&fst, &input.to_lowercase());
        let elapsed = t.elapsed();

        let result_words: Vec<&str> = results.iter().map(|r| r.compound_word.as_str()).collect();
        let found = expected.iter().any(|exp| result_words.contains(exp));

        let top3: Vec<String> = results.iter().take(3)
            .map(|r| {
                let parts: Vec<String> = r.parts.iter()
                    .map(|p| format!("{}({})", p.matched_word, p.edits))
                    .collect();
                format!("{}[{}] e={}", r.compound_word, parts.join("+"), r.total_edits)
            })
            .collect();

        if found {
            println!("  PASS ({:>5?}): {} → [{}]", elapsed, desc, top3.join(", "));
            pass += 1;
        } else {
            println!("  FAIL ({:>5?}): {} — input='{}' got [{}] expected {:?}",
                elapsed, desc, input, top3.join(", "), expected);
            fail += 1;
        }
    }

    println!("\nResults: {}/{} passed", pass, pass + fail);
    if fail > 0 { std::process::exit(1); }
}
