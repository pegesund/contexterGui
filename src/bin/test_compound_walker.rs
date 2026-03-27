/// Test the compound FST walker.
/// Loads the Norwegian FST dictionary and tests compound word decomposition
/// with fuzzy matching per part.

use acatts_rust::compound_walker::{compound_fuzzy_walk, load_fst_from_mfst};
use std::path::PathBuf;
use std::time::Instant;

fn main() {
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let mfst_path = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");

    println!("Loading FST from {}...", mfst_path.display());
    let t = Instant::now();
    let fst = load_fst_from_mfst(mfst_path.to_str().unwrap())
        .expect("Failed to load FST");
    println!("Loaded in {:?}\n", t.elapsed());

    // (input, expected_in_results, description)
    let tests: Vec<(&str, Vec<&str>, &str)> = vec![
        // === Single words (baseline) ===
        ("kjøkken", vec!["kjøkken"], "exact single word"),
        ("sjøkken", vec!["kjøkken"], "fuzzy single: s→k"),

        // === Two-part compounds, exact ===
        ("kjøkkenbord", vec!["kjøkkenbord"], "exact compound"),
        ("fotballkamp", vec!["fotballkamp"], "exact: fotball+kamp"),
        ("arbeidsplass", vec!["arbeidsplass"], "exact with binding s"),
        ("skolegård", vec!["skolegård"], "exact: skole+gård"),

        // === Error in first part ===
        ("sjøkkenbord", vec!["kjøkkenbord"], "part1 error: sj→kj"),
        ("fotbalskamp", vec!["fotballkamp"], "part1 error: missing l"),
        ("skollegård", vec!["skolegård"], "part1 error: ll→l"),

        // === Error in second part ===
        ("kjøkkenbort", vec!["kjøkkenbord"], "part2 error: t→d"),
        ("fotballkamb", vec!["fotballkamp"], "part2 error: b→p"),
        ("skolegårt", vec!["skolegård"], "part2 error: t→d"),

        // === Errors in BOTH parts ===
        ("sjøkkenbort", vec!["kjøkkenbord"], "both parts: sj→kj + t→d"),
        ("fotbalskamb", vec!["fotballkamp"], "both parts: l missing + b→p"),

        // === Binding letter 's' ===
        ("arbeidsplas", vec!["arbeidsplass"], "binding s: missing final s"),
        ("arbeidsplasss", vec!["arbeidsplass"], "binding s: extra s"),

        // === Phonetic confusions in compounds ===
        ("gåttebord", vec!["guttebord"], "phonetic: å→u in part1"),
        ("lekeplaas", vec!["lekeplass"], "phonetic: missing s"),
        ("barnehagge", vec!["barnehage"], "double consonant: gg→g"),

        // === Common Norwegian compound misspellings ===
        ("datamaskin", vec!["datamaskin"], "exact: data+maskin"),
        ("datamaskinn", vec!["datamaskin"], "extra n at end"),
        ("helsesøster", vec!["helsesøster"], "exact: helse+søster"),
        ("husholding", vec!["husholdning"], "missing n: holdning"),

        // === Three-part compounds ===
        ("barnehageplass", vec!["barnehageplass"], "three-part: barne+hage+plass"),

        // === Dyslexic-style errors ===
        ("sjokkolade", vec!["sjokolade"], "double k: kk→k"),
        ("biblåtek", vec!["bibliotek"], "phonetic: io→å"),
        ("restourang", vec!["restaurant"], "phonetic: au→ou"),
        ("informassjon", vec!["informasjon"], "double s: ss→s"),

        // === Edge cases ===
        ("bord", vec!["bord"], "short word, exact"),
        ("bort", vec!["bord", "bort"], "short word, fuzzy"),
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
