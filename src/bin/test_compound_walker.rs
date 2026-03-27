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
        // Single word, no compound
        ("kjøkken", vec!["kjøkken"], "exact single word"),
        ("sjøkken", vec!["kjøkken"], "fuzzy single word (s→k)"),

        // Two-part compound, exact
        ("kjøkkenbord", vec!["kjøkkenbord"], "exact compound"),

        // Two-part compound, error in first part
        ("sjøkkenbord", vec!["kjøkkenbord"], "error in part 1 (s→k)"),

        // Two-part compound, error in second part
        ("kjøkkenbort", vec!["kjøkkenbord"], "error in part 2 (t→d)"),

        // Two-part compound, errors in BOTH parts
        ("sjøkkenbort", vec!["kjøkkenbord"], "errors in both parts"),

        // Binding letter 's'
        ("arbeidsplass", vec!["arbeidsplass"], "exact with binding s"),

        // Phonetic confusion in compound
        ("fotballag", vec!["fotballag"], "football team"),
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
