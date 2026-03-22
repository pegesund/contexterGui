use std::path::PathBuf;

mod user_dict;

// Levenshtein distance (same as in main.rs)
fn levenshtein_distance(a: &str, b: &str) -> u32 {
    let (a, b): (Vec<char>, Vec<char>) = (a.chars().collect(), b.chars().collect());
    let (m, n) = (a.len(), b.len());
    let mut dp = vec![vec![0u32; n + 1]; m + 1];
    for i in 0..=m { dp[i][0] = i as u32; }
    for j in 0..=n { dp[0][j] = j as u32; }
    for i in 1..=m {
        for j in 1..=n {
            let cost = if a[i-1] == b[j-1] { 0 } else { 1 };
            dp[i][j] = (dp[i-1][j] + 1).min(dp[i][j-1] + 1).min(dp[i-1][j-1] + cost);
        }
    }
    dp[m][n]
}

fn main() {
    let dict_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../rustSpell/mtag-rs/data/fullform_bm.mfst");

    // Load analyzer
    eprintln!("Loading analyzer...");
    let analyzer = mtag::Analyzer::new(&dict_path).expect("Failed to load dictionary");
    eprintln!("Ready. Dict size: {}\n", analyzer.dict_size());

    // Open user dict in temp location for test
    let db_path = std::env::temp_dir().join("test_user_words.db");
    let _ = std::fs::remove_file(&db_path); // start fresh
    let udict = user_dict::UserDict::open(&db_path).expect("Failed to open user dict");

    let mut pass = 0;
    let mut fail = 0;

    // --- Test 1: "nevrale" is unknown in standard dict ---
    {
        let known = analyzer.has_word("nevrale");
        if !known {
            println!("PASS: 'nevrale' is unknown in standard dictionary");
            pass += 1;
        } else {
            println!("FAIL: 'nevrale' should be unknown in standard dictionary");
            fail += 1;
        }
    }

    // --- Test 2: "nevrale" is not in user dict yet ---
    {
        if !udict.has_word("nevrale") {
            println!("PASS: 'nevrale' not in user dict yet");
            pass += 1;
        } else {
            println!("FAIL: 'nevrale' should not be in user dict yet");
            fail += 1;
        }
    }

    // --- Test 3: Add "nevrale" to user dict ---
    {
        udict.add_word("nevrale").expect("Failed to add word");
        if udict.has_word("nevrale") {
            println!("PASS: 'nevrale' added to user dict");
            pass += 1;
        } else {
            println!("FAIL: 'nevrale' should be in user dict after add");
            fail += 1;
        }
    }

    // --- Test 4: Combined lookup (analyzer + user dict) ---
    {
        let known = analyzer.has_word("nevrale") || udict.has_word("nevrale");
        if known {
            println!("PASS: 'nevrale' found via combined lookup");
            pass += 1;
        } else {
            println!("FAIL: 'nevrale' should be found via combined lookup");
            fail += 1;
        }
    }

    // --- Test 5: Wildcard readings cover key POS categories ---
    {
        let readings = user_dict::UserDict::wildcard_readings("nevrale");
        let has_noun = readings.iter().any(|r| r.pos == mtag::types::Pos::Subst);
        let has_adj = readings.iter().any(|r| r.pos == mtag::types::Pos::Adj);
        let has_verb = readings.iter().any(|r| r.pos == mtag::types::Pos::Verb);
        let has_adv = readings.iter().any(|r| r.pos == mtag::types::Pos::Adv);
        let all_normert = readings.iter().all(|r| r.tags.contains(&mtag::types::Tag::Normert));
        if has_noun && has_adj && has_verb && has_adv && all_normert {
            println!("PASS: wildcard readings cover noun/adj/verb/adv, all Normert");
            pass += 1;
        } else {
            println!("FAIL: wildcard readings incomplete (noun={} adj={} verb={} adv={} normert={})",
                has_noun, has_adj, has_verb, has_adv, all_normert);
            fail += 1;
        }
        println!("  {} readings generated:", readings.len());
        for r in &readings {
            println!("    {} {} {:?}", r.lemma, r.pos, r.tags);
        }
    }

    // --- Test 6: "fotball" is known — user dict doesn't interfere ---
    {
        let known = analyzer.has_word("fotball");
        if known {
            println!("PASS: 'fotball' still known in standard dict");
            pass += 1;
        } else {
            println!("FAIL: 'fotball' should be known");
            fail += 1;
        }
    }

    // --- Test 7: Remove word from user dict ---
    {
        udict.remove_word("nevrale").expect("Failed to remove");
        if !udict.has_word("nevrale") {
            println!("PASS: 'nevrale' removed from user dict");
            pass += 1;
        } else {
            println!("FAIL: 'nevrale' should be gone after remove");
            fail += 1;
        }
    }

    // --- Test 8: Case insensitive ---
    {
        udict.add_word("Nevrale").expect("Failed to add");
        if udict.has_word("nevrale") && udict.has_word("NEVRALE") && udict.has_word("Nevrale") {
            println!("PASS: case-insensitive lookup works");
            pass += 1;
        } else {
            println!("FAIL: case-insensitive lookup broken");
            fail += 1;
        }
        udict.remove_word("nevrale").unwrap();
    }

    // --- Test 9: list_words ---
    {
        udict.add_word("nevrale").unwrap();
        udict.add_word("tensorflow").unwrap();
        udict.add_word("kubernetes").unwrap();
        let words = udict.list_words();
        if words.len() == 3 && words.contains(&"nevrale".to_string()) {
            println!("PASS: list_words returns {} words", words.len());
            pass += 1;
        } else {
            println!("FAIL: list_words returned {:?}", words);
            fail += 1;
        }
    }

    // --- Test 10: Grammar checker with user word ---
    // Load SWI-Prolog and check that "nevrale" in a sentence doesn't
    // appear as unknown when we filter with user dict
    {
        let swipl_path = "/Applications/SWI-Prolog.app/Contents/Frameworks/libswipl.dylib";
        let grammar_rules = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../syntaxer/grammar_rules.pl");
        let syntaxer_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../syntaxer");

        match nostos_cognio::grammar::swipl_checker::SwiGrammarChecker::new(
            swipl_path,
            dict_path.to_str().unwrap(),
            grammar_rules.to_str().unwrap(),
            syntaxer_dir.to_str().unwrap(),
        ) {
            Ok(mut checker) => {
                let result = checker.check_sentence_full("Lærer som et barn med nevrale nettverk.");

                // Without user dict: "nevrale" should be in unknown_words
                let unknown_before: Vec<&str> = result.unknown_words.iter()
                    .map(|u| u.word.as_str()).collect();
                let has_nevrale = unknown_before.contains(&"nevrale");

                // With user dict filter: remove user words from unknowns
                let filtered: Vec<&str> = result.unknown_words.iter()
                    .filter(|u| !udict.has_word(&u.word))
                    .map(|u| u.word.as_str()).collect();
                let nevrale_gone = !filtered.contains(&"nevrale");

                if has_nevrale && nevrale_gone {
                    println!("PASS: grammar checker — 'nevrale' unknown before filter, gone after");
                    pass += 1;
                } else {
                    println!("FAIL: grammar checker — before={:?} after={:?}", unknown_before, filtered);
                    fail += 1;
                }
            }
            Err(e) => {
                println!("SKIP: SWI-Prolog not available ({})", e);
            }
        }
    }

    // --- Test 11: Prefix matching for completions ---
    {
        // "nevrale" is in user dict, "tensorflow", "kubernetes" too
        let prefix_matches = |prefix: &str| -> Vec<String> {
            udict.list_words().into_iter()
                .filter(|w| w.starts_with(&prefix.to_lowercase()))
                .collect()
        };
        let matches = prefix_matches("nevr");
        if matches == vec!["nevrale".to_string()] {
            println!("PASS: prefix 'nevr' matches 'nevrale'");
            pass += 1;
        } else {
            println!("FAIL: prefix 'nevr' should match 'nevrale', got {:?}", matches);
            fail += 1;
        }
        let matches = prefix_matches("ten");
        if matches == vec!["tensorflow".to_string()] {
            println!("PASS: prefix 'ten' matches 'tensorflow'");
            pass += 1;
        } else {
            println!("FAIL: prefix 'ten' should match 'tensorflow', got {:?}", matches);
            fail += 1;
        }
        let matches = prefix_matches("xyz");
        if matches.is_empty() {
            println!("PASS: prefix 'xyz' matches nothing");
            pass += 1;
        } else {
            println!("FAIL: prefix 'xyz' should match nothing, got {:?}", matches);
            fail += 1;
        }
    }

    // --- Test 14: Levenshtein spelling candidates from user dict ---
    {
        // "nevrle" (missing 'a') → "nevrale" should be within distance 2
        let dist = levenshtein_distance("nevrle", "nevrale");
        if dist <= 2 {
            println!("PASS: 'nevrle' → 'nevrale' distance={} (within 2)", dist);
            pass += 1;
        } else {
            println!("FAIL: 'nevrle' → 'nevrale' distance={} (should be ≤2)", dist);
            fail += 1;
        }

        // "tensorflow" vs "fotbollx" — should NOT match (too far)
        let dist = levenshtein_distance("fotbollx", "tensorflow");
        if dist > 2 {
            println!("PASS: 'fotbollx' → 'tensorflow' distance={} (too far)", dist);
            pass += 1;
        } else {
            println!("FAIL: 'fotbollx' → 'tensorflow' distance={} (should be >2)", dist);
            fail += 1;
        }

        // Simulate spelling candidate selection from user dict
        let misspelled = "nevrle";
        let user_candidates: Vec<String> = udict.list_words().into_iter()
            .filter(|w| levenshtein_distance(misspelled, w) <= 2)
            .collect();
        if user_candidates == vec!["nevrale".to_string()] {
            println!("PASS: spelling candidates for 'nevrle' = {:?}", user_candidates);
            pass += 1;
        } else {
            println!("FAIL: spelling candidates for 'nevrle' should be ['nevrale'], got {:?}", user_candidates);
            fail += 1;
        }
    }

    // Cleanup
    let _ = std::fs::remove_file(&db_path);

    println!("\n=== Results: {} passed, {} failed ===", pass, fail);
    std::process::exit(if fail > 0 { 1 } else { 0 });
}
