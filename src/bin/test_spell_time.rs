//! Time each source in find_spelling_suggestions for "hundder".
//! Same code as main.rs, extracted standalone.

use std::collections::{HashMap, HashSet};
use std::time::Instant;
use std::sync::Arc;

fn main() {
    let word = "hundder";
    let sentence_ctx = "Jeg liker ikke katter og hundder.";

    println!("=== Timing find_spelling_suggestions('{}') ===\n", word);

    // Load analyzer (mtag)
    let dict_path = "C:/Users/pette/dev/contexter/rustSpell/mtag-rs/data/fullform_bm.mfst";
    println!("Loading analyzer from {}...", dict_path);
    let t = Instant::now();
    let analyzer = mtag::Analyzer::new(dict_path).expect("Failed to load analyzer");
    println!("  Analyzer loaded in {}ms\n", t.elapsed().as_millis());

    // Load wordfreq
    let wf_path = "C:/Users/pette/dev/contexter/contexter-repo/training-data/wordfreq.tsv";
    println!("Loading wordfreq...");
    let t = Instant::now();
    let mut wordfreq: HashMap<String, u64> = HashMap::new();
    if let Ok(content) = std::fs::read_to_string(wf_path) {
        for line in content.lines() {
            let parts: Vec<&str> = line.split('\t').collect();
            if parts.len() >= 2 {
                if let Ok(freq) = parts[1].parse::<u64>() {
                    wordfreq.insert(parts[0].to_string(), freq);
                }
            }
        }
    }
    println!("  Wordfreq: {} words in {}ms\n", wordfreq.len(), t.elapsed().as_millis());

    // Load compound FST
    println!("Loading compound FST...");
    let t = Instant::now();
    let compound_fst = acatts_rust::compound_walker::load_fst_from_mfst(dict_path).ok();
    println!("  FST loaded in {}ms\n", t.elapsed().as_millis());

    let word_lower = word.to_lowercase();
    let word_first = word_lower.chars().next().unwrap_or(' ');

    fn trigrams(word: &str) -> Vec<String> {
        let chars: Vec<char> = word.chars().collect();
        if chars.len() < 3 { return vec![word.to_string()]; }
        (0..chars.len() - 2).map(|i| chars[i..i+3].iter().collect()).collect()
    }
    let word_trigrams = trigrams(&word_lower);

    let mut candidates: Vec<String> = Vec::new();
    let mut seen: HashSet<String> = HashSet::new();
    let mut edit_distances: HashMap<String, u32> = HashMap::new();

    // Source 1
    let t = Instant::now();
    for (w, dist) in analyzer.fuzzy_lookup(&word_lower, 2) {
        let wl = w.to_lowercase();
        if wl == word_lower || wl.len() < 2 { continue; }
        edit_distances.insert(wl.clone(), dist);
        if seen.insert(wl.clone()) { candidates.push(wl); }
    }
    println!("Source 1 (fuzzy lev): {}ms, {} candidates", t.elapsed().as_millis(), candidates.len());

    // Source 2
    let t = Instant::now();
    let before = candidates.len();
    for w in analyzer.prefix_lookup(&word_lower, 20) {
        let wl = w.to_lowercase();
        let extra = wl.len() as i32 - word_lower.len() as i32;
        if extra >= 1 && extra <= 3 {
            edit_distances.entry(wl.clone()).or_insert(extra as u32);
            if wl != word_lower && wl.len() >= 2 && seen.insert(wl.clone()) { candidates.push(wl); }
        }
    }
    println!("Source 2 (prefix): {}ms, +{} candidates", t.elapsed().as_millis(), candidates.len() - before);

    // Source 3
    let t = Instant::now();
    let before = candidates.len();
    let char_count = word_lower.chars().count();
    if char_count >= 3 {
        let end_byte = word_lower.char_indices().rev().next().map(|(i, _)| i).unwrap_or(0);
        let shorter = &word_lower[..end_byte];
        for w in analyzer.prefix_lookup(shorter, 20) {
            let wl = w.to_lowercase();
            let diff = (wl.len() as i32 - word_lower.len() as i32).unsigned_abs() + 1;
            edit_distances.entry(wl.clone()).or_insert(diff);
            if wl != word_lower && wl.len() >= 2 && seen.insert(wl.clone()) { candidates.push(wl); }
        }
    }
    println!("Source 3 (prefix-1): {}ms, +{} candidates", t.elapsed().as_millis(), candidates.len() - before);

    // Source 4
    let t = Instant::now();
    let before = candidates.len();
    for strip in 1..=2u32 {
        let chars: Vec<char> = word_lower.chars().collect();
        if chars.len() <= 3 + strip as usize { continue; }
        let truncated: String = chars[..chars.len() - strip as usize].iter().collect();
        for (w, dist) in analyzer.fuzzy_lookup(&truncated, 2) {
            let wl = w.to_lowercase();
            edit_distances.entry(wl.clone()).or_insert(dist + strip);
            if wl != word_lower && wl.len() >= 2 && seen.insert(wl.clone()) { candidates.push(wl); }
        }
    }
    println!("Source 4 (trunc fuzzy): {}ms, +{} candidates", t.elapsed().as_millis(), candidates.len() - before);

    // Source 6
    let t = Instant::now();
    let before = candidates.len();
    if let Some(split) = acatts_rust::spelling_scorer::try_split_function_word(&word_lower, &analyzer) {
        let sl = split.to_lowercase();
        if seen.insert(sl.clone()) { candidates.push(sl); }
    }
    println!("Source 6 (split func): {}ms, +{} candidates", t.elapsed().as_millis(), candidates.len() - before);

    // Source 7
    let t = Instant::now();
    let before = candidates.len();
    for (w, _freq) in wordfreq.iter() {
        let wl = w.to_lowercase();
        if wl == word_lower || seen.contains(&wl) { continue; }
        if wl.chars().next().unwrap_or(' ') != word_first { continue; }
        let w_tri = trigrams(&wl);
        let common = word_trigrams.iter().filter(|t| w_tri.contains(t)).count();
        if common >= 2 && seen.insert(wl.clone()) { candidates.push(wl); }
    }
    println!("Source 7 (wordfreq): {}ms, +{} candidates", t.elapsed().as_millis(), candidates.len() - before);

    // Source 10
    let t = Instant::now();
    let before = candidates.len();
    if word_lower.len() >= 3 {
        let rest = &word_lower[word_first.len_utf8()..];
        for c in "abcdefghijklmnopqrstuvwxyzæøå".chars() {
            if c == word_first { continue; }
            let candidate = format!("{}{}", c, rest);
            if analyzer.has_word(&candidate) && seen.insert(candidate.clone()) {
                edit_distances.insert(candidate.clone(), 1);
                candidates.push(candidate);
            }
        }
    }
    println!("Source 10 (first-char): {}ms, +{} candidates", t.elapsed().as_millis(), candidates.len() - before);

    // Source 11
    let t = Instant::now();
    let before = candidates.len();
    let mut s11_lookups = 0u64;
    let mut s11_forms = 0u64;
    let mut s11_lev_calls = 0u64;
    let mut s11_slow_bases: Vec<(String, u128)> = Vec::new();
    {
        use mtag::types::{Pos, Tag};
        let base_candidates: Vec<String> = candidates.clone();
        for base in &base_candidates {
            let t_base = Instant::now();
            if let Some(readings) = analyzer.dict_lookup(base) {
                s11_lookups += 1;
                for r in &readings {
                    if !matches!(r.pos, Pos::Subst) { continue; }
                    for tag in &[Tag::Be, Tag::Fl] {
                        let forms = analyzer.forms_for_lemma(&r.lemma, &Pos::Subst, tag);
                        s11_forms += forms.len() as u64;
                        for form in forms {
                            let fl = form.to_lowercase();
                            if fl != word_lower && fl.len() >= 2 && seen.insert(fl.clone()) {
                                s11_lev_calls += 1;
                                let dist = acatts_rust::spelling_scorer::levenshtein_distance(&word_lower, &fl);
                                if dist <= 4 {
                                    edit_distances.insert(fl.clone(), dist);
                                    candidates.push(fl);
                                }
                            }
                        }
                    }
                }
            }
            let base_ms = t_base.elapsed().as_millis();
            if base_ms > 100 {
                s11_slow_bases.push((base.clone(), base_ms));
            }
        }
    }
    println!("Source 11 (inflections): {}ms, +{} candidates", t.elapsed().as_millis(), candidates.len() - before);
    println!("  lookups={}, forms={}, lev_calls={}", s11_lookups, s11_forms, s11_lev_calls);
    if !s11_slow_bases.is_empty() {
        println!("  SLOW bases (>100ms):");
        for (base, ms) in &s11_slow_bases {
            println!("    '{}': {}ms", base, ms);
        }
    }

    // Source 13
    let t = Instant::now();
    let before = candidates.len();
    if word_lower.len() >= 7 {
        if let Some(ref fst) = compound_fst {
            let word_check = |w: &str| -> bool {
                analyzer.dict_lookup(w).map_or(false, |rs|
                    rs.iter().any(|r| r.pos != mtag::types::Pos::Prop))
            };
            let noun_check = |w: &str| -> bool {
                analyzer.dict_lookup(w).map_or(false, |rs| {
                    let n = rs.iter().filter(|r| r.pos == mtag::types::Pos::Subst).count();
                    let a = rs.iter().filter(|r| r.pos == mtag::types::Pos::Adj).count();
                    n > 0 && n >= a
                })
            };
            let results = acatts_rust::compound_walker::compound_fuzzy_walk(
                fst, &word_lower,
                Some(&wordfreq),
                Some(&word_check), Some(&noun_check),
            );
            for r in results.iter().take(10) {
                let cw = r.compound_word.to_lowercase();
                if seen.insert(cw.clone()) {
                    edit_distances.insert(cw.clone(), r.total_edits);
                    candidates.push(cw);
                }
            }
        }
    }
    println!("Source 13 (compound): {}ms, +{} candidates", t.elapsed().as_millis(), candidates.len() - before);

    println!("\nTotal: {} candidates", candidates.len());
}
