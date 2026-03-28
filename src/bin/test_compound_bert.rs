/// Test compound walker + BERT re-ranking.
/// For each misspelled compound word, generates candidates via the walker,
/// then uses BERT subword_score to pick the best one in context.

use acatts_rust::compound_walker::{compound_fuzzy_walk, load_fst_from_mfst};
use acatts_rust::spelling_scorer::subword_score;
use nostos_cognio::model::Model;
use std::collections::HashSet;
use std::path::PathBuf;
use std::time::Instant;

/// Orthographic similarity between misspelled input and candidate compound.
/// Combines prefix match, suffix match (preserves user's intended inflection),
/// character trigram overlap, and length similarity.
fn ortho_score(input: &str, candidate: &str) -> f32 {
    let inp = input.as_bytes();
    let cand = candidate.as_bytes();
    let max_len = inp.len().max(cand.len()).max(1);

    // 1. Common prefix length — dyslexics get beginnings right
    let prefix_len = inp.iter().zip(cand.iter())
        .take_while(|(a, b)| a == b)
        .count();
    let prefix_ratio = prefix_len as f32 / max_len as f32;

    // 2. Common suffix length — preserves user's intended inflection
    // "fakturaggebyr" ends in "byr" → "fakturagebyr" (byr=match) beats "fakturagebyret" (ret=no match)
    let suffix_len = inp.iter().rev().zip(cand.iter().rev())
        .take_while(|(a, b)| a == b)
        .count();
    let suffix_ratio = suffix_len as f32 / max_len as f32;

    // 3. Character trigram Dice coefficient — overall sequence overlap
    let trigram_dice = if inp.len() >= 3 && cand.len() >= 3 {
        let inp_tri: HashSet<&[u8]> = inp.windows(3).collect();
        let cand_tri: HashSet<&[u8]> = cand.windows(3).collect();
        let common = inp_tri.intersection(&cand_tri).count();
        let total = inp_tri.len() + cand_tri.len();
        if total > 0 { 2.0 * common as f32 / total as f32 } else { 0.0 }
    } else {
        0.0
    };

    // 4. Length similarity — penalize very different lengths
    let len_ratio = inp.len().min(cand.len()) as f32 / max_len as f32;

    // Combined: prefix 35%, suffix 25%, trigrams 25%, length 15%
    prefix_ratio * 0.35 + suffix_ratio * 0.25 + trigram_dice * 0.25 + len_ratio * 0.15
}

fn main() {
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let mfst_path = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let dict_path = base.join("../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let wf_path = base.join("../contexter-repo/training-data/wordfreq.tsv");
    let training = base.join("../contexter-repo/training-data");

    println!("Loading FST...");
    let fst = load_fst_from_mfst(mfst_path.to_str().unwrap())
        .expect("Failed to load FST");

    let analyzer = mtag::Analyzer::new(dict_path.to_str().unwrap())
        .expect("Failed to load analyzer");

    let wordfreq = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);

    let onnx_path = training.join("onnx/norbert4_base_int8.onnx");
    let tok_path = training.join("onnx/tokenizer.json");
    println!("Loading NorBERT4...");
    let mut model = Model::load(onnx_path.to_str().unwrap(), tok_path.to_str().unwrap())
        .expect("Failed to load model");
    println!("Loaded. Vocab: {}\n", model.vocab_size());

    let word_check = |w: &str| -> bool {
        if let Some(readings) = analyzer.dict_lookup(w) {
            readings.iter().any(|r| r.pos != mtag::types::Pos::Prop)
        } else {
            false
        }
    };
    // Last compound part must be PRIMARILY a noun.
    // "finsk" has 4 adj + 2 subst readings → primarily adjective → reject
    // "fisk" has 1 subst + 1 verb → primarily noun → accept
    let noun_check = |w: &str| -> bool {
        if let Some(readings) = analyzer.dict_lookup(w) {
            let noun_count = readings.iter().filter(|r| r.pos == mtag::types::Pos::Subst).count();
            let adj_count = readings.iter().filter(|r| r.pos == mtag::types::Pos::Adj).count();
            noun_count > 0 && noun_count >= adj_count
        } else {
            false
        }
    };

    // (sentence_with_misspelling, misspelled_word, expected_correct, description)
    let tests: Vec<(&str, &str, Vec<&str>, &str)> = vec![
        // === Baseline: single words + known compounds ===
        ("Vi renoverte sjøkken i fjor.", "sjøkken", vec!["kjøkken"], "single: s→k"),
        ("Maten står på kjøkkenbort.", "kjøkkenbort", vec!["kjøkkenbord"], "known compound: t→d"),
        ("Vi kjøpte nytt sjøkkenbord.", "sjøkkenbord", vec!["kjøkkenbord"], "known compound: sj→kj"),
        ("Maten står på sjøkkenbort.", "sjøkkenbort", vec!["kjøkkenbord"], "known compound: both parts"),

        // === Productive compounds — error in part 1 ===
        ("Vi fanget innsjefisk i går.", "innsjefisk", vec!["innsjøfisk"], "productive: e→ø in innsjø"),
        ("Jeg kjøpte kyllingsfilet til middag.", "kyllingsfilet", vec!["kyllingfilet"], "productive: extra s"),
        ("Hun spiste en lakzebit til lunsj.", "lakzebit", vec!["laksebit"], "productive: z→s"),
        ("Bestemor laget jordbergrøt.", "jordbergrøt", vec!["jordbærgrøt"], "productive: e→æ in bær"),
        ("Vi spiste lunsj i skollekantine.", "skollekantine", vec!["skolekantine"], "productive: ll→l"),
        ("Det er eksamennsperiode på skolen.", "eksamennsperiode", vec!["eksamensperiode"], "productive: nn→n"),
        ("Han leverte prosjektrapport i dag.", "prosjektrapport", vec!["prosjektrapport"], "productive: exact"),
        ("Vi betalte fakturaggebyr.", "fakturaggebyr", vec!["fakturagebyr"], "productive: gg→g"),
        ("Åpne en netbuttikk på nett.", "netbuttikk", vec!["nettbutikk"], "productive: t→tt, tt→t"),
        ("Eleven trenger lekssehjlep.", "lekssehjlep", vec!["leksehjelp"], "productive: hj swap"),

        // === Productive compounds — error in part 2 ===
        ("Vi fanget innsjøfissk i vannet.", "innsjøfissk", vec!["innsjøfisk"], "productive: ss→s in fisk"),
        ("Hun stekte kyllingfilét til middag.", "kyllingfilét", vec!["kyllingfilet"], "productive: é→e"),
        ("Legg osteskivve på brødet.", "osteskivve", vec!["osteskive"], "productive: vv→v"),
        ("Barna var på svømmestevnne.", "svømmestevnne", vec!["svømmestevne"], "productive: nn→n"),
        ("Laget vant kampressultat.", "kampressultat", vec!["kampresultat"], "productive: ss→s"),
        ("Husk å endre passordbytte.", "passordbytte", vec!["passordbytte"], "productive: exact"),
        ("Vi kjøpte ny ladekabbel.", "ladekabbel", vec!["ladekabel"], "productive: bb→b"),
        ("Det kom et nyhettsvarsel.", "nyhettsvarsel", vec!["nyhetsvarsel"], "productive: tt→t"),
        ("Han mistet busskort.", "busskort", vec!["busskort"], "productive: exact compound"),

        // === Productive compounds — errors in BOTH parts ===
        ("Vi fanget innsjefissk i elva.", "innsjefissk", vec!["innsjøfisk"], "both: e→ø + ss→s"),
        ("Hun kjøpte kyllingsfilét.", "kyllingsfilét", vec!["kyllingfilet", "kyllingsfilet"], "both: extra s + é→e"),
        ("Den gamle taklysekronne var pen.", "taklysekronne", vec!["taklysekrone"], "both: nn→n"),

        // === Phonetic errors ===
        ("Elven nådde frostgrennse.", "frostgrennse", vec!["frostgrense"], "phonetic: nn→n"),
        ("De så nordlyskvell på himmelen.", "nordlyskvell", vec!["nordlyskveld"], "phonetic: l→ld"),
        ("Det var et brått temperaturfel.", "temperaturfel", vec!["temperaturfall"], "phonetic: e→a in fall"),
        ("Vi hadde sommeravsluttning på skolen.", "sommeravsluttning", vec!["sommeravslutning"], "phonetic: tt→t"),
        ("Det ble en fin solvskinnshellg.", "solvskinnshellg", vec!["solskinnshelg"], "phonetic: v→ø, ll→l"),

        // === Binding letter errors ===
        ("Sjekk møteinknalling i kalenderen.", "møteinknalling", vec!["møteinnkalling"], "binding: n→nn"),
        ("Bruk en rengjøringklut.", "rengjøringklut", vec!["rengjøringsklut"], "binding: missing s"),
        ("De trener i treningsstuddo.", "treningsstuddo", vec!["treningsstudio"], "binding: dd→d, o→io"),

        // === Modern/tech compounds ===
        ("Barna har for mye skjermtitt.", "skjermtitt", vec!["skjermtid"], "tech: tt→d"),
        ("Det var et datakrasj på serveren.", "datakrasj", vec!["datakrasj"], "tech: exact"),
        ("Vi hadde videomette i dag.", "videomette", vec!["videomøte"], "tech: e→ø"),
        ("Bruk en strømetjeneste for musikk.", "strømetjeneste", vec!["strømmetjeneste"], "tech: missing m"),

        // === Three-part compounds ===
        ("Sjekk vaskemaskinslannge bak.", "vaskemaskinslannge", vec!["vaskemaskinslange"], "3-part: nn→n"),
        ("Vi bodde på flyplashotell.", "flyplashotell", vec!["flyplasshotell"], "3-part: missing s"),

        // === Double consonant confusion ===
        ("Hun trener på jogamatte.", "jogamatte", vec!["yogamatte"], "double: j→y"),
        ("Barnet laget leerfigur.", "leerfigur", vec!["leirfigur"], "double: ee→ei"),
        ("Hun tok en allergittes.", "allergittes", vec!["allergitest"], "double: tt→t, extra s"),
        ("Han har dårlig sevnkvalitet.", "sevnkvalitet", vec!["søvnkvalitet"], "phonetic: e→ø"),
    ];

    let mut pass = 0;
    let mut fail = 0;
    let mut total_walker_ms = 0.0_f64;
    let mut total_bert_ms = 0.0_f64;

    for (sentence, misspelled, expected, desc) in &tests {
        // Step 1: Compound walker — get candidates
        let t_walk = Instant::now();
        let results = compound_fuzzy_walk(
            &fst, &misspelled.to_lowercase(),
            Some(&wordfreq), Some(&word_check), Some(&noun_check),
        );
        let walk_ms = t_walk.elapsed().as_secs_f64() * 1000.0;
        total_walker_ms += walk_ms;

        // Combine two selection strategies to get candidates for BERT:
        // 1. Walker's top 20 (by edit distance — catches BERT-friendly candidates)
        // 2. Ortho-reranked top 50 → top 20 (by prefix/trigram — catches similar-looking ones)
        // Union + dedup gives BERT a diverse set to score.
        let input_lower = misspelled.to_lowercase();
        let mut seen = HashSet::new();
        // Strategy 1: walker's top 20
        let walker_top: Vec<&str> = results.iter()
            .take(20)
            .map(|r| r.compound_word.as_str())
            .filter(|w| seen.insert(*w))
            .collect();
        // Strategy 2: ortho-reranked from top 50
        let mut top50: Vec<(&str, f32)> = results.iter()
            .take(50)
            .map(|r| (r.compound_word.as_str(), ortho_score(&input_lower, &r.compound_word)))
            .collect();
        top50.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
        let ortho_top: Vec<&str> = top50.iter()
            .filter(|(w, _)| seen.insert(*w))
            .take(10)
            .map(|(w, _)| *w)
            .collect();
        let mut candidates = walker_top;
        candidates.extend(ortho_top);

        if candidates.is_empty() {
            println!("  FAIL ({:>5.1}ms + 0ms): {} — no candidates", walk_ms, desc);
            fail += 1;
            continue;
        }

        // Step 2: BERT re-ranking + orthographic similarity
        let bert_sentence = sentence.to_lowercase().replace(
            &input_lower, "PLACEHOLDER"
        );

        // Get edit distance for each candidate from walker results
        let edit_map: std::collections::HashMap<&str, u32> = results.iter()
            .map(|r| (r.compound_word.as_str(), r.total_edits))
            .collect();

        let t_bert = Instant::now();
        let mut scored: Vec<(&str, f32, f32, f32)> = candidates.iter()
            .map(|&cand| {
                let s = bert_sentence.replace("PLACEHOLDER", cand);
                let bert = subword_score(&mut model, &s, cand);
                let ortho = ortho_score(&input_lower, cand);
                // Compound score: BERT picks the right word in context,
                // ortho similarity is a bonus (prefix + trigram match).
                // No edit distance penalty — the input IS wrong, corrections
                // with edits should compete equally with exact decompositions.
                let bert_norm = (bert / 25.0).clamp(0.0, 1.0);
                let combined = bert_norm * 7.0 + ortho * 3.0;
                (cand, combined, bert, ortho)
            })
            .collect();
        let bert_ms = t_bert.elapsed().as_secs_f64() * 1000.0;
        total_bert_ms += bert_ms;

        scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

        // Post-BERT: if top result and another candidate are the SAME WORD
        // (same root, different inflection), prefer the one whose suffix
        // matches the input. E.g., input "gebyr" → pick "gebyr" over "gebyret".
        if scored.len() >= 2 {
            let top = scored[0].0;
            let top_bytes = top.as_bytes();
            let inp_bytes = input_lower.as_bytes();

            for i in 1..scored.len().min(10) {
                let alt = scored[i].0;
                let alt_bytes = alt.as_bytes();

                // Same word = share 80%+ of shorter word as common prefix
                let common = top_bytes.iter().zip(alt_bytes.iter())
                    .take_while(|(a, b)| a == b).count();
                // Same word = one is a prefix of the other, or they differ
                // only in the very last byte (e.g., osteskive/osteskiva)
                let shorter = top_bytes.len().min(alt_bytes.len());
                if shorter == 0 || common + 1 < shorter {
                    continue; // different words
                }

                // Same root — compare suffix match to input
                let top_suffix: usize = top_bytes.iter().rev().zip(inp_bytes.iter().rev())
                    .take_while(|(a, b)| a == b).count();
                let alt_suffix: usize = alt_bytes.iter().rev().zip(inp_bytes.iter().rev())
                    .take_while(|(a, b)| a == b).count();

                if alt_suffix > top_suffix {
                    scored.swap(0, i);
                    break;
                }
            }
        }

        let bert_top1 = scored[0].0;
        let found = expected.iter().any(|exp| *exp == bert_top1);

        let top3: Vec<String> = scored.iter().take(3)
            .map(|(w, combined, bert, ortho)| format!("{}({:.1}b={:.1}o={:.2})", w, combined, bert, ortho))
            .collect();

        if found {
            println!("  PASS ({:>5.1}ms+{:>5.1}ms): {} → {} | [{}]",
                walk_ms, bert_ms, desc, bert_top1, top3.join(", "));
            pass += 1;
        } else {
            // Check if expected is anywhere in BERT-ranked list
            let bert_rank = expected.iter().find_map(|exp| {
                scored.iter().position(|(w, _, _, _)| w == exp).map(|p| p + 1)
            });
            let rank_info = bert_rank.map_or("not in top 20".to_string(), |r| format!("#{}", r));
            println!("  FAIL ({:>5.1}ms+{:>5.1}ms): {} — bert#1='{}' expected {:?} ({}) | [{}]",
                walk_ms, bert_ms, desc, bert_top1, expected, rank_info, top3.join(", "));
            fail += 1;
        }
    }

    println!("\nResults: {}/{} passed", pass, pass + fail);
    println!("Avg walker: {:.1}ms, avg BERT: {:.1}ms, avg total: {:.1}ms",
        total_walker_ms / tests.len() as f64,
        total_bert_ms / tests.len() as f64,
        (total_walker_ms + total_bert_ms) / tests.len() as f64,
    );
    if fail > 0 { std::process::exit(1); }
}
