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

    // Debug: check "innsjøefisk" decomposition in walker
    {
        let r = compound_fuzzy_walk(&fst, "innsjefisk", Some(&wordfreq), Some(&word_check), Some(&noun_check));
        println!("  === innsjefisk: results containing 'efisk' ===");
        for x in r.iter().filter(|x| x.compound_word.contains("efisk")).take(5) {
            let parts: Vec<String> = x.parts.iter()
                .map(|p| format!("'{}' (e={})", p.matched_word, p.edits)).collect();
            println!("    {} [{}] total_edits={}", x.compound_word, parts.join(" + "), x.total_edits);
        }
        if r.iter().all(|x| !x.compound_word.contains("efisk")) {
            println!("    (none found)");
        }
    }

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

        // Pre-BERT: walker's top 20 candidates (by edit distance + freq ranking)
        let input_lower = misspelled.to_lowercase();
        let mut seen = HashSet::new();
        let candidates: Vec<&str> = results.iter()
            .take(20)
            .map(|r| r.compound_word.as_str())
            .filter(|w| seen.insert(*w))
            .collect();

        if candidates.is_empty() {
            println!("  FAIL ({:>5.1}ms + 0ms): {} — no candidates", walk_ms, desc);
            fail += 1;
            continue;
        }

        // Step 2: BERT re-ranking using production approach
        // One base forward pass with <mask> at word position, then score
        // each candidate's tokens. BERT sees only surrounding context,
        // NOT the compound's internal structure.
        // Keep original casing — BERT is case-sensitive!
        // "Vi renoverte" gives much better scores than "vi renoverte"
        let sent_for_bert = sentence.to_string();
        // Find the misspelled word position case-insensitively
        let lower_sent = sentence.to_lowercase();
        let word_pos = lower_sent.find(&input_lower).unwrap_or(0);
        let before = &sentence[..word_pos];
        let after = &sentence[word_pos + misspelled.len()..];
        let ctx_parts: Vec<&str> = vec![before, after];
        let ctx_before = ctx_parts[0].trim_end();
        let ctx_after = ctx_parts[1].trim_start();

        // Tokenize each candidate
        let cand_tokens: Vec<(&str, Vec<u32>)> = candidates.iter()
            .filter_map(|&c| {
                let enc = model.tokenizer.encode(format!(" {}", c), false).ok()?;
                let ids: Vec<u32> = enc.get_ids().to_vec();
                if ids.is_empty() { None } else { Some((c, ids)) }
            })
            .collect();

        let t_bert = Instant::now();

        // Base forward pass: "<ctx_before> <mask> <ctx_after>"
        let masked_sent = format!("{} <mask> {}", ctx_before, ctx_after);
        let base_logits = model.single_forward(&masked_sent)
            .map(|(logits, _)| logits)
            .unwrap_or_default();

        // First-token score from base logits (free — just a lookup)
        let mut scores: Vec<f32> = cand_tokens.iter()
            .map(|(_, ids)| {
                if base_logits.is_empty() { 0.0 }
                else { base_logits[ids[0] as usize] }
            })
            .collect();

        // Multi-token scoring: incremental prefix + mask
        let max_tokens = cand_tokens.iter().map(|(_, ids)| ids.len()).max().unwrap_or(1);
        for t in 1..max_tokens {
            let to_score: Vec<usize> = cand_tokens.iter().enumerate()
                .filter(|(_, (_, ids))| ids.len() > t)
                .map(|(i, _)| i)
                .collect();
            if to_score.is_empty() { break; }

            // Batch all candidates that need token t scored
            let batch_texts: Vec<String> = to_score.iter()
                .map(|&i| {
                    let partial = model.tokenizer.decode(&cand_tokens[i].1[..t], false)
                        .unwrap_or_default();
                    format!("{} {}<mask> {}", ctx_before, partial.trim(), ctx_after)
                })
                .collect();

            if let Ok((batch_logits, _)) = model.batched_forward(&batch_texts) {
                for (k, &i) in to_score.iter().enumerate() {
                    if k < batch_logits.len() {
                        scores[i] += batch_logits[k][cand_tokens[i].1[t] as usize];
                    }
                }
            }
        }

        // Score = average logit per token, penalized for extra tokens.
        // Multi-token words get inflated scores because BERT easily
        // predicts continuations. Penalty: -1.5 per extra token.
        // A 1-token word (in BERT's vocab) is one BERT knows well.
        let mut scored: Vec<(&str, f32, f32, f32)> = cand_tokens.iter().enumerate()
            .map(|(i, (w, ids))| {
                let avg = scores[i] / ids.len() as f32;
                let penalty = (ids.len() as f32 - 1.0) * 1.5;
                let bert = avg - penalty;
                let ortho = ortho_score(&input_lower, w);
                let bert_norm = (bert / 25.0).clamp(0.0, 1.0);
                let combined = bert_norm * 7.0 + ortho * 3.0;
                (*w, combined, bert, ortho)
            })
            .collect();
        let bert_ms = t_bert.elapsed().as_secs_f64() * 1000.0;
        total_bert_ms += bert_ms;

        scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

        // Anti-typo: if top result is near-exact match of the misspelling
        // (ortho ≥ 0.95) but BERT scores another candidate 2+ points higher,
        // the exact match is probably the typo — trust BERT.
        // E.g., "kyllingsfilet"(b=10.4,o=1.0) → swap to "kyllingfilet"(b=14.5)
        if scored.len() >= 2 && scored[0].3 >= 0.95 {
            for i in 1..scored.len().min(5) {
                if scored[i].2 > scored[0].2 + 2.0 {
                    scored.swap(0, i);
                    break;
                }
            }
        }

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
