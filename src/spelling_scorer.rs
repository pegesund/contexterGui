//! Shared spelling suggestion pipeline.
//!
//! There is **one** candidate-generation function: `find_candidates_pipeline`.
//! Both `main.rs::find_spelling_suggestions` (the GUI path) and
//! `test_spelling.rs` (the regression test) MUST call this function. Test
//! coverage that bypasses it is a lie — if the GUI runs different code than
//! the test, "X/Y passed" tells you nothing about user-visible behaviour.
//! See `feedback_spelling_pipeline_duplicated.md` for the incident this
//! arrangement was created to prevent.
//!
//! Language dispatch for fuzzy matching lives inside this function:
//!   - `lang.uses_compound_lookup() == true` (Bokmål, Nynorsk): the compound
//!     walker is the ONLY source of fuzzy-distance candidates. It handles
//!     single-word matches AND multi-part decompositions with Damerau
//!     transposition and UTF-8 vowel swaps. We do NOT also call
//!     `analyzer.fuzzy_lookup` — that loses transposition and dyslexic
//!     vowel-swap handling.
//!   - `lang.uses_compound_lookup() == false` (English, future others):
//!     `analyzer.fuzzy_lookup` is the only fuzzy source. We do NOT call
//!     the compound walker — those languages don't form productive
//!     compounds and the walker would just generate junk.
//!
//! Phase 1: `find_candidates_pipeline` — candidate generation + ortho scoring + dict filter
//! Phase 2: `score_and_rerank` — BERT boundary scoring + grammar correction + hybrid sentence re-ranking

use std::collections::{HashMap, HashSet};
use nostos_cognio::grammar::types::GrammarError;
use nostos_cognio::model::Model;
use nostos_cognio::spelling;

/// Levenshtein edit distance between two strings.
pub fn levenshtein_distance(a: &str, b: &str) -> u32 {
    let (a, b): (Vec<char>, Vec<char>) = (a.chars().collect(), b.chars().collect());
    let (m, n) = (a.len(), b.len());
    let mut d = vec![vec![0u32; n + 1]; m + 1];
    for i in 0..=m { d[i][0] = i as u32; }
    for j in 0..=n { d[0][j] = j as u32; }
    for i in 1..=m {
        for j in 1..=n {
            let cost = if a[i - 1] == b[j - 1] { 0 } else { 1 };
            d[i][j] = (d[i - 1][j] + 1).min(d[i][j - 1] + 1).min(d[i - 1][j - 1] + cost);
        }
    }
    d[m][n]
}

fn common_prefix_len(a: &str, b: &str) -> usize {
    a.chars()
        .zip(b.chars())
        .take_while(|(left, right)| left == right)
        .count()
}

fn promote_orthographic_anchor(
    word_lower: &str,
    ranked: &mut Vec<(String, f32)>,
    ortho_map: &HashMap<String, f32>,
) {
    if ranked.len() < 2 || word_lower.chars().count() < 8 {
        return;
    }

    let best_lower = ranked[0].0.to_lowercase();
    let best_prefix = common_prefix_len(word_lower, &best_lower);
    let best_dist = levenshtein_distance(word_lower, &best_lower);
    let best_ortho = ortho_map
        .get(ranked[0].0.as_str())
        .copied()
        .unwrap_or(0.0);

    let anchor = ranked
        .iter()
        .enumerate()
        .take(10)
        .filter_map(|(idx, (candidate, _))| {
            let candidate_lower = candidate.to_lowercase();
            let prefix = common_prefix_len(word_lower, &candidate_lower);
            if prefix < 4 {
                return None;
            }
            let ortho = ortho_map.get(candidate.as_str()).copied().unwrap_or(0.0);
            let dist = levenshtein_distance(word_lower, &candidate_lower);
            Some((idx, prefix, ortho, dist))
        })
        .max_by(|left, right| {
            left.2
                .partial_cmp(&right.2)
                .unwrap_or(std::cmp::Ordering::Equal)
                .then_with(|| left.1.cmp(&right.1))
                .then_with(|| right.3.cmp(&left.3))
        });

    let Some((anchor_idx, _prefix, anchor_ortho, anchor_dist)) = anchor else {
        return;
    };
    if anchor_idx == 0 {
        return;
    }

    let weak_best_shape = best_prefix < 4
        && anchor_ortho >= 0.55
        && anchor_ortho >= best_ortho * 1.30;
    let much_closer_shape = anchor_ortho >= 0.82
        && anchor_dist.saturating_add(1) <= best_dist
        && anchor_ortho >= best_ortho * 1.15;

    if !weak_best_shape && !much_closer_shape {
        return;
    }

    let anchor = ranked.remove(anchor_idx);
    ranked.insert(0, anchor);
}

#[cfg(test)]
mod orthographic_anchor_tests {
    use super::promote_orthographic_anchor;
    use std::collections::HashMap;

    #[test]
    fn promotes_close_norwegian_form_over_contextual_plural() {
        let mut ranked = vec![
            ("dokumenta".to_string(), 12.0),
            ("dokumentet".to_string(), 11.0),
        ];
        let ortho_map = HashMap::from([
            ("dokumenta".to_string(), 0.72),
            ("dokumentet".to_string(), 0.96),
        ]);

        promote_orthographic_anchor("dokummentet", &mut ranked, &ortho_map);

        assert_eq!(ranked[0].0, "dokumentet");
    }

    #[test]
    fn promotes_same_stem_over_unrelated_semantic_candidate() {
        let mut ranked = vec![
            ("attraksjonen".to_string(), 12.0),
            ("applikasjonen".to_string(), 11.0),
        ];
        let ortho_map = HashMap::from([
            ("attraksjonen".to_string(), 0.31),
            ("applikasjonen".to_string(), 0.87),
        ]);

        promote_orthographic_anchor("appllicationen", &mut ranked, &ortho_map);

        assert_eq!(ranked[0].0, "applikasjonen");
    }

    #[test]
    fn keeps_bert_pick_when_spelling_gap_is_small() {
        let mut ranked = vec![
            ("message".to_string(), 12.0),
            ("massage".to_string(), 11.0),
        ];
        let ortho_map = HashMap::from([
            ("message".to_string(), 0.86),
            ("massage".to_string(), 0.90),
        ]);

        promote_orthographic_anchor("mesage", &mut ranked, &ortho_map);

        assert_eq!(ranked[0].0, "message");
    }
}

/// Compute trigrams for a word.
pub fn trigrams(word: &str) -> Vec<String> {
    let chars: Vec<char> = word.chars().collect();
    if chars.len() < 3 {
        return vec![word.to_string()];
    }
    (0..chars.len() - 2)
        .map(|i| chars[i..i + 3].iter().collect())
        .collect()
}

/// Try splitting a word into function_word + remainder.
/// Returns None if the word is already a valid dictionary word (no split needed).
pub fn try_split_function_word(word: &str, analyzer: &mtag::Analyzer, lang: &dyn language::LanguageSpelling) -> Option<String> {
    // Don't split valid compound words (avstand, tilstand, imorgen, etc.)
    if analyzer.has_word(&word.to_lowercase()) {
        return None;
    }
    let function_words = lang.function_words();
    let lower = word.to_lowercase();
    for prefix in function_words {
        if lower.len() <= prefix.len() + 1 { continue; }
        if !lower.starts_with(prefix) { continue; }
        let remainder = &lower[prefix.len()..];
        if remainder.len() < 2 { continue; }
        if analyzer.has_word(remainder) {
            return Some(format!("{} {}", prefix, remainder));
        }
    }
    let chars: Vec<char> = lower.chars().collect();
    let mut best_split: Option<(String, usize)> = None;
    for split_at in 3..=(chars.len().saturating_sub(3)) {
        let left: String = chars[..split_at].iter().collect();
        let right: String = chars[split_at..].iter().collect();
        if analyzer.has_word(&left) && analyzer.has_word(&right) {
            let balance = left.len().min(right.len());
            if best_split.as_ref().map(|(_, b)| balance > *b).unwrap_or(true) {
                best_split = Some((format!("{} {}", left, right), balance));
            }
        }
    }
    best_split.map(|(s, _)| s)
}

/// Compute boost multiplier for document-frequency and user-dictionary.
///
/// Previously this returned a flat 1.0 for any candidate not in the
/// document/user dictionary, so a rare-but-valid dictionary entry like
/// "blakr" tied with a far more common word like "blåbær" and the
/// BERT contextual score was the only tiebreaker. For sentences where
/// the rare word also looks contextually plausible (BERT 14.8 vs 14.2)
/// the wrong candidate would win.
///
/// We now apply a smooth log10-based wordfreq boost (capped at +25 % for
/// the most common words) on top of the existing in-doc / in-user
/// multipliers. Candidates absent from wordfreq still score 1.0, so
/// fabricated junk that happens to be in the dictionary gets no boost.
pub fn compute_boost(
    word: &str,
    doc_word_counts: &HashMap<String, u16>,
    user_dict_words: &[String],
    wordfreq: Option<&HashMap<String, u64>>,
) -> f32 {
    let lower = word.to_lowercase();
    let freq = wordfreq.and_then(|wf| wf.get(&lower)).copied().unwrap_or(0);
    let wf_boost = 1.0 + ((freq as f32 + 1.0).log10() * 0.10).min(0.25);

    let in_doc = doc_word_counts.get(&lower).copied().unwrap_or(0) >= 2;
    let in_user = user_dict_words.iter().any(|uw| uw.eq_ignore_ascii_case(&lower));
    let ctx_mult = match (in_doc, in_user) {
        (true, true)   => 1.6,
        (false, true)  => 1.3,
        (true, false)  => 1.25,
        (false, false) => 1.0,
    };

    wf_boost * ctx_mult
}

/// Phase 1: Generate spelling candidates, ortho-score them, and dictionary-filter.
///
/// THIS IS THE ONLY CANDIDATE-GENERATION ENTRY POINT. Both the GUI path
/// (`main.rs::find_spelling_suggestions`) and the regression test
/// (`test_spelling.rs`) call this function. Do not create a parallel
/// pipeline in main.rs or anywhere else — see the file-level docstring.
///
/// Language dispatch for fuzzy matching:
///   - `lang.uses_compound_lookup()` → compound walker via `compound_fst`
///     (Bokmål, Nynorsk).
///   - Otherwise → `analyzer.fuzzy_lookup` (English, others).
///
/// `compound_fst` must be `Some` whenever `uses_compound_lookup()` is true
/// or no fuzzy candidates will be produced. Conversely, `compound_fst` is
/// ignored when `uses_compound_lookup()` is false.
pub fn find_candidates_pipeline(
    analyzer: &mtag::Analyzer,
    compound_fst: Option<&fst::raw::Fst<Vec<u8>>>,
    wordfreq: Option<&HashMap<String, u64>>,
    user_dict_words: &[String],
    doc_word_counts: &HashMap<String, u16>,
    word: &str,
    _sentence_ctx: &str,
    lang: &dyn language::LanguageSpelling,
) -> Vec<(String, f32)> {
    // Sabotage probe — set `SPELLING_PIPELINE_SABOTAGE=1` and run anything
    // that should produce spelling suggestions. If suggestions still appear
    // somewhere, that caller is bypassing this function and is a duplicated
    // pipeline that must be deleted. Verified 2026-06-26: test_spelling
    // returns 0 candidates under the probe (so test_spelling reaches here),
    // and the GUI also reaches here via main.rs::find_spelling_suggestions.
    if std::env::var("SPELLING_PIPELINE_SABOTAGE").is_ok() {
        return Vec::new();
    }
    let word_lower = word.to_lowercase();
    let word_trigrams = trigrams(&word_lower);
    let word_first = word_lower.chars().next().unwrap_or(' ');

    let mut candidates: Vec<String> = Vec::new();
    let mut seen = HashSet::new();
    let mut edit_distances: HashMap<String, u32> = HashMap::new();
    // Single-word matches discovered by the compound walker (1-part results).
    // Kept separate from the multi-part `compound_candidates` set so callers
    // that distinguish "this is a real compound" from "this is a fuzzy
    // single-word hit" still can.
    let mut compound_candidates: HashSet<String> = HashSet::new();

    let uses_compound = lang.uses_compound_lookup();

    // ── Fuzzy lookup — language-dispatched ──────────────────────────────
    // BM/NN: compound walker handles BOTH single-word fuzzy AND multi-part
    // decompositions, with Damerau transposition + free vowel swaps for
    // dyslexic substitutions. Plain `analyzer.fuzzy_lookup` does NOT have
    // those, so we intentionally skip it for these languages. Mixing the
    // two created the duplicated-pipeline bug logged in memory.
    //
    // Other languages: plain `analyzer.fuzzy_lookup` is the source. The
    // compound walker is skipped — English (and most languages) don't form
    // productive compounds the way Norwegian does, and running the walker
    // would generate junk multi-part suggestions.
    if uses_compound {
        if let Some(fst) = compound_fst {
            // Allow the walker to fire on any length so it can replace the
            // plain fuzzy_lookup for short words too. The walker's own
            // edit-budget (MAX_EDITS_PER_PART = 2, MAX_TOTAL_EDITS = 4)
            // keeps the candidate count bounded.
            let word_check = |w: &str| -> bool {
                analyzer.dict_lookup(w).map_or(false, |rs|
                    rs.iter().any(|r| r.pos != mtag::types::Pos::Prop))
            };
            let results = crate::compound_walker::compound_fuzzy_walk(
                fst, &word_lower, lang, wordfreq,
                Some(&word_check), None,
            );
            // Do NOT cap with `take(N)` here. The walker returns its hits
            // sorted by edit count ascending, so a small cap (was take(30))
            // drops every edits=2 match the moment there are 30+ edits=1
            // single-vowel siblings of the misspelling. For consonant-only
            // input like "lgn" that's exactly what happens: 30+ dist-1
            // variants (lån, len, lon, lin, lyn …) crowd out "lege" at
            // dist=2 and "lege" never reaches the ortho phase. Ortho +
            // dict-filter further down already cap the pool the BERT
            // scorer sees, so adding all walker hits here is free.
            for r in &results {
                let cw = r.compound_word.to_lowercase();
                if cw == word_lower || cw.len() < 2 { continue; }
                if seen.insert(cw.clone()) {
                    edit_distances.insert(cw.clone(), r.total_edits.max(1));
                    compound_candidates.insert(cw.clone());
                    candidates.push(cw);
                }
            }
        }
    } else {
        // Source 1: Fuzzy Levenshtein (distance 2)
        for (w, dist) in analyzer.fuzzy_lookup(&word_lower, 2) {
            let wl = w.to_lowercase();
            if wl == word_lower || wl.len() < 2 { continue; }
            edit_distances.insert(wl.clone(), dist);
            if seen.insert(wl.clone()) { candidates.push(wl); }
        }

        // Source 4: Truncated fuzzy (strip 1-2 trailing chars then fuzzy)
        for strip in 1..=2u32 {
            let chars: Vec<char> = word_lower.chars().collect();
            if chars.len() <= 3 + strip as usize { continue; }
            let truncated: String = chars[..chars.len() - strip as usize].iter().collect();
            for (w, dist) in analyzer.fuzzy_lookup(&truncated, 2) {
                let wl = w.to_lowercase();
                edit_distances.entry(wl.clone()).or_insert(dist + strip);
                if wl != word_lower && wl.len() >= 2 && seen.insert(wl.clone()) {
                    candidates.push(wl);
                }
            }
        }
    }

    // Source 2: Prefix lookup (missing-letter typos)
    // Applies to all languages — it's a targeted source, not a fuzzy lookup.
    for w in analyzer.prefix_lookup(&word_lower, 20) {
        let wl = w.to_lowercase();
        let extra = wl.len() as i32 - word_lower.len() as i32;
        if extra >= 1 && extra <= 3 {
            edit_distances.entry(wl.clone()).or_insert(extra as u32);
            if wl != word_lower && wl.len() >= 2 && seen.insert(wl.clone()) {
                candidates.push(wl);
            }
        }
    }

    // Source 3: Prefix with last char removed
    let char_count = word_lower.chars().count();
    if char_count >= 3 {
        let end_byte = word_lower.char_indices().rev().next().map(|(i, _)| i).unwrap_or(0);
        let shorter = &word_lower[..end_byte];
        for w in analyzer.prefix_lookup(shorter, 20) {
            let wl = w.to_lowercase();
            let diff = (wl.len() as i32 - word_lower.len() as i32).unsigned_abs() + 1;
            edit_distances.entry(wl.clone()).or_insert(diff);
            if wl != word_lower && wl.len() >= 2 && seen.insert(wl.clone()) {
                candidates.push(wl);
            }
        }
    }

    // Source 3b: Long-word prefix rescue.
    //
    // BM/NN normally get long fuzzy candidates from the compound walker, but
    // one-part corrections still obey MAX_EDITS_PER_PART=2 there. Loanword
    // typos such as "appllicationen" -> "applikasjonen" are edit-distance 4
    // (extra l, c->k, t->s, i->j). Bokmal happened to recover these via
    // wordfreq, while Nynorsk can miss them when the valid form is in the FST
    // but absent from wordfreq. Reuse the existing prefix source with a short
    // stable prefix, then keep only close dictionary words.
    if char_count >= 8 {
        let prefix_len = 4.min(char_count);
        let prefix_end = word_lower
            .char_indices()
            .nth(prefix_len)
            .map(|(idx, _)| idx)
            .unwrap_or(word_lower.len());
        let broad_prefix = &word_lower[..prefix_end];
        if broad_prefix.len() >= 4 {
            for w in analyzer.prefix_lookup(broad_prefix, 80) {
                let wl = w.to_lowercase();
                if wl == word_lower || wl.len() < 5 || seen.contains(&wl) {
                    continue;
                }
                let len_delta = (wl.chars().count() as i32 - char_count as i32).unsigned_abs();
                if len_delta > 3 {
                    continue;
                }
                let dist = levenshtein_distance(&word_lower, &wl);
                if dist > 4 {
                    continue;
                }
                let w_tri = trigrams(&wl);
                let common = word_trigrams.iter().filter(|t| w_tri.contains(t)).count();
                if common >= 2 && seen.insert(wl.clone()) {
                    edit_distances.entry(wl.clone()).or_insert(dist);
                    candidates.push(wl);
                }
            }
        }
    }

    // Source 6: Split function word
    if let Some(split) = try_split_function_word(&word_lower, analyzer, lang) {
        let sl = split.to_lowercase();
        if seen.insert(sl.clone()) { candidates.push(sl); }
    }

    // Source 7: Wordfreq — common words with trigram overlap
    if let Some(wf) = wordfreq {
        for (w, _freq) in wf.iter() {
            let wl = w.to_lowercase();
            if wl == word_lower || seen.contains(&wl) { continue; }
            if wl.chars().next().unwrap_or(' ') != word_first { continue; }
            let w_tri = trigrams(&wl);
            let common = word_trigrams.iter().filter(|t| w_tri.contains(t)).count();
            if common >= 2 && seen.insert(wl.clone()) {
                candidates.push(wl);
            }
        }
    }

    // Source 8: User dictionary — words within edit distance 2
    for uw in user_dict_words {
        let uwl = uw.to_lowercase();
        if uwl == word_lower || seen.contains(&uwl) { continue; }
        let dist = levenshtein_distance(&word_lower, &uwl);
        if dist <= 2 {
            edit_distances.entry(uwl.clone()).or_insert(dist);
            if seen.insert(uwl.clone()) { candidates.push(uwl); }
        }
    }

    // Source 9: Long word truncation (>= 10 chars).
    // The "is this word real" check accepts either a direct dict hit OR a
    // 0-edit compound decomposition (BM/NN). For non-compound languages the
    // compound_fst is None so only the dict hit applies.
    if word_lower.len() >= 10 {
        let is_known = |w: &str| -> bool {
            if analyzer.has_word(w) { return true; }
            if uses_compound {
                if let Some(fst) = compound_fst {
                    let r = crate::compound_walker::compound_fuzzy_walk(
                        fst, w, lang, wordfreq, None, None,
                    );
                    return r.iter().any(|c| c.total_edits == 0);
                }
            }
            false
        };
        for strip in 1..=2usize {
            if word_lower.is_char_boundary(strip) {
                let trimmed = &word_lower[strip..];
                if trimmed.len() >= 5 && is_known(trimmed) && seen.insert(trimmed.to_string()) {
                    edit_distances.insert(trimmed.to_string(), strip as u32);
                    candidates.push(trimmed.to_string());
                }
            }
            let end = word_lower.len() - strip;
            if word_lower.is_char_boundary(end) {
                let trimmed = &word_lower[..end];
                if trimmed.len() >= 5 && is_known(trimmed) && seen.insert(trimmed.to_string()) {
                    edit_distances.insert(trimmed.to_string(), strip as u32);
                    candidates.push(trimmed.to_string());
                }
            }
        }
    }

    // Source 10: First-character swap
    if word_lower.len() >= 3 {
        let rest = &word_lower[word_first.len_utf8()..];
        for c in lang.first_char_alphabet().chars() {
            if c == word_first { continue; }
            let candidate = format!("{}{}", c, rest);
            if analyzer.has_word(&candidate) && seen.insert(candidate.clone()) {
                edit_distances.insert(candidate.clone(), 1);
                candidates.push(candidate);
            }
        }
    }

    // Source 11: Inflected forms of candidates
    {
        use mtag::types::{Pos, Tag};
        let base_candidates: Vec<String> = candidates.clone();
        for base in &base_candidates {
            // Skeleton-match heritage: an inflected form derived from a
            // walker-confirmed lemma should ride the same wf-cap floor as
            // the lemma itself. Without this, "skog" wins ortho 1.000 via
            // its compound_candidates flag while "skogen" (added here from
            // skog's Be inflection) sits at 0.998 and gets crowded out of
            // the dict filter by the 100+ siblings of the misspelling that
            // also hit the cap. For consonant-skeleton typos this is the
            // exact failure mode the test exercises (skgn→skogen,
            // bnkn→benken, lgn→legen).
            let base_is_compound_hit = compound_candidates.contains(base);
            if let Some(readings) = analyzer.dict_lookup(base) {
                for r in &readings {
                    if !matches!(r.pos, Pos::Subst) { continue; }
                    for tag in &[Tag::Be, Tag::Fl] {
                        let forms = analyzer.forms_for_lemma(&r.lemma, &Pos::Subst, tag);
                        for form in forms {
                            let fl = form.to_lowercase();
                            if fl != word_lower && fl.len() >= 2 && seen.insert(fl.clone()) {
                                let dist = levenshtein_distance(&word_lower, &fl);
                                if dist <= 4 {
                                    edit_distances.insert(fl.clone(), dist);
                                    if base_is_compound_hit {
                                        compound_candidates.insert(fl.clone());
                                    }
                                    candidates.push(fl);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Source 12: Phonetic substitutions for dyslexic users
    // Language-specific sound confusions from the LanguageSpelling trait.
    // For BM/NN the compound walker's built-in free_vowel_swaps already
    // covers the dyslexic substitutions (e↔ø, o↔å, a↔æ), so we skip this
    // pass to avoid duplicating that work — and to keep all fuzzy-style
    // candidate generation funneled through the compound walker per the
    // language dispatch rule above.
    if !uses_compound {
        let phonetic_subs = lang.phonetic_substitutions();

        // Single substitution pass
        let mut phonetic_candidates: Vec<String> = Vec::new();
        for &(from, to) in phonetic_subs {
            let mut pos = 0;
            while let Some(idx) = word_lower[pos..].find(from) {
                let abs_idx = pos + idx;
                let result = format!("{}{}{}", &word_lower[..abs_idx], to, &word_lower[abs_idx + from.len()..]);
                if result != word_lower && analyzer.has_word(&result) && seen.insert(result.clone()) {
                    edit_distances.insert(result.clone(), 1);
                    candidates.push(result.clone());
                    phonetic_candidates.push(result);
                }
                pos = abs_idx + from.len();
            }
        }

        // Two-step phonetic chain: apply a second substitution to results of the first
        // Catches "gåtterier" → "gotterier" (å→o) → "godterier" (tt→dt)
        let chain_candidates = phonetic_candidates.clone();
        for base in &chain_candidates {
            for &(from, to) in phonetic_subs {
                let mut pos = 0;
                while let Some(idx) = base[pos..].find(from) {
                    let abs_idx = pos + idx;
                    let result = format!("{}{}{}", &base[..abs_idx], to, &base[abs_idx + from.len()..]);
                    if result != word_lower && result != *base && analyzer.has_word(&result) && seen.insert(result.clone()) {
                        edit_distances.insert(result.clone(), 2);
                        candidates.push(result);
                    }
                    pos = abs_idx + from.len();
                }
            }
        }
    }

    // Phase 2: Ortho score
    //
    // Skeleton heuristic: if the misspelled word is short and contains no
    // vowels, the user almost certainly typed an abbreviation (lgn for
    // legen, skgn for skogen, bnkn for benken). Candidates whose consonant
    // skeleton matches the input letter-for-letter are far more likely
    // intentions than dist-1 single-vowel siblings of the misspelling.
    // Without this lift, "legen" sits at ortho 1.00 while 100+ short
    // siblings (lån, len, lin, lon, lyn, løgn, lign, lun …) all hit the
    // same cap and crowd it out of the top-30 dict filter.
    fn consonant_skeleton(s: &str) -> String {
        s.chars()
            .filter(|c| c.is_alphabetic() && !matches!(c,
                'a' | 'e' | 'i' | 'o' | 'u' | 'y' | 'å' | 'ø' | 'æ'))
            .collect()
    }
    let input_skel = consonant_skeleton(&word_lower);
    let input_is_skeleton = !input_skel.is_empty()
        && word_lower.chars().count() <= 5
        && !word_lower.chars().any(|c|
            matches!(c, 'a' | 'e' | 'i' | 'o' | 'u' | 'y' | 'å' | 'ø' | 'æ'));

    let mut ortho_scored: Vec<(String, f32)> = Vec::new();
    for w in &candidates {
        let w_tri = trigrams(w);
        let common = word_trigrams.iter().filter(|t| w_tri.contains(t)).count();
        let max_t = word_trigrams.len().max(w_tri.len()).max(1);
        let trigram_sim = common as f32 / max_t as f32;

        let prefix_len = word_lower.chars().zip(w.chars())
            .take_while(|(a, b)| a == b).count();
        let max_len = word_lower.chars().count().max(w.chars().count()).max(1);
        let prefix_sim = prefix_len as f32 / max_len as f32;

        let edit_sim = match edit_distances.get(w.as_str()) {
            Some(1) => 0.85,
            Some(2) => 0.65,
            Some(3) => 0.45,
            _ => 0.0,
        };
        let mut ortho_sim = trigram_sim.max(edit_sim).max(prefix_sim);
        if w.chars().next() == Some(word_first) {
            ortho_sim += 0.15;
        }
        let boost = compute_boost(w, doc_word_counts, user_dict_words, wordfreq);
        // Compound-walker single-word matches must not be penalized for being
        // absent from wordfreq. The walker only emits dict-real words within
        // its strict edit budget, so the match itself is a strong signal;
        // floor the boost at the wf-cap so a rare-but-walker-confirmed hit
        // can compete with common siblings of the misspelling.
        let mut effective_boost = if compound_candidates.contains(w) { boost.max(1.25) } else { boost };

        // Skeleton-match lift: see Phase 2 docstring above. Compares the
        // misspelling's consonant skeleton with the candidate's, and lifts
        // ortho by a factor that puts skeleton hits above all the dist-1
        // single-vowel siblings of the input. 1.5 was chosen empirically:
        // 1.25 was not enough to clear the cap, 2.0 dragged junk skeleton
        // matches above legitimate dist-1 fixes for non-skeleton inputs
        // (but the gate above prevents that branch entirely for those).
        if input_is_skeleton && consonant_skeleton(w) == input_skel {
            effective_boost = effective_boost.max(1.5);
        }
        ortho_sim *= effective_boost;
        ortho_scored.push((w.clone(), ortho_sim));
    }
    ortho_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

    // Phase 4: Dictionary filter (top 30)
    //
    // We need a generous pool so that for typos where competitors share a
    // common prefix (e.g. "sykell" matches the entire "syk*" inflection
    // family) the actually-intended word ("sykkel") still survives ortho
    // sorting and reaches BERT. BERT re-ranks the top 30 fast (~24ms) so
    // the extra candidates are cheap.
    let mut passing: Vec<(String, f32)> = Vec::new();
    let mut checked = 0;
    for (candidate, score) in &ortho_scored {
        if checked >= 30 { break; }
        if !word_lower.contains('-') && candidate.contains('-') { continue; }
        let words: Vec<&str> = candidate.split_whitespace().collect();
        if words.iter().any(|w| {
            !analyzer.has_word(w)
            && !user_dict_words.iter().any(|uw| uw.eq_ignore_ascii_case(w))
        }) {
            continue;
        }
        checked += 1;
        passing.push((candidate.clone(), *score));
    }

    passing
}

/// Score a candidate word in sentence context.
/// Masks each word and sums logits, weighting the TARGET word position 3x.
/// `target` is the candidate word to emphasize (the one that differs between candidates).
pub fn sentence_score(model: &mut Model, sentence: &str, target: &str) -> f32 {
    let words: Vec<&str> = sentence.split_whitespace().collect();
    if words.is_empty() { return f32::NEG_INFINITY; }
    let target_lower = target.to_lowercase();
    let mut total: f32 = 0.0;
    let mask_str = model.mask_token_str();
    let mut weight_sum: f32 = 0.0;
    for i in 0..words.len() {
        let word_clean = words[i].trim_matches(|c: char| c.is_ascii_punctuation());
        if word_clean.is_empty() { continue; }
        let masked: String = words.iter().enumerate()
            .map(|(j, w)| if j == i { mask_str.as_str() } else { *w })
            .collect::<Vec<_>>().join(" ");
        if let Ok((logits, _)) = model.single_forward(&masked) {
            if let Ok(enc) = model.tokenizer.encode(format!(" {}", word_clean.to_lowercase()), false) {
                let ids = enc.get_ids();
                if let Some(&first_id) = ids.first() {
                    let w = 1.0_f32; // equal weight for all positions
                    total += logits[first_id as usize] * w;
                    weight_sum += w;
                }
            }
        }
    }
    if weight_sum == 0.0 { return f32::NEG_INFINITY; }
    total / weight_sum
}

/// Score a multi-token candidate by masking each subword token individually.
/// "Katten sitter på stolen" with candidate "sitter" (tokens: "s", "itter"):
///   - Mask "s": "Katten <mask> itter på stolen" → logit for "s"
///   - Mask "itter": "Katten s <mask> på stolen" → logit for "itter"
///   - Sum both → total score for "sitter"
pub fn subword_score(model: &mut Model, sentence: &str, candidate: &str) -> f32 {
    // Tokenize full sentence to get token IDs and find candidate position
    let enc = match model.tokenizer.encode(sentence.to_string(), false) {
        Ok(e) => e,
        Err(_) => return f32::NEG_INFINITY,
    };
    let sent_ids: Vec<u32> = enc.get_ids().to_vec();

    // Tokenize candidate to get its subword tokens
    let cand_enc = match model.tokenizer.encode(format!(" {}", candidate.to_lowercase()), false) {
        Ok(e) => e,
        Err(_) => return f32::NEG_INFINITY,
    };
    let cand_ids: Vec<u32> = cand_enc.get_ids().to_vec();
    if cand_ids.is_empty() { return f32::NEG_INFINITY; }

    // Find where candidate tokens appear in sentence tokens
    // Try exact match first, then fuzzy (first token match only)
    let mut start_pos = None;
    for i in 0..sent_ids.len().saturating_sub(cand_ids.len() - 1) {
        if sent_ids[i] == cand_ids[0] {
            let matches = cand_ids.iter().enumerate().all(|(k, &cid)| {
                i + k < sent_ids.len() && sent_ids[i + k] == cid
            });
            if matches {
                start_pos = Some(i);
                break;
            }
        }
    }
    // Fallback: if exact match fails, find the candidate word in the sentence text
    // and use its character position to estimate the token position
    if start_pos.is_none() {
        let sent_lower = sentence.to_lowercase();
        let cand_lower = candidate.to_lowercase();
        if let Some(char_pos) = sent_lower.find(&cand_lower) {
            // Count tokens before this character position
            let prefix = &sentence[..char_pos];
            if let Ok(prefix_enc) = model.tokenizer.encode(prefix.to_string(), false) {
                let tok_pos = prefix_enc.get_ids().len();
                if tok_pos + cand_ids.len() <= sent_ids.len() {
                    start_pos = Some(tok_pos);
                }
            }
        }
    }

    let start = match start_pos {
        Some(s) => s,
        None => {
            // Last resort: fall back to sentence_score
            return sentence_score(model, sentence, candidate);
        }
    };

    // Direct token-level masking — no text decode/re-encode roundtrip.
    // Encode the full sentence with special tokens (CLS/SEP), find candidate tokens,
    // mask each one in the token array, forward pass with raw IDs.
    let enc_full = match model.tokenizer.encode(sentence.to_string(), true) {
        Ok(e) => e,
        Err(_) => return f32::NEG_INFINITY,
    };
    let full_ids: Vec<u32> = enc_full.get_ids().to_vec();

    // Find candidate tokens in the full encoding (with special tokens)
    let mut start_s = None;
    for i in 0..full_ids.len().saturating_sub(cand_ids.len().saturating_sub(1)) {
        if full_ids[i] == cand_ids[0] {
            let matches = cand_ids.iter().enumerate().all(|(k, &cid)| {
                i + k < full_ids.len() && full_ids[i + k] == cid
            });
            if matches {
                start_s = Some(i);
                break;
            }
        }
    }
    let start_pos = match start_s {
        Some(s) => s,
        None => return f32::NEG_INFINITY,
    };

    let mask_id = model.tokenizer.token_to_id("<mask>")
        .or_else(|| model.tokenizer.token_to_id("[MASK]"))
        .unwrap_or(4);
    let mut total: f32 = 0.0;
    let mut scored: usize = 0;
    for k in 0..cand_ids.len() {
        let pos = start_pos + k;
        if pos >= full_ids.len() { break; }
        let mut masked = full_ids.clone();
        masked[pos] = mask_id;
        if let Ok((logits, _)) = model.forward_ids(&masked, pos) {
            total += logits[cand_ids[k] as usize];
            scored += 1;
        }
    }
    if scored == 0 { return f32::NEG_INFINITY; }
    total / scored as f32
}

/// Phase 2a: BERT scoring only — no grammar.
///
/// The grammar batch used to live inside `score_and_rerank` and blocked the
/// caller's thread for hundreds of ms (one SWI-Prolog call per candidate
/// against a single-threaded actor). When the BERT worker called this
/// synchronously, completion requests piled up behind every spelling check.
///
/// This function does only what BERT can do: boundary score the candidates
/// in context, then hybrid sentence-rerank the top ones. Grammar filtering
/// is now the caller's job and runs on the grammar actor in parallel.
///
/// Returns BERT-ranked candidates (best first), ortho-weighted, ready for a
/// separate `apply_grammar_filter` pass.
pub fn bert_score_only(
    model: &mut Model,
    candidates: &[(String, f32)],
    context_before: &str,
    context_after: &str,
    sentence: &str,
) -> Vec<(String, f32)> {
    if candidates.is_empty() { return Vec::new(); }

    let all_cands: Vec<String> = candidates.iter().take(30).map(|(c, _)| c.clone()).collect();

    // Step 1: Boundary scoring
    let scored: Vec<(String, f32)> = match spelling::score_spelling(model, context_before, context_after, &all_cands) {
        Ok(result) => result.scored_candidates.into_iter()
            .map(|(c, s)| (c.trim().to_string(), s))
            .collect(),
        Err(_) => return candidates.to_vec(),
    };
    if scored.is_empty() { return candidates.to_vec(); }

    let ortho_map: HashMap<String, f32> = candidates.iter().cloned().collect();

    // Apply ortho weighting (sqrt) over BERT-scored
    let mut weighted: Vec<(String, f32)> = scored.iter().map(|(c, bert_score)| {
        let ct = c.trim().to_string();
        let ortho = ortho_map.get(ct.as_str()).copied().unwrap_or(0.5);
        (ct, bert_score * ortho.sqrt())
    }).collect();
    weighted.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

    // Step 4 (formerly Step 3 of score_and_rerank): hybrid sentence rerank
    if weighted.len() > 1 {
        let mut top_set: Vec<String> = Vec::new();
        let mut top_seen = HashSet::new();
        for (c, _) in weighted.iter().take(5) {
            if top_seen.insert(c.clone()) { top_set.push(c.clone()); }
        }
        for (c, _) in candidates.iter().take(15) {
            if top_seen.insert(c.clone()) { top_set.push(c.clone()); }
        }
        let top_set: Vec<String> = top_set.into_iter().map(|c| c.trim().to_string()).collect();
        let sentence_lower = sentence.to_lowercase();
        let sentence_cased = sentence.to_string();
        let ctx_before_lower = context_before.to_lowercase();
        let ctx_after_lower = context_after.to_lowercase();
        let word_lower = if let Some(pos) = sentence_lower.find(&ctx_before_lower) {
            let after_before = pos + ctx_before_lower.len();
            if let Some(apos) = sentence_lower[after_before..].find(&ctx_after_lower.trim_start()) {
                sentence_lower[after_before..after_before + apos].trim().to_string()
            } else {
                sentence_lower[after_before..].trim_end_matches(|c: char| c.is_ascii_punctuation() || c.is_whitespace()).to_string()
            }
        } else {
            all_cands.first().map(|c| c.to_lowercase()).unwrap_or_default()
        };
        let mut reranked: Vec<(String, f32)> = top_set.iter().map(|candidate| {
            let corrected_sent = if let Some(pos) = sentence_cased.to_lowercase().find(&word_lower) {
                format!("{}{}{}", &sentence_cased[..pos], candidate, &sentence_cased[pos + word_lower.len()..])
            } else {
                format!("{}{}{}", context_before, candidate, context_after)
            };
            let sent_score = subword_score(model, &corrected_sent, candidate);
            let ortho = ortho_map.get(candidate.as_str()).copied().unwrap_or(0.5);
            (candidate.clone(), sent_score * ortho.sqrt())
        }).collect();
        reranked.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
        promote_orthographic_anchor(&word_lower, &mut reranked, &ortho_map);
        reranked
    } else {
        weighted
    }
}

/// Phase 2b: Apply grammar filter to BERT-ranked candidates.
///
/// Runs on the main thread (or wherever the grammar response lands).
/// `grammar_results[i]` are the errors for candidate `bert_ranked[i]` after
/// substitution into the sentence. Caller must pre-build the test sentences
/// the same way `build_grammar_test_sentences` does so indices line up.
///
/// Behaviour matches the original score_and_rerank Step 2:
///   - Candidate with no errors → keep at current score.
///   - Candidate with error carrying a single-word suggestion → replace
///     the candidate with the suggestion (so "teksta" becomes "tekst"
///     after "en"). Original candidate kept at 0.8× score as fallback.
///   - Candidate with error and no usable suggestion → demote to 0.8×.
pub fn apply_grammar_filter(
    bert_ranked: &[(String, f32)],
    grammar_results: &[Vec<GrammarError>],
) -> Vec<(String, f32)> {
    if bert_ranked.is_empty() { return Vec::new(); }
    let mut out: Vec<(String, f32)> = Vec::new();
    let mut seen = HashSet::new();
    for ((candidate, score), errs) in bert_ranked.iter().zip(grammar_results.iter()) {
        if errs.is_empty() {
            if seen.insert(candidate.clone()) {
                out.push((candidate.clone(), *score));
            }
            continue;
        }
        for err in errs {
            if !err.suggestion.is_empty() && err.suggestion != *candidate && !err.suggestion.contains('|') {
                let sug = err.suggestion.to_lowercase();
                if seen.insert(sug.clone()) {
                    out.push((sug, *score));
                }
            }
        }
        if seen.insert(candidate.clone()) {
            out.push((candidate.clone(), score * 0.8));
        }
    }
    out.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
    out
}

/// Build the test sentences `apply_grammar_filter` expects: one sentence
/// per candidate with the candidate substituted in place of the misspelled
/// word. The grammar actor batches these and returns per-sentence errors.
pub fn build_grammar_test_sentences(
    bert_ranked: &[(String, f32)],
    context_before: &str,
    context_after: &str,
) -> Vec<String> {
    let last_start = context_before.rfind(|c: char| ".!?".contains(c))
        .map(|i| i + 1).unwrap_or(0);
    let fragment = context_before[last_start..].trim();
    bert_ranked.iter()
        .map(|(c, _)| if context_after.is_empty() {
            format!("{} {}.", fragment, c)
        } else {
            format!("{} {} {}", fragment, c, context_after.trim_start())
        })
        .collect()
}

/// Phase 2: BERT scoring + grammar correction + hybrid sentence re-ranking.
/// `grammar_check` takes sentences and returns errors per sentence.
///
/// Kept as a synchronous wrapper for `test_spelling`, `dyslexia_tests`,
/// and any caller that doesn't have the worker/actor async plumbing.
/// Production GUI uses `bert_score_only` + `apply_grammar_filter` over
/// the async grammar actor instead — see `feedback_*` memory.
pub fn score_and_rerank(
    model: &mut Model,
    grammar_check: &mut dyn FnMut(&[String]) -> Vec<Vec<GrammarError>>,
    candidates: &[(String, f32)],
    context_before: &str,
    context_after: &str,
    sentence: &str,
) -> Vec<(String, f32)> {
    if candidates.is_empty() { return Vec::new(); }

    let all_cands: Vec<String> = candidates.iter().take(30).map(|(c, _)| c.clone()).collect();

    // Step 1: Boundary scoring (~24ms)
    let scored: Vec<(String, f32)> = match spelling::score_spelling(model, context_before, context_after, &all_cands) {
        Ok(result) => result.scored_candidates.into_iter()
            .map(|(c, s)| (c.trim().to_string(), s)) // trim BPE artifacts
            .collect(),
        Err(_) => return candidates.to_vec(),
    };

    // Step 2: Grammar correction — replace candidates with grammar-suggested forms
    let last_start = context_before.rfind(|c: char| ".!?".contains(c))
        .map(|i| i + 1).unwrap_or(0);
    let fragment = context_before[last_start..].trim();
    let test_sentences: Vec<String> = scored.iter()
        .map(|(c, _)| if context_after.is_empty() {
            format!("{} {}.", fragment, c)
        } else {
            format!("{} {} {}", fragment, c, context_after.trim_start())
        })
        .collect();
    let grammar_results = grammar_check(&test_sentences);

    let mut corrected: Vec<(String, f32)> = Vec::new();
    let mut grammar_suggested = HashSet::new();
    let mut seen = HashSet::new();
    for ((candidate, score), errs) in scored.into_iter().zip(grammar_results.iter()) {
        if errs.is_empty() {
            if seen.insert(candidate.clone()) {
                corrected.push((candidate, score));
            }
        } else {
            for err in errs {
                if !err.suggestion.is_empty() && err.suggestion != candidate && !err.suggestion.contains('|') {
                    let sug = err.suggestion.to_lowercase();
                    if seen.insert(sug.clone()) {
                        grammar_suggested.insert(sug.clone());
                        corrected.push((sug, score));
                    }
                }
            }
            if seen.insert(candidate.clone()) {
                corrected.push((candidate, score * 0.8));
            }
        }
    }

    if corrected.is_empty() { return candidates.to_vec(); }

    // Re-score corrected candidates with boundary scorer
    let all_corrected: Vec<String> = corrected.iter().map(|(c, _)| c.clone()).collect();
    let boundary_scored: Vec<(String, f32)> = match spelling::score_spelling(model, context_before, context_after, &all_corrected) {
        Ok(result) => result.scored_candidates.into_iter()
            .map(|(c, s)| (c.trim().to_string(), s)).collect(),
        Err(_) => corrected,
    };

    // Build ortho map from original candidates
    let ortho_map: HashMap<String, f32> = candidates.iter().cloned().collect();

    // Apply ortho weighting (trim any BPE space artifacts)
    let mut weighted: Vec<(String, f32)> = boundary_scored.iter().map(|(c, bert_score)| {
        let ct = c.trim().to_string();
        let ortho = ortho_map.get(ct.as_str()).copied().unwrap_or(0.5);
        let eff = if grammar_suggested.contains(&ct) { 1.0 } else { ortho };
        (ct, bert_score * eff.sqrt())
    }).collect();
    weighted.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

    // Step 3: Hybrid — sentence scoring on top candidates
    // Only score single-token candidates with sentence_score (multi-token get inflated prefix logits)
    // For multi-token candidates, keep the boundary score
    if weighted.len() > 1 {
        let mut top_set: Vec<String> = Vec::new();
        let mut top_seen = HashSet::new();
        for (c, _) in weighted.iter().take(5) {
            if top_seen.insert(c.clone()) { top_set.push(c.clone()); }
        }
        for (c, _) in candidates.iter().take(15) {
            if top_seen.insert(c.clone()) { top_set.push(c.clone()); }
        }
        let top_set: Vec<String> = top_set.into_iter().map(|c| c.trim().to_string()).collect();
        let sentence_lower = sentence.to_lowercase();
        // Use original casing for BERT (case-sensitive model)
        let sentence_cased = sentence.to_string();
        // Extract misspelled word from gap between context_before and context_after
        let ctx_before_lower = context_before.to_lowercase();
        let ctx_after_lower = context_after.to_lowercase();
        let word_lower = if let Some(pos) = sentence_lower.find(&ctx_before_lower) {
            let after_before = pos + ctx_before_lower.len();
            if let Some(apos) = sentence_lower[after_before..].find(&ctx_after_lower.trim_start()) {
                sentence_lower[after_before..after_before + apos].trim().to_string()
            } else {
                sentence_lower[after_before..].trim_end_matches(|c: char| c.is_ascii_punctuation() || c.is_whitespace()).to_string()
            }
        } else {
            all_cands.first().map(|c| c.to_lowercase()).unwrap_or_default()
        };
        let weighted_map: HashMap<String, f32> = weighted.iter().cloned().collect();
        let mut reranked: Vec<(String, f32)> = top_set.iter().map(|candidate| {
            let corrected_sent = if let Some(pos) = sentence_cased.to_lowercase().find(&word_lower) {
                format!("{}{}{}", &sentence_cased[..pos], candidate, &sentence_cased[pos + word_lower.len()..])
            } else {
                format!("{}{}{}", context_before, candidate, context_after)
            };
            let sent_score = subword_score(model, &corrected_sent, candidate);
            let ortho = ortho_map.get(candidate.as_str()).copied().unwrap_or(0.5);
            let eff = if grammar_suggested.contains(candidate) { 1.0 } else { ortho };
            (candidate.clone(), sent_score * eff.sqrt())
        }).collect();
        reranked.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
        promote_orthographic_anchor(&word_lower, &mut reranked, &ortho_map);
        reranked
    } else {
        weighted
    }
}
