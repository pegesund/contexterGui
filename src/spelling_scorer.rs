//! Shared spelling suggestion pipeline.
//! Used by both the app (main.rs + bert_worker.rs) and tests (test_spelling.rs).
//!
//! Phase 1: `generate_spelling_candidates` — candidate generation + ortho scoring + dict filter
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
pub fn try_split_function_word(word: &str, analyzer: &mtag::Analyzer) -> Option<String> {
    // Don't split valid compound words (avstand, tilstand, imorgen, etc.)
    if analyzer.has_word(&word.to_lowercase()) {
        return None;
    }
    const FUNCTION_WORDS: &[&str] = &[
        "gjennom", "mellom", "under", "etter", "langs", "rundt",
        "foran", "bortover", "innover", "utover",
        "forbi", "siden", "etter", "blant",
        "over", "inne", "borte",
        "uten", "utenfor", "innenfor",
        "med", "mot", "ved", "hos", "fra",
        "for", "som", "men",
        "til", "per", "via",
        "på", "av", "om",
        "en", "et", "ei",
        "og", "at",
        "i",
    ];
    let lower = word.to_lowercase();
    for prefix in FUNCTION_WORDS {
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
pub fn compute_boost(
    word: &str,
    doc_word_counts: &HashMap<String, u16>,
    user_dict_words: &[String],
    wordfreq: Option<&HashMap<String, u64>>,
) -> f32 {
    let lower = word.to_lowercase();
    const COMMON_THRESHOLD: u64 = 40_000;
    if wordfreq.and_then(|wf| wf.get(&lower)).map_or(false, |&f| f >= COMMON_THRESHOLD) {
        return 1.0;
    }
    let in_doc = doc_word_counts.get(&lower).copied().unwrap_or(0) >= 2;
    let in_user = user_dict_words.iter().any(|uw| uw.eq_ignore_ascii_case(&lower));
    match (in_doc, in_user) {
        (true, true)   => 1.6,
        (false, true)  => 1.3,
        (true, false)  => 1.25,
        (false, false) => 1.0,
    }
}

/// Phase 1: Generate spelling candidates, ortho-score them, and dictionary-filter.
/// Returns ortho-scored, dictionary-valid candidates sorted best-first.
pub fn generate_spelling_candidates(
    analyzer: &mtag::Analyzer,
    wordfreq: Option<&HashMap<String, u64>>,
    user_dict_words: &[String],
    doc_word_counts: &HashMap<String, u16>,
    word: &str,
    _sentence_ctx: &str,
) -> Vec<(String, f32)> {
    let word_lower = word.to_lowercase();
    let word_trigrams = trigrams(&word_lower);
    let word_first = word_lower.chars().next().unwrap_or(' ');

    let mut candidates: Vec<String> = Vec::new();
    let mut seen = HashSet::new();
    let mut edit_distances: HashMap<String, u32> = HashMap::new();

    // Source 1: Fuzzy Levenshtein (distance 2)
    for (w, dist) in analyzer.fuzzy_lookup(&word_lower, 2) {
        let wl = w.to_lowercase();
        if wl == word_lower || wl.len() < 2 { continue; }
        edit_distances.insert(wl.clone(), dist);
        if seen.insert(wl.clone()) { candidates.push(wl); }
    }

    // Source 2: Prefix lookup (missing-letter typos)
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

    // Source 6: Split function word
    if let Some(split) = try_split_function_word(&word_lower, analyzer) {
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

    // Source 9: Long word truncation (>= 10 chars)
    if word_lower.len() >= 10 {
        let is_known_or_compound = |w: &str| -> bool {
            if analyzer.has_word(w) { return true; }
            for j in 3..w.len().saturating_sub(2) {
                if !w.is_char_boundary(j) { continue; }
                let left = &w[..j];
                let right = &w[j..];
                if right.len() >= 3 && analyzer.has_word(left) && analyzer.has_word(right) { return true; }
                if right.starts_with('s') && right.len() > 3 && analyzer.has_word(left) && analyzer.has_word(&right[1..]) { return true; }
            }
            false
        };
        for strip in 1..=2usize {
            if word_lower.is_char_boundary(strip) {
                let trimmed = &word_lower[strip..];
                if trimmed.len() >= 5 && is_known_or_compound(trimmed) && seen.insert(trimmed.to_string()) {
                    edit_distances.insert(trimmed.to_string(), strip as u32);
                    candidates.push(trimmed.to_string());
                }
            }
            let end = word_lower.len() - strip;
            if word_lower.is_char_boundary(end) {
                let trimmed = &word_lower[..end];
                if trimmed.len() >= 5 && is_known_or_compound(trimmed) && seen.insert(trimmed.to_string()) {
                    edit_distances.insert(trimmed.to_string(), strip as u32);
                    candidates.push(trimmed.to_string());
                }
            }
        }
    }

    // Source 10: First-character swap
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

    // Source 11: Inflected forms of candidates
    {
        use mtag::types::{Pos, Tag};
        let base_candidates: Vec<String> = candidates.clone();
        for base in &base_candidates {
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
                                    candidates.push(fl);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Phase 2: Ortho score
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
        ortho_sim *= compute_boost(w, doc_word_counts, user_dict_words, wordfreq);
        ortho_scored.push((w.clone(), ortho_sim));
    }
    ortho_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

    // Phase 4: Dictionary filter (top 15)
    let mut passing: Vec<(String, f32)> = Vec::new();
    let mut checked = 0;
    for (candidate, score) in &ortho_scored {
        if checked >= 15 { break; }
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

/// Score a candidate word in context. Same as boundary scorer but called
/// with the full corrected sentence for context.
/// This is used for the hybrid re-ranking step.
pub fn sentence_score(model: &mut Model, sentence: &str) -> f32 {
    let words: Vec<&str> = sentence.split_whitespace().collect();
    if words.is_empty() { return f32::NEG_INFINITY; }
    let mut total: f32 = 0.0;
    let mut scored_count: usize = 0;
    for i in 0..words.len() {
        let word_clean = words[i].trim_matches(|c: char| c.is_ascii_punctuation());
        if word_clean.is_empty() { continue; }
        let masked: String = words.iter().enumerate()
            .map(|(j, w)| if j == i { "<mask>" } else { *w })
            .collect::<Vec<_>>().join(" ");
        if let Ok((logits, _)) = model.single_forward(&masked) {
            if let Ok(enc) = model.tokenizer.encode(format!(" {}", word_clean.to_lowercase()), false) {
                let ids = enc.get_ids();
                if let Some(&first_id) = ids.first() {
                    total += logits[first_id as usize];
                    scored_count += 1;
                }
            }
        }
    }
    if scored_count == 0 { return f32::NEG_INFINITY; }
    total / scored_count as f32
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

    let start = match start_pos {
        Some(s) => s,
        None => return f32::NEG_INFINITY,
    };

    // Mask each subword token and sum logits
    let mask_id = model.tokenizer.token_to_id("<mask>").unwrap_or(4);
    let mut total: f32 = 0.0;
    let mut scored: usize = 0;
    for k in 0..cand_ids.len() {
        let pos = start + k;
        if pos >= sent_ids.len() { break; }
        let mut masked_ids = sent_ids.clone();
        masked_ids[pos] = mask_id;
        // Build masked text from token IDs
        let masked_text = model.tokenizer.decode(&masked_ids, true)
            .unwrap_or_default();
        if let Ok((logits, _)) = model.single_forward(&masked_text) {
            total += logits[cand_ids[k] as usize];
            scored += 1;
        }
    }
    if scored == 0 { return f32::NEG_INFINITY; }
    total / scored as f32
}

/// Phase 2: BERT scoring + grammar correction + hybrid sentence re-ranking.
/// `grammar_check` takes sentences and returns errors per sentence.
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
                if !err.suggestion.is_empty() && err.suggestion != candidate {
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
        for (c, _) in candidates.iter().take(5) {
            if top_seen.insert(c.clone()) { top_set.push(c.clone()); }
        }
        let top_set: Vec<String> = top_set.into_iter().map(|c| c.trim().to_string()).collect();
        let sentence_lower = sentence.to_lowercase();
        let word_lower = all_cands.first().map(|c| c.to_lowercase()).unwrap_or_default();
        let weighted_map: HashMap<String, f32> = weighted.iter().cloned().collect();
        let mut reranked: Vec<(String, f32)> = top_set.iter().map(|candidate| {
            // Check if candidate is a single BPE token
            let n_tokens = model.tokenizer.encode(format!(" {}", candidate.to_lowercase()), false)
                .ok().map(|enc| enc.get_ids().len()).unwrap_or(0);
            let corrected_sent = if let Some(pos) = sentence_lower.find(&word_lower) {
                format!("{}{}{}", &sentence_lower[..pos], candidate, &sentence_lower[pos + word_lower.len()..])
            } else {
                format!("{}{}{}", context_before, candidate, context_after)
            };
            let sent_score = if n_tokens == 1 {
                // Single token: standard sentence scoring
                sentence_score(model, &corrected_sent)
            } else {
                // Multi-token: score each subword token at its position
                subword_score(model, &corrected_sent, candidate)
            };
            let ortho = ortho_map.get(candidate.as_str()).copied().unwrap_or(0.5);
            let eff = if grammar_suggested.contains(candidate) { 1.0 } else { ortho };
            (candidate.clone(), sent_score * eff.sqrt())
        }).collect();
        reranked.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
        reranked
    } else {
        weighted
    }
}
