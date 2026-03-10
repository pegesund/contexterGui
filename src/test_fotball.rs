//! Console test replicating the exact GUI completion pipeline.
//! Tests that "fotba" → "fotball" appears in completions.

use std::collections::{HashMap, HashSet};
use std::path::PathBuf;

fn data_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../contexter-repo/training-data")
}

fn main() -> anyhow::Result<()> {
    let ort_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../onnxruntime/onnxruntime-win-x64-1.23.0/lib/onnxruntime.dll");
    unsafe { std::env::set_var("ORT_DYLIB_PATH", &ort_path); }

    let data = data_dir();
    let onnx = data.join("onnx/norbert4_base_int8.onnx");
    let tok = data.join("onnx/tokenizer.json");
    let wf_path = data.join("wordfreq.tsv");

    eprintln!("Loading NorBERT4...");
    let mut model = nostos_cognio::model::Model::load(onnx.to_str().unwrap(), tok.to_str().unwrap())?;
    eprintln!("Building prefix index...");
    let pi = nostos_cognio::prefix_index::build_prefix_index(&model.tokenizer);
    let wf = nostos_cognio::wordfreq::load_wordfreq(wf_path.as_path(), 10);

    // Load mtag analyzer directly for prefix_lookup fallback
    eprintln!("Loading mtag analyzer...");
    let dict_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../rustSpell/mtag-rs/data/fullform_bm.mfst");
    let analyzer = mtag::Analyzer::new(dict_path.to_str().unwrap())
        .map_err(|e| anyhow::anyhow!("{}", e))?;
    eprintln!("Ready.\n");

    // ── Test cases ──
    let tests = vec![
        // Simple context
        ("Jeg liker å spille", "fotba", "fotball"),
        // Multi-sentence context where "fotball" appears in prior sentence (the bug!)
        ("Jeg spiller fotball. Jeg liker å spille", "fotba", "fotball"),
        ("Han er en god", "fotba", "fotball"),
        ("Jeg liker å", "s", "spille"),
        ("Hun er veldig", "f", "flink"),
    ];

    for (context, prefix, expected) in &tests {
        println!("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
        println!("Context: '{}'  prefix: '{}'  expected: '{}'", context, prefix, expected);
        println!("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");

        let prefix_lower = prefix.to_lowercase();
        let capitalize = prefix.chars().next().map_or(false, |c| c.is_uppercase());

        // Build masked sentence (same as GUI: bridge::build_masked_sentence)
        // In GUI: before = "Jeg liker å spille fotba", after = ""
        // Word is stripped from before, mask is glued (no space before <mask>)
        // Result: "Jeg liker å spille<mask> ."
        let masked = format!("{}<mask> .", context);
        println!("Masked: '{}'", masked);

        // Step 1: BPE prefix index lookup (same as GUI line 2744-2747)
        let matches: Vec<(u32, String)> = pi.get(&prefix_lower)
            .cloned()
            .unwrap_or_default();
        println!("BPE matches for '{}': {} tokens", prefix_lower, matches.len());
        if matches.len() <= 10 {
            for (tid, w) in &matches {
                println!("  token {} = '{}'", tid, w);
            }
        } else {
            for (tid, w) in matches.iter().take(5) {
                println!("  token {} = '{}'", tid, w);
            }
            println!("  ... and {} more", matches.len() - 5);
        }

        // Step 2: mtag fallback (same as GUI line 2748-2752)
        let mtag_candidates: Vec<String> = if matches.is_empty() && !prefix.is_empty() {
            let cands = analyzer.prefix_lookup(&prefix_lower, 50);
            println!("mtag fallback: {} candidates", cands.len());
            for w in cands.iter().take(10) {
                println!("  '{}'", w);
            }
            cands
        } else {
            println!("mtag fallback: SKIPPED (BPE has {} matches)", matches.len());
            vec![]
        };

        // Step 3: Nearby words — only from current sentence (not prior context sentences)
        let before_mask = masked.split("<mask>").next().unwrap_or("");
        let sent_start = before_mask.rfind(|c: char| ".!?".contains(c))
            .map(|i| i + 1).unwrap_or(0);
        let current_sent = &before_mask[sent_start..];
        let nearby_words: HashSet<String> = current_sent.split_whitespace()
            .map(|w| w.trim_matches(|c: char| !c.is_alphanumeric()).to_lowercase())
            .filter(|w| w.len() > 1)
            .collect();
        println!("Nearby words (excluded): {:?}", nearby_words);

        // Step 4: single_forward (same as GUI line 2810)
        let t0 = std::time::Instant::now();
        let logits = match model.single_forward(&masked) {
            Ok((l, _)) => l,
            Err(e) => { println!("ERROR: forward failed: {}", e); continue; }
        };
        let forward_ms = t0.elapsed().as_millis();
        println!("Forward pass: {}ms ({} chars context)", forward_ms, masked.len());

        let cancel = std::sync::atomic::AtomicBool::new(false);

        // Step 5: Build left completions (same as GUI line 2817-2823)
        let left = if matches.is_empty() && !prefix_lower.is_empty() {
            println!("Using mtag path ({} candidates)...", mtag_candidates.len());
            build_mtag_completions(&mut model, &masked, &mtag_candidates, &logits, capitalize, &cancel)
        } else if !prefix_lower.is_empty() {
            println!("Using BPE path ({} matches)...", matches.len());
            build_bpe_completions(&mut model, &masked, &prefix_lower, &matches, &logits, Some(&wf), &nearby_words, capitalize, &cancel)
        } else {
            vec![]
        };

        // Step 6: Right completions
        let left_words: HashSet<String> = left.iter().map(|c| c.word.to_lowercase()).collect();
        let right = build_right_completions(&model, &logits, Some(&wf), &nearby_words, &left_words);

        println!("\nLeft completions:");
        for (i, c) in left.iter().enumerate().take(10) {
            let marker = if c.word.to_lowercase() == expected.to_lowercase() { " ◄◄◄" } else { "" };
            println!("  {:>2}. {} ({:.1}){}", i + 1, c.word, c.score, marker);
        }
        println!("Right completions:");
        for (i, c) in right.iter().enumerate().take(5) {
            println!("  {:>2}. {} ({:.1})", i + 1, c.word, c.score);
        }

        let found = left.iter().any(|c| c.word.to_lowercase() == expected.to_lowercase());
        println!("\n{} '{}' {} in completions\n",
            if found { "✓" } else { "✗" },
            expected,
            if found { "FOUND" } else { "NOT FOUND" });
    }

    Ok(())
}

// ── Exact copies of GUI pipeline functions ──

use nostos_cognio::complete::Completion;
use nostos_cognio::model::Model;

fn build_bpe_completions(
    model: &mut Model,
    masked: &str,
    prefix_lower: &str,
    matches: &[(u32, String)],
    logits: &[f32],
    wordfreq: Option<&HashMap<String, u64>>,
    nearby_words: &HashSet<String>,
    capitalize: bool,
    cancel: &std::sync::atomic::AtomicBool,
) -> Vec<Completion> {
    use std::sync::atomic::Ordering;

    let is_valid = |w: &str| -> bool {
        let key = w.to_lowercase();
        if nearby_words.contains(&key) { return false; }
        wordfreq.map_or(true, |wf| wf.contains_key(&key))
    };
    let cap = |s: &str| -> String {
        let mut c = s.chars();
        match c.next() {
            None => String::new(),
            Some(f) => f.to_uppercase().to_string() + c.as_str(),
        }
    };

    let mut token_scored: Vec<(String, f32)> = matches.iter()
        .map(|(tid, word)| (word.clone(), logits[*tid as usize]))
        .collect();
    token_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

    struct Candidate {
        token_ids: Vec<u32>,
        word: String,
        score: f32,
        done: bool,
    }
    let mut candidate_set: HashSet<String> = HashSet::new();
    let mut candidates: Vec<Candidate> = Vec::new();

    for (tok_word, tok_score) in token_scored.iter().take(20) {
        if candidate_set.insert(tok_word.clone()) {
            if let Some((tid, _)) = matches.iter().find(|(_, w)| w == tok_word) {
                candidates.push(Candidate {
                    token_ids: vec![*tid],
                    word: tok_word.clone(),
                    score: *tok_score,
                    done: false,
                });
            }
        }
    }
    let mut long_tokens: Vec<&(String, f32)> = token_scored.iter()
        .filter(|(w, s)| w.len() >= 5 && *s > 0.0)
        .collect();
    long_tokens.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
    for (tok_word, tok_score) in long_tokens.iter().take(20) {
        if candidate_set.insert(tok_word.clone()) {
            if let Some((tid, _)) = matches.iter().find(|(_, w)| w == tok_word) {
                candidates.push(Candidate {
                    token_ids: vec![*tid],
                    word: tok_word.clone(),
                    score: *tok_score,
                    done: false,
                });
            }
        }
    }

    let mask_parts: Vec<&str> = masked.splitn(2, "<mask>").collect();
    let ctx_before = mask_parts[0].trim_end();
    let ctx_after = mask_parts.get(1).map(|s| s.trim_start()).unwrap_or(".");

    let max_steps = if prefix_lower.len() <= 3 { 1 } else { 0 };
    for _step in 0..max_steps {
        if cancel.load(Ordering::Acquire) { return vec![]; }
        let best_score = candidates.iter()
            .filter(|c| !c.done)
            .map(|c| c.score)
            .fold(f32::NEG_INFINITY, f32::max);
        let threshold = best_score - 15.0;
        let mut to_extend: Vec<usize> = candidates.iter().enumerate()
            .filter(|(_, c)| !c.done && c.score >= threshold)
            .map(|(i, _)| i)
            .collect();
        for c in candidates.iter_mut() {
            if !c.done && c.score < threshold { c.done = true; }
        }
        let batch_cap = if prefix_lower.len() <= 2 { 5 } else { 10 };
        to_extend.truncate(batch_cap);
        if to_extend.is_empty() { break; }

        let batch_texts: Vec<String> = to_extend.iter()
            .map(|&i| {
                let accumulated = model.tokenizer
                    .decode(&candidates[i].token_ids, false)
                    .unwrap_or_default();
                let accumulated = accumulated.trim();
                format!("{} {}<mask> {}", ctx_before, accumulated, ctx_after)
            })
            .collect();

        match model.batched_forward_argmax(&batch_texts) {
            Ok((argmaxes, _)) => {
                for (k, &i) in to_extend.iter().enumerate() {
                    let best_id = argmaxes[k];
                    let best_token = &model.id_to_token[best_id as usize];
                    let is_continuation = !best_token.starts_with('Ġ')
                        && !matches!(best_token.as_str(), "." | "," | "!" | "?" | ";" | ":");
                    if is_continuation {
                        candidates[i].token_ids.push(best_id);
                        candidates[i].word = model.tokenizer
                            .decode(&candidates[i].token_ids, false)
                            .unwrap_or_default().trim().to_string();
                    } else {
                        candidates[i].done = true;
                    }
                }
            }
            Err(_) => break,
        }
    }

    // Debug: show all candidates before filtering
    println!("  BPE candidates before is_valid filter:");
    for c in &candidates {
        let valid = is_valid(&c.word.to_lowercase());
        println!("    '{}' score={:.1} done={} valid={}", c.word, c.score, c.done, valid);
    }

    let mut left_scored: Vec<(String, f32)> = Vec::new();
    let mut seen_words: HashSet<String> = HashSet::new();
    for c in &candidates {
        let key = c.word.to_lowercase();
        if is_valid(&key) && seen_words.insert(key.clone()) {
            left_scored.push((key, c.score));
        }
    }

    left_scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
    left_scored.into_iter()
        .take(25)
        .map(|(w, s)| Completion {
            word: if capitalize { cap(&w) } else { w },
            score: s,
            elapsed_ms: 0.0,
        })
        .collect()
}

fn build_mtag_completions(
    model: &mut Model,
    masked: &str,
    mtag_candidates: &[String],
    logits: &[f32],
    capitalize: bool,
    cancel: &std::sync::atomic::AtomicBool,
) -> Vec<Completion> {
    use std::sync::atomic::Ordering;
    if mtag_candidates.is_empty() { return vec![]; }

    let cap = |s: &str| -> String {
        let mut c = s.chars();
        match c.next() {
            None => String::new(),
            Some(f) => f.to_uppercase().to_string() + c.as_str(),
        }
    };

    let mask_parts: Vec<&str> = masked.splitn(2, "<mask>").collect();
    let ctx_before = mask_parts[0].trim_end();
    let ctx_after = mask_parts.get(1).map(|s| s.trim_start()).unwrap_or(".");

    let candidates_with_tokens: Vec<(String, Vec<u32>)> = mtag_candidates.iter()
        .filter_map(|w| {
            let enc = model.tokenizer.encode(format!(" {}", w).as_str(), false).ok()?;
            let ids: Vec<u32> = enc.get_ids().to_vec();
            if ids.is_empty() { return None; }
            Some((w.clone(), ids))
        })
        .collect();

    let mut scores: Vec<f32> = candidates_with_tokens.iter()
        .map(|(_, ids)| logits[ids[0] as usize])
        .collect();

    let max_tokens = candidates_with_tokens.iter().map(|(_, ids)| ids.len()).max().unwrap_or(1);
    for t in 1..max_tokens {
        if cancel.load(Ordering::Acquire) { return vec![]; }
        let to_score: Vec<usize> = candidates_with_tokens.iter().enumerate()
            .filter(|(_, (_, ids))| ids.len() > t)
            .map(|(i, _)| i)
            .collect();
        if to_score.is_empty() { break; }

        let mut unique_prefixes: Vec<Vec<u32>> = Vec::new();
        let mut prefix_to_idx: HashMap<Vec<u32>, usize> = HashMap::new();
        let mut candidate_to_prefix: Vec<usize> = Vec::new();
        for &i in &to_score {
            let token_prefix = candidates_with_tokens[i].1[..t].to_vec();
            let pidx = if let Some(&existing) = prefix_to_idx.get(&token_prefix) {
                existing
            } else {
                let idx = unique_prefixes.len();
                prefix_to_idx.insert(token_prefix.clone(), idx);
                unique_prefixes.push(token_prefix);
                idx
            };
            candidate_to_prefix.push(pidx);
        }

        let batch_texts: Vec<String> = unique_prefixes.iter()
            .map(|ids| {
                let partial = model.tokenizer.decode(ids, false).unwrap_or_default();
                format!("{} {}<mask> {}", ctx_before, partial.trim(), ctx_after)
            })
            .collect();

        if let Ok((batch_logits, _)) = model.batched_forward(&batch_texts) {
            for (k, &i) in to_score.iter().enumerate() {
                let pidx = candidate_to_prefix[k];
                scores[i] += batch_logits[pidx][candidates_with_tokens[i].1[t] as usize];
            }
        }
    }

    // Debug: show mtag candidate scores
    let mut debug: Vec<(String, f32, usize)> = candidates_with_tokens.iter().enumerate()
        .map(|(i, (w, ids))| (w.clone(), scores[i] / ids.len() as f32, ids.len()))
        .collect();
    debug.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
    println!("  mtag candidates scored:");
    for (w, s, ntok) in debug.iter().take(15) {
        println!("    '{}' score={:.1} tokens={}", w, s, ntok);
    }

    let mut scored: Vec<(String, f32)> = candidates_with_tokens.iter().enumerate()
        .map(|(i, (w, ids))| (w.clone(), scores[i] / ids.len() as f32))
        .collect();
    scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
    scored.into_iter()
        .take(25)
        .map(|(w, s)| Completion {
            word: if capitalize { cap(&w) } else { w },
            score: s,
            elapsed_ms: 0.0,
        })
        .collect()
}

fn build_right_completions(
    model: &Model,
    logits: &[f32],
    wordfreq: Option<&HashMap<String, u64>>,
    nearby_words: &HashSet<String>,
    left_words: &HashSet<String>,
) -> Vec<Completion> {
    let mut right: Vec<Completion> = Vec::new();
    // Top tokens by logit that are word-initial and in wordfreq
    let mut scored: Vec<(usize, f32)> = logits.iter().enumerate()
        .map(|(i, &s)| (i, s))
        .collect();
    scored.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));

    for (tid, score) in scored.iter().take(200) {
        let token = &model.id_to_token[*tid];
        if !token.starts_with('Ġ') { continue; }
        let word = model.tokenizer.decode(&[*tid as u32], false)
            .unwrap_or_default().trim().to_string();
        let key = word.to_lowercase();
        if key.len() < 2 { continue; }
        if nearby_words.contains(&key) { continue; }
        if left_words.contains(&key) { continue; }
        if let Some(wf) = wordfreq {
            if !wf.contains_key(&key) { continue; }
        }
        right.push(Completion { word: key, score: *score, elapsed_ms: 0.0 });
        if right.len() >= 10 { break; }
    }
    right
}
