//! Compound word FST walker.
//!
//! Walks the FST node-by-node to find compound word decompositions
//! with fuzzy matching per part. Handles Norwegian compound words
//! like "kjøkkenbord" = "kjøkken" + "bord" even when misspelled.

use fst::raw::{CompiledAddr, Fst, Node};
use std::collections::{BinaryHeap, HashMap};
use std::cmp::Reverse;

/// A matched part of a compound word.
#[derive(Debug, Clone)]
pub struct CompoundPart {
    pub matched_word: String,
    pub input_start: usize,  // byte offset in input
    pub input_end: usize,
    pub edits: u32,
}

/// A complete compound decomposition result.
#[derive(Debug, Clone)]
pub struct CompoundResult {
    pub parts: Vec<CompoundPart>,
    pub total_edits: u32,
    pub compound_word: String, // parts joined
}

/// Internal walk state.
#[derive(Clone)]
struct WalkState {
    fst_addr: CompiledAddr,
    input_pos: usize,
    edits: u32,              // edits for current part
    word_bytes: Vec<u8>,     // FST bytes matched for current part
    word_start: usize,       // input byte offset where current part started
    parts: Vec<CompoundPart>,
    total_edits: u32,
}

impl WalkState {
    fn priority(&self) -> u32 {
        // Lower is better: total edits + current part edits
        self.total_edits + self.edits
    }
}

impl PartialEq for WalkState {
    fn eq(&self, other: &Self) -> bool { self.priority() == other.priority() }
}
impl Eq for WalkState {}
impl PartialOrd for WalkState {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> { Some(self.cmp(other)) }
}
impl Ord for WalkState {
    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        Reverse(self.priority()).cmp(&Reverse(other.priority()))
    }
}

const MAX_EDITS_PER_PART: u32 = 2;
const MAX_TOTAL_EDITS: u32 = 4;
const MAX_PARTS: usize = 3;
const MIN_PART_BYTES: usize = 3;
const BEAM_WIDTH: usize = 5000;
const BINDING_LETTERS: &[u8] = b"se";

/// Norwegian phonetic vowel pairs common in dyslexic writing.
/// These substitutions cost 0 edits since they're near-equivalences.
/// ascii_byte ↔ UTF-8 continuation byte (after 0xC3 prefix)
#[inline]
fn is_free_vowel_swap(ascii: u8, utf8_cont: u8) -> bool {
    matches!((ascii, utf8_cont),
        (b'e', 0xB8) |  // e ↔ ø
        (b'o', 0xA5) |  // o ↔ å
        (b'a', 0xA6)    // a ↔ æ
    )
}

/// Check if a word's suffix triggers binding -s- in Norwegian compounds.
/// Based on morphological rules: -ing, -ning, -het, -skap, -sjon, -tet,
/// -dom, -else, -sel, -nad, -itet, -leik always take -s-.
#[inline]
fn needs_binding_s(word: &[u8]) -> bool {
    word.ends_with(b"ing") || word.ends_with(b"het") || word.ends_with(b"skap") ||
    word.ends_with(b"sjon") || word.ends_with(b"tet") || word.ends_with(b"dom") ||
    word.ends_with(b"else") || word.ends_with(b"sel") || word.ends_with(b"nad") ||
    word.ends_with(b"leik")
}

/// Quick exact-match walk from root for the remaining input bytes.
/// Returns all accepting words found (word, end_pos) without any edits.
/// Used as a lookahead to bypass beam pruning for exact second parts.
fn exact_walk_from_root<D: AsRef<[u8]>>(fst: &Fst<D>, input: &[u8], start: usize) -> Vec<(String, usize)> {
    let mut results = Vec::new();
    let mut node = fst.root();
    let mut wb = Vec::new();
    for i in start..input.len() {
        if let Some(idx) = node.find_input(input[i]) {
            let trans = node.transition(idx);
            wb.push(trans.inp);
            node = fst.node(trans.addr);
            if node.is_final() && wb.len() >= MIN_PART_BYTES {
                results.push((String::from_utf8_lossy(&wb).to_string(), i + 1));
            }
        } else {
            break;
        }
    }
    results
}

/// Walk the FST to find compound word decompositions for the input.
/// Returns decompositions sorted by total edit distance (best first).
pub fn compound_fuzzy_walk<D: AsRef<[u8]>>(
    fst: &Fst<D>,
    input: &str,
    wordfreq: Option<&HashMap<String, u64>>,
    is_valid_word: Option<&dyn Fn(&str) -> bool>,
    is_noun: Option<&dyn Fn(&str) -> bool>,
) -> Vec<CompoundResult> {
    let input_bytes = input.as_bytes();
    let root_addr = fst.root().addr();
    let mut results: Vec<CompoundResult> = Vec::new();
    let mut seen_compounds: std::collections::HashSet<String> = std::collections::HashSet::new();

    // Short inputs: 2 parts max (3-part is almost always junk)
    // Long inputs (15+ bytes): allow 3 parts for real compounds like "vaskemaskinslange"
    let max_parts = if input_bytes.len() >= 15 { MAX_PARTS } else { 2 };

    // Initialize with root state
    let initial = WalkState {
        fst_addr: root_addr,
        input_pos: 0,
        edits: 0,
        word_bytes: Vec::new(),
        word_start: 0,
        parts: Vec::new(),
        total_edits: 0,
    };

    let mut current_states = vec![initial];
    let mut next_states: Vec<WalkState> = Vec::new();

    // Process states level by level (BFS with beam pruning)
    for _iteration in 0..input_bytes.len() * 4 {
        if current_states.is_empty() { break; }

        for state in &current_states {
            if state.total_edits + state.edits > MAX_TOTAL_EDITS { continue; }
            if state.edits > MAX_EDITS_PER_PART { continue; }

            let node = fst.node(state.fst_addr);

            // FORK: at accepting states (found a complete word part)
            // Only accept parts that are real words (in wordfreq with freq ≥ 10)
            if node.is_final() && state.word_bytes.len() >= MIN_PART_BYTES {
                let matched = String::from_utf8_lossy(&state.word_bytes).to_string();
                // Validate ALL compound parts are real words:
                // Must be in wordfreq (freq ≥ 10) OR recognized by analyzer.
                // Filters junk like "innsjøe", "efisk", "øs", "ikt".
                let in_wf = wordfreq.map_or(true, |wf| wf.contains_key(&matched));
                let in_dict = is_valid_word.map_or(true, |check| check(&matched));
                let is_real_word = in_wf || in_dict;
              if is_real_word {
                let new_part = CompoundPart {
                    matched_word: matched.clone(),
                    input_start: state.word_start,
                    input_end: state.input_pos,
                    edits: state.edits,
                };
                let mut new_parts = state.parts.clone();
                new_parts.push(new_part);
                let new_total = state.total_edits + state.edits;

                // Extend-and-match: at 0-edit accepting states, try extending
                // by 1 DELETE byte to reach a LONGER word, then check if
                // remaining input matches exactly. Bypasses beam pruning
                // for cases like "strøm"→"strømme"+tjeneste.
                // Only at 0-edit states to avoid flooding results.
                if state.edits == 0
                    && new_total + 1 <= MAX_TOTAL_EDITS
                    && state.input_pos + 5 <= input_bytes.len()
                    && new_parts.len() <= max_parts
                {
                    for trans in node.transitions() {
                        let ext_node = fst.node(trans.addr);
                        // 1 delete: check if extension + match reaches final
                        if state.input_pos < input_bytes.len() {
                            let inp = input_bytes[state.input_pos];
                            if let Some(idx) = ext_node.find_input(inp) {
                                let t2 = ext_node.transition(idx);
                                let n2 = fst.node(t2.addr);
                                if n2.is_final() {
                                    // Extended word found (1 edit: delete trans.inp)
                                    let mut ewb = state.word_bytes.clone();
                                    ewb.push(trans.inp);
                                    ewb.push(inp);
                                    let ext_word = String::from_utf8_lossy(&ewb).to_string();
                                    let ext_end = state.input_pos + 1;
                                    let ext_edits = state.edits + 1;
                                    let ext_total = state.total_edits + ext_edits;
                                    // Check if remaining input matches a word from root
                                    if ext_end < input_bytes.len() {
                                        for (w2, end2) in exact_walk_from_root(fst, input_bytes, ext_end) {
                                            if end2 == input_bytes.len()
                                                && is_noun.map_or(true, |check| check(&w2))
                                            {
                                                let mut ext_parts = state.parts.clone();
                                                ext_parts.push(CompoundPart {
                                                    matched_word: ext_word.clone(),
                                                    input_start: state.word_start,
                                                    input_end: ext_end,
                                                    edits: ext_edits,
                                                });
                                                ext_parts.push(CompoundPart {
                                                    matched_word: w2.clone(),
                                                    input_start: ext_end,
                                                    input_end: end2,
                                                    edits: 0,
                                                });
                                                let compound = ext_parts.iter().map(|p| p.matched_word.as_str()).collect::<String>();
                                                if ext_parts.len() <= MAX_PARTS && seen_compounds.insert(compound.clone()) {
                                                    results.push(CompoundResult {
                                                        parts: ext_parts,
                                                        total_edits: ext_total,
                                                        compound_word: compound,
                                                    });
                                                }
                                            }
                                        }
                                    } else if ext_end == input_bytes.len()
                                        && is_noun.map_or(true, |check| check(&ext_word))
                                    {
                                        // Extended word consumes all input
                                        let mut ext_parts = state.parts.clone();
                                        ext_parts.push(CompoundPart {
                                            matched_word: ext_word.clone(),
                                            input_start: state.word_start,
                                            input_end: ext_end,
                                            edits: ext_edits,
                                        });
                                        let compound = ext_parts.iter().map(|p| p.matched_word.as_str()).collect::<String>();
                                        if seen_compounds.insert(compound.clone()) {
                                            results.push(CompoundResult {
                                                parts: ext_parts,
                                                total_edits: ext_total,
                                                compound_word: compound,
                                            });
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                if state.input_pos == input_bytes.len() {
                    // Complete match — all input consumed
                    // Last part must be a noun (not verb, name, adverb etc.)
                    let last_is_noun = is_noun.map_or(true, |check| {
                        check(&new_parts.last().unwrap().matched_word)
                    });
                    // For 3+ part compounds, require ALL parts freq ≥ 50
                    let parts_ok = if new_parts.len() >= 3 {
                        new_parts.iter().all(|p| {
                            wordfreq.map_or(true, |wf|
                                wf.get(&p.matched_word).copied().unwrap_or(0) >= 50)
                        })
                    } else {
                        true
                    };
                    if last_is_noun && parts_ok {
                        let compound = new_parts.iter().map(|p| p.matched_word.as_str()).collect::<String>();
                        if seen_compounds.insert(compound.clone()) {
                            results.push(CompoundResult {
                                parts: new_parts.clone(),
                                total_edits: new_total,
                                compound_word: compound,
                            });
                        }
                    }
                } else if new_parts.len() < max_parts {
                    // Restart from root for next part
                    next_states.push(WalkState {
                        fst_addr: root_addr,
                        input_pos: state.input_pos,
                        edits: 0,
                        word_bytes: Vec::new(),
                        word_start: state.input_pos,
                        parts: new_parts.clone(),
                        total_edits: new_total,
                    });

                    // Try skipping binding letters (s, e) present in input
                    if state.input_pos < input_bytes.len()
                        && BINDING_LETTERS.contains(&input_bytes[state.input_pos])
                    {
                        let binding_char = input_bytes[state.input_pos] as char;
                        let mut parts_with_binding = new_parts.clone();
                        if let Some(last) = parts_with_binding.last_mut() {
                            last.matched_word.push(binding_char);
                            last.input_end = state.input_pos + 1;
                        }
                        next_states.push(WalkState {
                            fst_addr: root_addr,
                            input_pos: state.input_pos + 1,
                            edits: 0,
                            word_bytes: Vec::new(),
                            word_start: state.input_pos + 1,
                            parts: parts_with_binding,
                            total_edits: new_total,
                        });
                    }

                    // Try INSERTING binding -s- when morphology requires it
                    // Only insert when the matched word ends with a suffix
                    // that triggers -s- (e.g., -ing, -het, -skap, -sjon, -tet)
                    if needs_binding_s(&state.word_bytes)
                        && (state.input_pos >= input_bytes.len()
                            || input_bytes[state.input_pos] != b's')
                    {
                        let mut parts_with_s = new_parts;
                        if let Some(last) = parts_with_s.last_mut() {
                            last.matched_word.push('s');
                        }
                        next_states.push(WalkState {
                            fst_addr: root_addr,
                            input_pos: state.input_pos,
                            edits: 0,
                            word_bytes: Vec::new(),
                            word_start: state.input_pos,
                            parts: parts_with_s,
                            total_edits: new_total,
                        });
                    }
                }
              } // is_real_word
            }

            // ADVANCE: Levenshtein moves
            if state.edits >= MAX_EDITS_PER_PART { continue; }

            if state.input_pos < input_bytes.len() {
                let inp_byte = input_bytes[state.input_pos];

                // Check if input is at a 2-byte UTF-8 char (Norwegian å,ø,æ = 0xC3 + XX)
                let inp_is_multibyte = inp_byte == 0xC3 && state.input_pos + 1 < input_bytes.len();
                let inp_char_len = if inp_is_multibyte { 2 } else { 1 };

                // Try all FST transitions
                let is_utf8_cross = |t_inp: u8| -> bool {
                    // True when FST byte and input byte are in different UTF-8 "worlds"
                    // (one is 0xC3 prefix, the other is ASCII) — creates mid-UTF-8 junk
                    (t_inp == 0xC3 && !inp_is_multibyte) || (t_inp != 0xC3 && inp_is_multibyte)
                };

                for trans in node.transitions() {
                    if trans.inp == inp_byte {
                        // Match — advance both, no edit
                        let mut wb = state.word_bytes.clone();
                        wb.push(trans.inp);
                        next_states.push(WalkState {
                            fst_addr: trans.addr,
                            input_pos: state.input_pos + 1,
                            edits: state.edits,
                            word_bytes: wb,
                            word_start: state.word_start,
                            parts: state.parts.clone(),
                            total_edits: state.total_edits,
                        });
                    } else if !is_utf8_cross(trans.inp) {
                        // Substitute — advance both, +1 edit
                        // (only when both are same byte-width to avoid mid-UTF-8 junk)
                        let mut wb = state.word_bytes.clone();
                        wb.push(trans.inp);
                        next_states.push(WalkState {
                            fst_addr: trans.addr,
                            input_pos: state.input_pos + 1,
                            edits: state.edits + 1,
                            word_bytes: wb,
                            word_start: state.word_start,
                            parts: state.parts.clone(),
                            total_edits: state.total_edits,
                        });
                    }

                    // UTF-8 aware substitution: 2-byte input char (å,ø,æ) ↔ 1-byte FST char
                    // Free only in the COMMON dyslexic direction: ASCII→Norwegian (e→ø, o→å, a→æ)
                    // Reverse direction (ø→e, å→o, æ→a) costs 1 edit
                    if inp_is_multibyte && trans.inp != 0xC3 {
                        // Input has 2-byte char (ø/å/æ), FST has 1-byte char (e/o/a)
                        // This is the REVERSE direction — always costs 1 edit
                        let cost = 1;
                        let mut wb = state.word_bytes.clone();
                        wb.push(trans.inp);
                        next_states.push(WalkState {
                            fst_addr: trans.addr,
                            input_pos: state.input_pos + 2,
                            edits: state.edits + cost,
                            word_bytes: wb,
                            word_start: state.word_start,
                            parts: state.parts.clone(),
                            total_edits: state.total_edits,
                        });
                    }
                    if !inp_is_multibyte && trans.inp == 0xC3 {
                        // Input has 1-byte char, FST starts 2-byte char
                        // Substitute: consume 1 input byte, emit complete 2-byte FST char
                        // Free for vowel pairs, 1 edit otherwise
                        let next_node = fst.node(trans.addr);
                        for trans2 in next_node.transitions() {
                            let cost = if is_free_vowel_swap(inp_byte, trans2.inp) { 0 } else { 1 };
                            let mut wb = state.word_bytes.clone();
                            wb.push(0xC3);
                            wb.push(trans2.inp);
                            next_states.push(WalkState {
                                fst_addr: trans2.addr,
                                input_pos: state.input_pos + 1,
                                edits: state.edits + cost,
                                word_bytes: wb,
                                word_start: state.word_start,
                                parts: state.parts.clone(),
                                total_edits: state.total_edits,
                            });
                        }
                        // Delete: emit complete 2-byte FST char, don't consume input → 1 edit
                        if state.edits + 1 <= MAX_EDITS_PER_PART {
                            let next_node = fst.node(trans.addr);
                            for trans2 in next_node.transitions() {
                                let mut wb = state.word_bytes.clone();
                                wb.push(0xC3);
                                wb.push(trans2.inp);
                                next_states.push(WalkState {
                                    fst_addr: trans2.addr,
                                    input_pos: state.input_pos,
                                    edits: state.edits + 1,
                                    word_bytes: wb,
                                    word_start: state.word_start,
                                    parts: state.parts.clone(),
                                    total_edits: state.total_edits,
                                });
                            }
                        }
                    } else if trans.inp != 0xC3 || inp_is_multibyte {
                        // Delete (FST byte not in input) — advance FST only, +1 edit
                        // Skip when trans is 0xC3 and input is ASCII (handled above atomically)
                        if state.edits + 1 <= MAX_EDITS_PER_PART {
                            let mut wb = state.word_bytes.clone();
                            wb.push(trans.inp);
                            next_states.push(WalkState {
                                fst_addr: trans.addr,
                                input_pos: state.input_pos,
                                edits: state.edits + 1,
                                word_bytes: wb,
                                word_start: state.word_start,
                                parts: state.parts.clone(),
                                total_edits: state.total_edits,
                            });
                        }
                    }
                }

                // Insert (input byte not in FST) — advance input only, +1 edit
                // For multibyte input chars, skip entire 2-byte sequence
                if state.edits + 1 <= MAX_EDITS_PER_PART {
                    let skip = if inp_is_multibyte { 2 } else { 1 };
                    next_states.push(WalkState {
                        fst_addr: state.fst_addr,
                        input_pos: state.input_pos + skip,
                        edits: state.edits + 1,
                        word_bytes: state.word_bytes.clone(),
                        word_start: state.word_start,
                        parts: state.parts.clone(),
                        total_edits: state.total_edits,
                    });
                }

                // Transposition (Damerau): swap adjacent input bytes [a,b] → FST [b,a]
                // Counts as 1 edit instead of 2 substitutions
                if !inp_is_multibyte
                    && state.input_pos + 1 < input_bytes.len()
                    && state.edits + 1 <= MAX_EDITS_PER_PART
                {
                    let a = inp_byte;
                    let b = input_bytes[state.input_pos + 1];
                    if a != b && b != 0xC3 {
                        // Try FST path [b, a] (swapped order)
                        if let Some(idx1) = node.find_input(b) {
                            let t1 = node.transition(idx1);
                            let mid = fst.node(t1.addr);
                            if let Some(idx2) = mid.find_input(a) {
                                let t2 = mid.transition(idx2);
                                let mut wb = state.word_bytes.clone();
                                wb.push(b);
                                wb.push(a);
                                next_states.push(WalkState {
                                    fst_addr: t2.addr,
                                    input_pos: state.input_pos + 2,
                                    edits: state.edits + 1,
                                    word_bytes: wb,
                                    word_start: state.word_start,
                                    parts: state.parts.clone(),
                                    total_edits: state.total_edits,
                                });
                            }
                        }
                    }
                }
            } else {
                // Input exhausted — try deleting remaining FST bytes
                if state.edits + 1 <= MAX_EDITS_PER_PART {
                    for trans in node.transitions() {
                        if trans.inp == 0xC3 {
                            // Delete entire 2-byte char atomically — 1 edit
                            let next_node = fst.node(trans.addr);
                            for trans2 in next_node.transitions() {
                                let mut wb = state.word_bytes.clone();
                                wb.push(0xC3);
                                wb.push(trans2.inp);
                                next_states.push(WalkState {
                                    fst_addr: trans2.addr,
                                    input_pos: state.input_pos,
                                    edits: state.edits + 1,
                                    word_bytes: wb,
                                    word_start: state.word_start,
                                    parts: state.parts.clone(),
                                    total_edits: state.total_edits,
                                });
                            }
                        } else {
                            let mut wb = state.word_bytes.clone();
                            wb.push(trans.inp);
                            next_states.push(WalkState {
                                fst_addr: trans.addr,
                                input_pos: state.input_pos,
                                edits: state.edits + 1,
                                word_bytes: wb,
                                word_start: state.word_start,
                                parts: state.parts.clone(),
                                total_edits: state.total_edits,
                            });
                        }
                    }
                }
            }
        }

        // Beam pruning: keep best states
        next_states.sort_by_key(|s| s.priority());
        next_states.truncate(BEAM_WIDTH);

        std::mem::swap(&mut current_states, &mut next_states);
        next_states.clear();

        if results.len() >= 200 { break; } // enough results
    }

    // Sort by (total_edits, part_count, -min_frequency)
    // At equal edits and parts, prefer decompositions where all parts are common
    results.sort_by(|a, b| {
        let ord = a.total_edits.cmp(&b.total_edits);
        if ord != std::cmp::Ordering::Equal { return ord; }
        let ord = a.parts.len().cmp(&b.parts.len());
        if ord != std::cmp::Ordering::Equal { return ord; }
        // Higher minimum frequency = better (more likely real compound)
        let freq = |r: &CompoundResult| -> u64 {
            r.parts.iter()
                .map(|p| wordfreq.and_then(|wf| wf.get(&p.matched_word).copied()).unwrap_or(0))
                .min().unwrap_or(0)
        };
        freq(b).cmp(&freq(a))
    });
    results
}

/// Load the raw FST from an .mfst file.
pub fn load_fst_from_mfst(path: &str) -> Result<Fst<Vec<u8>>, Box<dyn std::error::Error>> {
    let data = std::fs::read(path)?;
    if data.len() < 28 || &data[0..4] != b"MFST" {
        return Err("Not a valid MFST file".into());
    }
    // Header: [magic 4B] [version 4B] [fst_offset 8B] [fst_len 8B] ...
    let fst_offset = u64::from_le_bytes(data[8..16].try_into()?) as usize;
    let fst_len = u64::from_le_bytes(data[16..24].try_into()?) as usize;
    if fst_offset + fst_len > data.len() {
        return Err("FST data out of bounds".into());
    }
    let fst_bytes = data[fst_offset..fst_offset + fst_len].to_vec();
    Ok(Fst::new(fst_bytes)?)
}
