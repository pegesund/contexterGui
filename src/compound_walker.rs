//! Compound word FST walker.
//!
//! Walks the FST node-by-node to find compound word decompositions
//! with fuzzy matching per part. Handles Norwegian compound words
//! like "kjøkkenbord" = "kjøkken" + "bord" even when misspelled.

use fst::raw::{CompiledAddr, Fst, Node};
use std::collections::BinaryHeap;
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
const BEAM_WIDTH: usize = 2000;
const BINDING_LETTERS: &[u8] = b"se";

/// Walk the FST to find compound word decompositions for the input.
/// Returns decompositions sorted by total edit distance (best first).
pub fn compound_fuzzy_walk<D: AsRef<[u8]>>(
    fst: &Fst<D>,
    input: &str,
) -> Vec<CompoundResult> {
    let input_bytes = input.as_bytes();
    let root_addr = fst.root().addr();
    let mut results: Vec<CompoundResult> = Vec::new();
    let mut seen_compounds: std::collections::HashSet<String> = std::collections::HashSet::new();

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
            if node.is_final() && state.word_bytes.len() >= MIN_PART_BYTES {
                let matched = String::from_utf8_lossy(&state.word_bytes).to_string();
                let new_part = CompoundPart {
                    matched_word: matched.clone(),
                    input_start: state.word_start,
                    input_end: state.input_pos,
                    edits: state.edits,
                };
                let mut new_parts = state.parts.clone();
                new_parts.push(new_part);
                let new_total = state.total_edits + state.edits;

                if state.input_pos == input_bytes.len() {
                    // Complete match — all input consumed
                    let compound = new_parts.iter().map(|p| p.matched_word.as_str()).collect::<String>();
                    if seen_compounds.insert(compound.clone()) {
                        results.push(CompoundResult {
                            parts: new_parts.clone(),
                            total_edits: new_total,
                            compound_word: compound,
                        });
                    }
                } else if new_parts.len() < MAX_PARTS {
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

                    // Try binding letters (s, e) between parts
                    if state.input_pos < input_bytes.len()
                        && BINDING_LETTERS.contains(&input_bytes[state.input_pos])
                    {
                        next_states.push(WalkState {
                            fst_addr: root_addr,
                            input_pos: state.input_pos + 1,
                            edits: 0,
                            word_bytes: Vec::new(),
                            word_start: state.input_pos + 1,
                            parts: new_parts,
                            total_edits: new_total,
                        });
                    }
                }
            }

            // ADVANCE: Levenshtein moves
            if state.edits >= MAX_EDITS_PER_PART { continue; }

            if state.input_pos < input_bytes.len() {
                let inp_byte = input_bytes[state.input_pos];

                // Check if input is at a 2-byte UTF-8 char (Norwegian å,ø,æ = 0xC3 + XX)
                let inp_is_multibyte = inp_byte == 0xC3 && state.input_pos + 1 < input_bytes.len();
                let inp_char_len = if inp_is_multibyte { 2 } else { 1 };

                // Try all FST transitions
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
                    } else {
                        // Substitute — advance both, +1 edit
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
                    // Count as 1 edit instead of 2
                    if inp_is_multibyte && trans.inp != 0xC3 {
                        // Input has 2-byte char, FST has 1-byte char
                        // Skip both input bytes, take FST byte → 1 edit
                        let mut wb = state.word_bytes.clone();
                        wb.push(trans.inp);
                        next_states.push(WalkState {
                            fst_addr: trans.addr,
                            input_pos: state.input_pos + 2, // skip both bytes of å/ø/æ
                            edits: state.edits + 1,
                            word_bytes: wb,
                            word_start: state.word_start,
                            parts: state.parts.clone(),
                            total_edits: state.total_edits,
                        });
                    }
                    if !inp_is_multibyte && trans.inp == 0xC3 {
                        // Input has 1-byte char, FST starts 2-byte char
                        // Follow the 0xC3 transition, then try all continuations consuming 1 input byte
                        let next_node = fst.node(trans.addr);
                        for trans2 in next_node.transitions() {
                            // This completes the 2-byte FST char; consume 1 input byte → 1 edit
                            let mut wb = state.word_bytes.clone();
                            wb.push(0xC3);
                            wb.push(trans2.inp);
                            next_states.push(WalkState {
                                fst_addr: trans2.addr,
                                input_pos: state.input_pos + 1, // consume 1 input byte
                                edits: state.edits + 1,
                                word_bytes: wb,
                                word_start: state.word_start,
                                parts: state.parts.clone(),
                                total_edits: state.total_edits,
                            });
                        }
                    }

                    // Delete (FST byte not in input) — advance FST only, +1 edit
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

                // Insert (input byte not in FST) — advance input only, +1 edit
                if state.edits + 1 <= MAX_EDITS_PER_PART {
                    next_states.push(WalkState {
                        fst_addr: state.fst_addr,
                        input_pos: state.input_pos + 1,
                        edits: state.edits + 1,
                        word_bytes: state.word_bytes.clone(),
                        word_start: state.word_start,
                        parts: state.parts.clone(),
                        total_edits: state.total_edits,
                    });
                }
            } else {
                // Input exhausted — try deleting remaining FST bytes
                if state.edits + 1 <= MAX_EDITS_PER_PART {
                    for trans in node.transitions() {
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

        // Beam pruning: keep best states
        next_states.sort_by_key(|s| s.priority());
        next_states.truncate(BEAM_WIDTH);

        std::mem::swap(&mut current_states, &mut next_states);
        next_states.clear();

        if results.len() >= 50 { break; } // enough results
    }

    results.sort_by_key(|r| r.total_edits);
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
