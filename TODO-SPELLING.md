# TODO — Spelling pipeline

Deferred items from the user-feedback round on v0.1.41. These need deeper
changes than the simple punktfikser we did in v0.1.42 and we want to do them
as one focused pass with good regression coverage.

## 1. Phonetic-equivalent ortho distance

**Symptom.** "blabar" → suggests "blader" (leaves) instead of "blåbær".

**Why current pipeline misses it.** The ortho ranker compares trigrams as
exact strings. `a`, `å` and `æ` are completely different bytes, so:
- "blabar" trigrams: bla, lab, aba, bar
- "blåbær" trigrams: blå, låb, åbæ, bær  (0 overlap)
- "blader" trigrams: bla, lad, ade, der  (1 overlap)

"blader" looks closer purely on string distance. wordfreq boost doesn't
bridge the gap because "blader" is also reasonably common.

**Fix shape.** Add a phonetic-normalised distance: fold `å → a`, `æ → a`,
`ø → e` (or `o`) before computing trigrams / prefix / edit distance, then
use the higher of phonetic-sim and raw-sim as the ortho input. Probably
also worth offering this as the candidate-generation gate for Source 12 so
phonetic chains can fire without needing a dict-valid intermediate.

**Risk.** Touches the core ortho computation. Every existing test goes
through this. Need to verify dyslexia_tests, test_spelling, test_spelling_nn
all keep their pass count.

## 2. Letter transposition source

**Symptom.** "brode" → suggests "brodne" (valid word, contextually
plausible) instead of "bordet" (what the user meant).

**Why current pipeline misses it.** "brode" → "bordet" is Damerau
distance 2 (one transposition r↔o + one insertion t) but plain
Levenshtein distance 3 (r→o sub + o→r sub + insert t). Our active fuzzy
backend, `mtag-rs::Analyzer::fuzzy_lookup`, uses `fst::automaton::
Levenshtein` — plain Levenshtein, no transposition handling — and the
default cap is dist 2, so "bordet" is never produced as a candidate.

**Code already in the repo (but unwired).**
`rustSpell/spellcheck/src/lib.rs:45-50` implements Damerau-Levenshtein
distance correctly. It is only used by the standalone `cg3-spell-*`
bins and the `dyslex-ios-ffi` crate. Nothing in acatts-rust or
nostos-cognio pulls it in. The README of that crate explicitly
advertises "Edit distance (Levenshtein with Damerau transpositions)"
— we lost the wiring at some point and never noticed.

**Fix shape, ranked.**
1. Cheapest: add an explicit transposition candidate source to
   `spelling_scorer.rs` — for each adjacent pair, swap them and check
   the dictionary. Mirror Source 13's pattern with a vowel-insert-style
   boost. Combine with the existing fuzzy/prefix sources to catch
   "brode" → "borde" → "bordet" via the chain.
2. More invasive: replace `fst::Levenshtein` in `mtag-rs::Analyzer`
   with a Damerau-Levenshtein FST automaton (or fall back to the
   length-bucketed Damerau in the spellcheck crate). Brings the
   benefit to every caller automatically but is a bigger change.

**Risk.** Generates many extra candidates for typical typos. The dict
filter cap might saturate and push out fuzzy hits. May need to bump
cap or give transposition candidates the same skeleton-style boost
the vowel-insertion source uses.

## 3. Verify earlier feedback claims are actually fixed in v0.1.42

The user's feedback round was tested against an older build. Cases
reported as "no response" or "wrong word" that the console pipeline now
handles correctly:

- piza → pizza (was reported as pipa)
- Bergeen → Bergen (was reported as no response)
- stajonen → stasjonen singular (was reported as plural stasjonene)
- kjore → kjøre (was reported as no response)
- pa → på (was reported as no response)
- blåøbær → blåbær (was reported as no response — 1-edit deletion, easy)
- veldiøg → veldig (was reported as no response — 1-edit deletion, easy)
- skæål → skal (was reported as no response — 1-edit deletion, easy)
- Bærna → Barna, møttæø → møtte, sykæøhuset → sykehuset etc.

Tester should re-run the full feedback document on v0.1.42 before we
spend effort on the harder items above. Likely most of the "missed"
group-4 cases (Norwegian-special-char insertions) already work.

## 4. Things we explicitly decided NOT to fix

- "hjelpa" — valid Bokmål definite-feminine form, not a typo
- "spisse piza" — "spisse" is a real word ("to sharpen"). Detecting
  "verb form doesn't match the context" is a grammar problem, not a
  spelling problem. Out of scope for this pipeline.

## State at time of writing

v0.1.42 ships these completed in this round:
- wordfreq-aware ortho boost (compute_boost)
- fuzzy-check capitalized typos (Bergeen-class)
- Source 13 vowel-insertion for short consonant skeletons (lgn/skgn/bnkn)

Test_spelling: 32/33 (only blabar fails).
dyslexia_tests: 811/834 (baseline, unchanged).
test_punctuation: 171/239 (baseline, unchanged).
