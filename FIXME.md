# FIXME — Completion pipeline broken

## Current state (DO NOT COMMIT)
Working version: git commit e78d19c
Last good commit on main: 0a75806

## What's broken

### 1. Left column suggests tokens, not always proper words
- `complete_word()` works in test binary (test-cw) — produces "sport", "spill", etc.
- In GUI, context extraction was wrong (masked sentence strips word parts)
- Fixed by using `self.context.sentence` directly instead of masked sentence
- BUT scoring is wrong: "spor" scores same as "sport" — baselines/PMI not applied?

### 2. Right column is BPE tokens, not words
- `build_right_completions()` produces raw BPE tokens
- Working version ran right column through `grammar_filter()` too
- Need to either run right through `complete_word()` or `grammar_filter()`

### 3. Grammar filter not filtering left column
- `check_sentence_sync` added to grammar actor
- Actor handles `ActorMessage::Sync` and calls `checker.check_sentence()`
- But grammar filter results still pass illegal words through
- Need to verify: is `check_sentence_sync` actually being called?
- Previous log showed `grammar_sync` working for first call but not second

### 4. OCR popup triggers on any clipboard image change
- Including screenshots taken during testing
- Needs to not trigger when our app is in background

## Working version pipeline (e78d19c)
1. `complete_word()` called on main thread with model, prefix_index, baselines, wordfreq, embedding_store, fallback_dict, fallback_prefix
2. Results dictionary-filtered: `checker.has_word()`
3. Results grammar-filtered: `grammar_filter()` with `checker.check_sentence()` as check_fn
4. Right column also grammar-filtered
5. Quality setting controls max_steps (0/1/3)

## What was changed to break it
- Model moved to BERT worker thread (can't call complete_word from main thread)
- Grammar checker moved to actor thread (can't call check_sentence from main thread)
- `grammar_filter` commented out with "// skipped"
- `complete_word` disabled with `if false`
- Replaced with `build_bpe_completions` which is simpler and produces tokens

## Fix strategy
- `complete_word()` now works in BERT worker (CompleteWord request) ✓
- Context extraction fixed (use sentence directly) ✓
- Grammar actor has `check_sentence_sync` ✓
- Still need: verify grammar filter actually runs, fix right column, fix scoring
