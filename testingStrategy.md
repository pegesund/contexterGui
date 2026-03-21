# Testing Strategy

## Overview

Automated testing of NorskTale uses three layers:
1. **Unit test binaries** — test individual components without GUI
2. **Integration test via /errors endpoint** — test full pipeline with Word
3. **Manual verification** — for UI and interaction testing

## 1. Unit Test Binaries

Run from `contexterGui/` directory.

### test-cw — Word completion
Tests `complete_word()` function directly with various contexts and prefixes.
```bash
cargo run --bin test-cw
```
Verifies: correct word suggestions for given context (e.g., "Fotball er en morsom" + "s" → "sport")

### test-prolog — Grammar check timing
Tests SWI-Prolog grammar checking speed and correctness.
```bash
cargo run --bin test-prolog
```
Verifies: grammar errors detected correctly, timing under 1ms per sentence.

### test-spelling — Unknown word detection
Tests spelling/unknown word detection pipeline.
```bash
cargo run --bin test-spelling
```
Verifies: misspelled words flagged, correct words not flagged, compound splitting rules.

### test-fotball — Full completion pipeline
Tests BPE completion pipeline end-to-end.
```bash
cargo run --bin test-fotball
```

## 2. Integration Testing via /errors Endpoint

### Prerequisites
- Microsoft Word open with a document
- NorskTale Word Add-in loaded and connected
- `acatts-rust` running (`cargo run --bin acatts-rust`)

### The /errors Endpoint

The Word Add-in HTTPS server exposes a test endpoint:

```bash
curl -k https://127.0.0.1:3000/errors | python3 -m json.tool
```

Returns JSON array of all current writing errors:
```json
[
  {
    "category": "spelling",
    "word": "somx",
    "suggestion": "som",
    "rule": "stavefeil_bert",
    "sentence": "sport somx er veldig morsom."
  }
]
```

Categories: `spelling`, `grammar`, `sentence_boundary`

### Automated Test Script

```bash
./test_errors.sh
```

This script:
1. Inserts text with known errors into Word via AppleScript
2. Waits for NorskTale to process (5 seconds)
3. Queries `/errors` endpoint
4. Verifies expected errors are found with correct suggestions
5. Cleans up inserted text

### AppleScript Capabilities

The test script uses these Word AppleScript commands:

**Insert text:**
```applescript
tell application "Microsoft Word"
    insert text "Text here." at end of text object of active document
end tell
```

**Delete text (last N chars):**
```applescript
tell application "Microsoft Word"
    set t to content of text object of active document
    set r to create range active document start ((length of t) - N) end (length of t)
    select r
    type text selection text ""
end tell
```

**Type at cursor:**
```applescript
tell application "Microsoft Word"
    type text selection text "typed text"
end tell
```

### Writing New Tests

Add test cases to `test_errors.sh` using:

```bash
# Insert text with error
osascript -e 'tell application "Microsoft Word" to insert text "
Error text here." at end of text object of active document'
sleep 5

# Check error detected
ERRORS=$(curl -sk https://127.0.0.1:3000/errors)
check_error "description" "error_word" "expected_suggestion" "$ERRORS"

# Or check no false positive
check_no_error "description" "correct_word" "$ERRORS"

# Clean up
osascript -e 'tell application "Microsoft Word"
    set t to content of text object of active document
    set r to create range active document start ((length of t) - N) end (length of t)
    select r
    type text selection text ""
end tell'
```

## 3. Performance Rules (MUST NOT violate)

These rules exist because violating them caused severe regressions. Every change must be verified against them.

### Rule 1: Only check what changed
- NEVER rescan the whole document on every keystroke
- When a paragraph changes, only recheck THAT paragraph's sentences
- Use `processed_sentence_hashes` to skip unchanged sentences
- **Test:** Open a 50+ paragraph document, type one character. The log should show only 1 paragraph being processed, not 50.

### Rule 2: Only check complete units
- Spelling: only after the user presses space (word is complete)
- Grammar: only after sentence ends with punctuation (. ! ? :)
- NEVER flag errors on the word currently being typed
- **Test:** Type "fotba" slowly. No spelling error should appear until after pressing space.

### Rule 3: Don't block the UI
- Grammar checks go through the actor (async)
- BERT scoring goes through the worker (async)
- Spelling suggestions must not freeze the GUI
- Grammar filter for completions uses batch (one round-trip)
- **Test:** While typing, the GUI should never freeze for >100ms

### Rule 4: Don't drop features for technical convenience
- If a refactor makes a feature harder to call, find a way to still call it
- NEVER comment out working code with "// skipped" or "// TODO"
- If the grammar filter was working, it must KEEP working after any change
- **Test:** Verify both columns show grammar-filtered words, not BPE tokens

### How to verify these rules

Add `/timing` endpoint (future) that reports:
- Time since last keystroke
- Number of paragraphs processed in last cycle
- Number of grammar checks sent
- Number of spelling checks sent

For now, verify manually via log:
```bash
# After typing one char, check how many paragraphs were processed:
grep "Addin changed paragraph" /path/to/acatts-rust.log | tail -5
# Should show only 1 paragraph, not the whole document
```

## 4. When to Test

- **After changing completion pipeline** — run test-cw
- **After changing grammar rules** — run test-prolog
- **After changing spelling detection** — run test-spelling
- **After any change to main.rs** — run test_errors.sh
- **Before merging to main** — run ALL tests
- **After ANY change** — verify performance rules (section 3) are not violated
