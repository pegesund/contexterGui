# FIXME — Remaining bugs on fix-completion-pipeline branch

## Test endpoint: GET /errors

The Word Add-in HTTPS server exposes a `/errors` endpoint for automated testing.

```bash
curl -k https://127.0.0.1:3000/errors | python3 -m json.tool
```

Returns JSON array of current writing errors:
```json
[
  {
    "category": "spelling",      // "spelling", "grammar", or "sentence_boundary"
    "word": "somx",              // the error word
    "suggestion": "som",         // suggested correction
    "rule": "stavefeil_bert",    // rule name
    "sentence": "sport somx er"  // sentence context
  }
]
```

Use this to build automated tests:
1. Use AppleScript to type text with known errors into Word
2. Wait for processing (2-3 seconds)
3. `curl -k https://127.0.0.1:3000/errors` to get detected errors
4. Verify expected errors are found with correct suggestions

## What works
- complete_word() produces proper words (sport, sportsgren, etc.)
- Grammar filter works and is fast (FST v3: 0.1ms instead of 315ms)
- Both left and right columns show real words
- Grammar filter batch via actor channel works
- UTF-8 compound splitter crash fixed

## What's broken

### 1. Spelling doesn't find all errors in a sentence
- "Fotball er en morsom sport somx er veldig morsson." only flags "somx", not "morsson"
- Grammar checker's check_sentence_full returns unknown_words but seems to miss some
- Need to investigate: does SWI-Prolog return all unknown words?

### 2. Paragraph change detection misses some paragraphs
- "Dettex er en test" in a separate paragraph not detected by add-in
- The add-in's drain_changed_paragraphs might not report all paragraphs on doc load

### 3. Spelling/grammar errors triggered mid-word
- Errors appear while user is still typing (before space)
- Previous fix attempts broke ALL error detection
- Need to fix at the RIGHT level — the add-in should only report
  completed words, not the one currently being typed
- The proper fix might be in the Word Add-in JavaScript, not Rust

### 4. OCR popup triggers on clipboard screenshots
- Minor annoyance during testing

## Branch status
- Branch: fix-completion-pipeline
- Main: v1.0-mac tag + toolbar/UI fixes
- Do NOT merge to main until spelling detection is reliable
