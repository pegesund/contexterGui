#!/bin/bash
# Automated error detection test
# Requires: Word open with NorskTale add-in connected, acatts-rust running
#
# Usage: ./test_errors.sh

ENDPOINT="https://127.0.0.1:3000/errors"
PASS=0
FAIL=0

check_error() {
    local desc="$1"
    local word="$2"
    local expected_suggestion="$3"
    local json="$4"

    if echo "$json" | python3 -c "
import json, sys
errors = json.load(sys.stdin)
found = [e for e in errors if e['word'] == '$word']
if found and (not '$expected_suggestion' or found[0]['suggestion'] == '$expected_suggestion'):
    sys.exit(0)
sys.exit(1)
" 2>/dev/null; then
        echo "  PASS: $desc ('$word' → '$expected_suggestion')"
        PASS=$((PASS + 1))
    else
        echo "  FAIL: $desc ('$word' expected '$expected_suggestion')"
        FAIL=$((FAIL + 1))
    fi
}

check_no_error() {
    local desc="$1"
    local word="$2"
    local json="$3"

    if echo "$json" | python3 -c "
import json, sys
errors = json.load(sys.stdin)
found = [e for e in errors if e['word'] == '$word']
if not found:
    sys.exit(0)
sys.exit(1)
" 2>/dev/null; then
        echo "  PASS: $desc ('$word' not flagged)"
        PASS=$((PASS + 1))
    else
        echo "  FAIL: $desc ('$word' should not be flagged)"
        FAIL=$((FAIL + 1))
    fi
}

echo "=== NorskTale Error Detection Test ==="
echo ""

# Test 1: Insert text with spelling error
echo "Test 1: Spelling error 'somx' → 'som'"
osascript -e 'tell application "Microsoft Word" to insert text "
Fotball er en morsom sport somx er veldig morsom." at end of text object of active document' 2>/dev/null
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "spelling: somx" "somx" "som" "$ERRORS"

# Clean up
osascript -e 'tell application "Microsoft Word"
    set t to content of text object of active document
    set r to create range active document start ((length of t) - 52) end (length of t)
    select r
    type text selection text ""
end tell' 2>/dev/null
sleep 2

# Test 2: Insert text with grammar error
echo ""
echo "Test 2: Grammar error"
osascript -e 'tell application "Microsoft Word" to insert text "
Jeg kan gikk til butikken." at end of text object of active document' 2>/dev/null
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "grammar: gikk after modal" "gikk" "" "$ERRORS"

# Clean up
osascript -e 'tell application "Microsoft Word"
    set t to content of text object of active document
    set r to create range active document start ((length of t) - 30) end (length of t)
    select r
    type text selection text ""
end tell' 2>/dev/null
sleep 2

# Test 3: Correct text should have no new errors
echo ""
echo "Test 3: Correct text — no errors"
osascript -e 'tell application "Microsoft Word" to insert text "
Fotball er en morsom sport." at end of text object of active document' 2>/dev/null
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "no error: sport" "sport" "$ERRORS"

# Clean up
osascript -e 'tell application "Microsoft Word"
    set t to content of text object of active document
    set r to create range active document start ((length of t) - 30) end (length of t)
    select r
    type text selection text ""
end tell' 2>/dev/null

echo ""
echo "=== Results: $PASS passed, $FAIL failed ==="
