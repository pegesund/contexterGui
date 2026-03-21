#!/bin/bash
# Automated error detection test — simulates real user typing
# Requires: Word open with NorskTale add-in connected, acatts-rust running
#
# Usage: ./test_errors.sh

ENDPOINT="https://127.0.0.1:3000/errors"
PASS=0
FAIL=0
DELAY=0.2  # delay between keystrokes (seconds)

check_error() {
    local desc="$1"
    local word="$2"
    local expected_suggestion="$3"
    local json="$4"

    if echo "$json" | python3 -c "
import json, sys
errors = json.load(sys.stdin)
# Match by word OR by word appearing in the error word (grammar errors use full sentence)
found = [e for e in errors if e['word'] == '$word' or '$word' in e['word'] or '$word' in e.get('sentence','')]
if found and (not '$expected_suggestion' or any(f['suggestion'] == '$expected_suggestion' or '$expected_suggestion' in f.get('suggestion','') for f in found)):
    sys.exit(0)
sys.exit(1)
" 2>/dev/null; then
        echo "  PASS: $desc ('$word' → '$expected_suggestion')"
        PASS=$((PASS + 1))
    else
        echo "  FAIL: $desc ('$word' expected '$expected_suggestion')"
        echo "        errors: $(echo "$json" | python3 -c "import json,sys; [print('    ',e['word'][:30],'|',e['rule']) for e in json.load(sys.stdin)]" 2>/dev/null)"
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

# Type text character by character, like a real user
type_text() {
    local text="$1"
    for (( i=0; i<${#text}; i++ )); do
        local char="${text:$i:1}"
        if [ "$char" = $'\n' ]; then
            osascript -e 'tell application "Microsoft Word" to type text selection text return' 2>/dev/null
        else
            osascript -e "tell application \"Microsoft Word\" to type text selection text \"$char\"" 2>/dev/null
        fi
        sleep $DELAY
    done
}

# Move cursor to end of document
move_to_end() {
    osascript << 'APPLEOF' 2>/dev/null
tell application "Microsoft Word"
    set sel to selection of active document
    end key move sel move end of story
end tell
APPLEOF
}

# Select and delete last N characters
delete_last() {
    local n="$1"
    osascript -e "tell application \"Microsoft Word\"
        set t to content of text object of active document
        set r to create range active document start ((length of t) - $n) end (length of t)
        select r
        type text selection text \"\"
    end tell" 2>/dev/null
}

# Type backspace N times
backspace() {
    local n="$1"
    for (( i=0; i<n; i++ )); do
        osascript -e 'tell application "System Events" to key code 51' 2>/dev/null
        sleep 0.1
    done
}

echo "=== NorskTale Error Detection Test ==="
echo "  Simulating real user typing"
echo ""

# Activate Word
osascript -e 'tell application "Microsoft Word" to activate' 2>/dev/null
sleep 1

# Move cursor to end
move_to_end
sleep 0.5

# --- Test 1: Spelling error ---
echo "Test 1: Spelling error 'somx' → 'som'"
type_text $'\n'"Fotball er en morsom sport somx er veldig morsom."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "spelling: somx" "somx" "som" "$ERRORS"

# Clean up
delete_last 52
sleep 2

# --- Test 2: Correct text ---
echo ""
echo "Test 2: Correct text — no false positives"
type_text $'\n'"Fotball er en morsom sport."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "no error: sport" "sport" "$ERRORS"
check_no_error "no error: morsom" "morsom" "$ERRORS"

# Clean up
delete_last 30
sleep 2

# --- Test 3: Multiple errors ---
echo ""
echo "Test 3: Multiple errors in one sentence"
type_text $'\n'"Jeg liker å spise matx og drikkx."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "spelling: matx" "matx" "" "$ERRORS"
check_error "spelling: drikkx" "drikkx" "" "$ERRORS"

# Clean up
delete_last 36
sleep 2

# --- Test 4: Type and correct ---
echo ""
echo "Test 4: Type misspelled word, then fix with backspace"
type_text $'\n'"Jeg liker fotbalx"
sleep 3
ERRORS_BEFORE=$(curl -sk "$ENDPOINT")
# Now backspace and fix
backspace 1
type_text "l. "
sleep 5
ERRORS_AFTER=$(curl -sk "$ENDPOINT")
check_no_error "fixed: fotball not flagged" "fotball" "$ERRORS_AFTER"

# Clean up
delete_last 22
sleep 2

# --- Test 5: Grammar error (adj gender mismatch) ---
echo ""
echo "Test 5: Grammar error — 'morsom spor' gender mismatch"
type_text $'\n'"Fotball er en morsom spor."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
# Grammar errors may use full sentence as word — check that the sentence with "spor" has a grammar error
if echo "$ERRORS" | python3 -c "
import json, sys
errors = json.load(sys.stdin)
found = [e for e in errors if e['category'] == 'grammar' and 'spor' in e.get('sentence','')]
sys.exit(0 if found else 1)
" 2>/dev/null; then
    echo "  PASS: grammar error detected for 'morsom spor'"
    PASS=$((PASS + 1))
else
    echo "  FAIL: no grammar error for 'morsom spor'"
    echo "        errors: $(echo "$ERRORS" | python3 -c "import json,sys; [print('    ',e['category'],e['word'][:30],'|',e['rule']) for e in json.load(sys.stdin)]" 2>/dev/null)"
    FAIL=$((FAIL + 1))
fi

# Clean up
delete_last 29
sleep 1

echo ""
echo "=== Results: $PASS passed, $FAIL failed ==="
