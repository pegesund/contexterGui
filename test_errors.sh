#!/bin/bash
# Automated error detection test — simulates real user typing and editing
# Requires: Word open with NorskTale add-in connected, acatts-rust running
#
# Usage: ./test_errors.sh

ENDPOINT="https://127.0.0.1:3000/errors"
PASS=0
FAIL=0
DELAY=0.15  # delay between keystrokes

# --- Helper functions ---

check_error() {
    local desc="$1"
    local word="$2"
    local expected_suggestion="$3"
    local json="$4"

    if echo "$json" | python3 -c "
import json, sys
errors = json.load(sys.stdin)
found = [e for e in errors if e['word'] == '$word' or '$word' in e['word'] or '$word' in e.get('sentence','')]
if found and (not '$expected_suggestion' or any(f['suggestion'] == '$expected_suggestion' or '$expected_suggestion' in f.get('suggestion','') for f in found)):
    sys.exit(0)
sys.exit(1)
" 2>/dev/null; then
        echo "  PASS: $desc"
        PASS=$((PASS + 1))
    else
        echo "  FAIL: $desc"
        echo "$json" | python3 -c "import json,sys; [print('        ',e['category'],e['word'][:40],'|',e['rule']) for e in json.load(sys.stdin)]" 2>/dev/null
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
found = [e for e in errors if e['word'] == '$word' or '$word' in e['word']]
sys.exit(0 if not found else 1)
" 2>/dev/null; then
        echo "  PASS: $desc"
        PASS=$((PASS + 1))
    else
        echo "  FAIL: $desc"
        FAIL=$((FAIL + 1))
    fi
}

check_grammar() {
    local desc="$1"
    local sentence_fragment="$2"
    local json="$3"

    if echo "$json" | python3 -c "
import json, sys
errors = json.load(sys.stdin)
found = [e for e in errors if e['category'] == 'grammar' and '$sentence_fragment' in e.get('sentence','')]
sys.exit(0 if found else 1)
" 2>/dev/null; then
        echo "  PASS: $desc"
        PASS=$((PASS + 1))
    else
        echo "  FAIL: $desc"
        FAIL=$((FAIL + 1))
    fi
}

# Type text character by character
type_text() {
    local text="$1"
    for (( i=0; i<${#text}; i++ )); do
        local char="${text:$i:1}"
        if [ "$char" = $'\n' ]; then
            key_press "return"
        else
            osascript -e "tell application \"Microsoft Word\" to type text selection text \"$char\"" 2>/dev/null
        fi
        sleep $DELAY
    done
}

# Press a key via System Events
key_press() {
    local key="$1"
    case "$key" in
        return)     osascript -e 'tell application "System Events" to keystroke return' 2>/dev/null ;;
        backspace)  osascript -e 'tell application "System Events" to key code 51' 2>/dev/null ;;
        delete)     osascript -e 'tell application "System Events" to key code 117' 2>/dev/null ;;  # forward delete
        left)       osascript -e 'tell application "System Events" to key code 123' 2>/dev/null ;;
        right)      osascript -e 'tell application "System Events" to key code 124' 2>/dev/null ;;
        up)         osascript -e 'tell application "System Events" to key code 126' 2>/dev/null ;;
        down)       osascript -e 'tell application "System Events" to key code 125' 2>/dev/null ;;
        home)       osascript -e 'tell application "System Events" to key code 115' 2>/dev/null ;;  # fn+left
        end)        osascript -e 'tell application "System Events" to key code 119' 2>/dev/null ;;  # fn+right
        cmd_end)    osascript -e 'tell application "System Events" to key code 125 using command down' 2>/dev/null ;;  # Cmd+Down = end of doc
        cmd_home)   osascript -e 'tell application "System Events" to key code 126 using command down' 2>/dev/null ;;  # Cmd+Up = start of doc
        cmd_left)   osascript -e 'tell application "System Events" to key code 123 using command down' 2>/dev/null ;;  # start of line
        cmd_right)  osascript -e 'tell application "System Events" to key code 124 using command down' 2>/dev/null ;;  # end of line
        opt_left)   osascript -e 'tell application "System Events" to key code 123 using option down' 2>/dev/null ;;  # word left
        opt_right)  osascript -e 'tell application "System Events" to key code 124 using option down' 2>/dev/null ;;  # word right
        shift_left) osascript -e 'tell application "System Events" to key code 123 using shift down' 2>/dev/null ;;  # select left
        shift_right) osascript -e 'tell application "System Events" to key code 124 using shift down' 2>/dev/null ;;  # select right
        cmd_a)      osascript -e 'tell application "System Events" to keystroke "a" using command down' 2>/dev/null ;;  # select all
        cmd_x)      osascript -e 'tell application "System Events" to keystroke "x" using command down' 2>/dev/null ;;  # cut
        cmd_v)      osascript -e 'tell application "System Events" to keystroke "v" using command down' 2>/dev/null ;;  # paste
        cmd_c)      osascript -e 'tell application "System Events" to keystroke "c" using command down' 2>/dev/null ;;  # copy
        cmd_z)      osascript -e 'tell application "System Events" to keystroke "z" using command down' 2>/dev/null ;;  # undo
        cmd_shift_z) osascript -e 'tell application "System Events" to keystroke "z" using {command down, shift down}' 2>/dev/null ;;  # redo
        # Select word: option+shift+left
        sel_word_left) osascript -e 'tell application "System Events" to key code 123 using {option down, shift down}' 2>/dev/null ;;
        sel_word_right) osascript -e 'tell application "System Events" to key code 124 using {option down, shift down}' 2>/dev/null ;;
        # Select to end of line
        sel_to_end) osascript -e 'tell application "System Events" to key code 124 using {command down, shift down}' 2>/dev/null ;;
        sel_to_start) osascript -e 'tell application "System Events" to key code 123 using {command down, shift down}' 2>/dev/null ;;
    esac
    sleep 0.1
}

# Press key N times
repeat_key() {
    local key="$1"
    local n="$2"
    for (( i=0; i<n; i++ )); do
        key_press "$key"
    done
}

# Select current line and delete
delete_current_line() {
    key_press cmd_left
    key_press sel_to_end
    key_press backspace
    key_press backspace  # also delete the newline
}

# Go to end of document
go_to_end() {
    key_press cmd_end
}

# Go to start of document
go_to_start() {
    key_press cmd_home
}

echo "=== NorskTale Error Detection Test ==="
echo "  Simulating real user keyboard interaction"
echo ""

# Activate Word
osascript -e 'tell application "Microsoft Word" to activate' 2>/dev/null
sleep 1

# ============================================================
# Test 1: Basic spelling error
# ============================================================
echo "Test 1: Spelling error 'somx' → 'som'"
go_to_end
key_press return
type_text "Fotball er en morsom sport somx er veldig morsom."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "somx detected with suggestion 'som'" "somx" "som" "$ERRORS"

# Clean up: select line and delete
key_press cmd_left
key_press sel_to_end
key_press backspace
key_press backspace
sleep 2

# ============================================================
# Test 2: Correct text — no false positives
# ============================================================
echo ""
echo "Test 2: Correct text — no false positives"
go_to_end
key_press return
type_text "Fotball er en morsom sport."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "sport not flagged" "sport" "$ERRORS"
check_no_error "morsom not flagged" "morsom" "$ERRORS"

# Clean up
key_press cmd_left
key_press sel_to_end
key_press backspace
key_press backspace
sleep 2

# ============================================================
# Test 3: Multiple errors in one sentence
# ============================================================
echo ""
echo "Test 3: Multiple errors in one sentence"
go_to_end
key_press return
type_text "Jeg liker aa spise matx og drikkx."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "matx detected" "matx" "" "$ERRORS"
check_error "drikkx detected" "drikkx" "" "$ERRORS"

# Clean up
key_press cmd_left
key_press sel_to_end
key_press backspace
key_press backspace
sleep 2

# ============================================================
# Test 4: Type misspelled, then fix with backspace
# ============================================================
echo ""
echo "Test 4: Type misspelled word, fix with backspace"
go_to_end
key_press return
type_text "Jeg liker fotbalx"
sleep 3
# Fix: backspace one char and type correct
key_press backspace
type_text "l. "
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "fotball not flagged after fix" "fotball" "$ERRORS"

# Clean up
key_press cmd_left
key_press sel_to_end
key_press backspace
key_press backspace
sleep 2

# ============================================================
# Test 5: Grammar error — gender mismatch
# ============================================================
echo ""
echo "Test 5: Grammar error — 'morsom spor' gender mismatch"
go_to_end
key_press return
type_text "Fotball er en morsom spor."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_grammar "gender mismatch detected" "spor" "$ERRORS"

# Clean up
key_press cmd_left
key_press sel_to_end
key_press backspace
key_press backspace
sleep 2

# ============================================================
# Test 6: Delete sentence with error — stale error should be gone
# ============================================================
echo ""
echo "Test 6: Delete sentence with error — no stale errors"
go_to_end
key_press return
type_text "Dette er en feilx i teksten."
sleep 5
ERRORS_BEFORE=$(curl -sk "$ENDPOINT")
check_error "feilx detected before delete" "feilx" "" "$ERRORS_BEFORE"

# Now delete the whole line
key_press cmd_left
key_press sel_to_end
key_press backspace
key_press backspace
sleep 5
ERRORS_AFTER=$(curl -sk "$ENDPOINT")
check_no_error "feilx gone after delete" "feilx" "$ERRORS_AFTER"
sleep 2

# ============================================================
# Test 7: Edit middle of word to fix error
# ============================================================
echo ""
echo "Test 7: Edit middle of word — navigate with arrows"
go_to_end
key_press return
type_text "Han liker fotboll veldig godt."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "fotboll detected" "fotboll" "" "$ERRORS"

# Fix: move cursor into "fotboll" — navigate to the 'o' after 'b', delete it, type 'a'
# "Han liker fotboll veldig godt."
# "Han liker fotb" = 14 chars, then 'o' is at position 14
key_press cmd_left    # start of line
repeat_key right 14   # move to 'o' in fotb|oll
key_press delete      # delete 'o' forward
type_text "a"         # now it's "fotball"
type_text " "         # trigger recheck
key_press backspace   # remove extra space
sleep 8
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "fotball correct after edit" "fotboll" "$ERRORS"

# Clean up
key_press cmd_left
key_press sel_to_end
key_press backspace
key_press backspace
sleep 2

# ============================================================
# Test 8: Split sentence by pressing Enter
# ============================================================
echo ""
echo "Test 8: Split sentence with Enter"
go_to_end
key_press return
type_text "Fotball er morsomt somx er fint."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "somx detected in single line" "somx" "" "$ERRORS"

# Split: move cursor between "morsomt" and "somx" and press Enter
key_press cmd_left
repeat_key right 19   # after "morsomt " (19 chars)
key_press return
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "somx still detected after split" "somx" "" "$ERRORS"

# Clean up both lines
key_press cmd_left
key_press sel_to_end
key_press backspace
key_press backspace
key_press cmd_left
key_press sel_to_end
key_press backspace
key_press backspace
sleep 2

# ============================================================
# Test 9: Replace correct word with misspelled
# ============================================================
echo ""
echo "Test 9: Replace correct word with misspelled"
go_to_end
key_press return
type_text "Jeg spiller fotball hver dag."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "fotball correct" "fotball" "$ERRORS"

# Select "fotball" and replace with "fotboll"
# "Jeg spiller fotball hver dag." — fotball starts at pos 12
key_press cmd_left
repeat_key right 12
# Select "fotball" (7 chars)
repeat_key shift_right 7
type_text "fotboll"
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "fotboll detected after replace" "fotboll" "" "$ERRORS"

# Clean up
key_press cmd_left
key_press sel_to_end
key_press backspace
key_press backspace
sleep 2

# ============================================================
# Test 10: Rapid typing — no crash
# ============================================================
echo ""
echo "Test 10: Rapid typing — no crash"
go_to_end
key_press return
DELAY=0.05  # fast typing
type_text "Dette er en rask test med mange ord som skal skrives fort uten feilx."
DELAY=0.15  # restore normal
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "feilx detected after rapid typing" "feilx" "" "$ERRORS"

# Clean up
key_press cmd_left
key_press sel_to_end
key_press backspace
key_press backspace
sleep 1

echo ""
echo "=== Results: $PASS passed, $FAIL failed ==="
