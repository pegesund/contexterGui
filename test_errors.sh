#!/bin/bash
# Automated error detection test — simulates real user typing and editing
# Requires: Word open with NorskTale add-in connected, acatts-rust running
#
# Usage: ./test_errors.sh

ENDPOINT="https://127.0.0.1:3000/errors"
PASS=0
FAIL=0
DELAY=0.15

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# --- Helper functions ---

check_alignment() {
    # Retry alignment check up to 3 times with 3s wait (grammar actor may still be processing)
    for attempt in 1 2 3; do
        local ec=$(curl -sk "$ENDPOINT" | python3 -c "import json,sys; print(len(json.load(sys.stdin)))" 2>/dev/null)
        local uc=$(osascript "$SCRIPT_DIR/scan_underlines.applescript" 2>/dev/null | head -1 | grep -o '^[0-9]*')
        uc=${uc:-0}
        if [ "$ec" = "$uc" ]; then
            return
        fi
        if [ "$attempt" -lt 3 ]; then
            sleep 3
        fi
    done
    echo "  ALIGNMENT FAIL: $ec errors != $uc underlines"
    FAIL=$((FAIL + 1))
    echo "=== ABORTING: underline alignment broken ==="
    echo "=== Results: $PASS passed, $FAIL failed ==="
    exit 1
}

check_error() {
    local desc="$1" word="$2" expected="$3" json="$4"
    if echo "$json" | python3 -c "
import json, sys
errors = json.load(sys.stdin)
found = [e for e in errors if e['word'] == '$word' or '$word' in e['word'] or '$word' in e.get('sentence','')]
if found and (not '$expected' or any('$expected' in f.get('suggestion','') for f in found)):
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
    local desc="$1" word="$2" json="$3"
    if echo "$json" | python3 -c "
import json, sys
errors = json.load(sys.stdin)
found = [e for e in errors if e['word'] == '$word']
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
    local desc="$1" fragment="$2" json="$3"
    if echo "$json" | python3 -c "
import json, sys
errors = json.load(sys.stdin)
found = [e for e in errors if e['category'] == 'grammar' and '$fragment' in e.get('sentence','')]
sys.exit(0 if found else 1)
" 2>/dev/null; then
        echo "  PASS: $desc"
        PASS=$((PASS + 1))
    else
        echo "  FAIL: $desc"
        FAIL=$((FAIL + 1))
    fi
}

type_text() {
    local text="$1"
    for (( i=0; i<${#text}; i++ )); do
        local char="${text:$i:1}"
        if [ "$char" = $'\n' ]; then
            key_press return
        else
            osascript -e "tell application \"Microsoft Word\" to type text selection text \"$char\"" 2>/dev/null
        fi
        sleep $DELAY
    done
}

key_press() {
    local key="$1"
    case "$key" in
        return)       osascript -e 'tell application "System Events" to keystroke return' ;;
        backspace)    osascript -e 'tell application "System Events" to key code 51' ;;
        delete)       osascript -e 'tell application "System Events" to key code 117' ;;
        left)         osascript -e 'tell application "System Events" to key code 123' ;;
        right)        osascript -e 'tell application "System Events" to key code 124' ;;
        cmd_end)      osascript -e 'tell application "System Events" to key code 125 using command down' ;;
        cmd_left)     osascript -e 'tell application "System Events" to key code 123 using command down' ;;
        cmd_right)    osascript -e 'tell application "System Events" to key code 124 using command down' ;;
        opt_left)     osascript -e 'tell application "System Events" to key code 123 using option down' ;;
        opt_right)    osascript -e 'tell application "System Events" to key code 124 using option down' ;;
        shift_right)  osascript -e 'tell application "System Events" to key code 124 using shift down' ;;
        sel_to_end)   osascript -e 'tell application "System Events" to key code 124 using {command down, shift down}' ;;
        sel_to_start) osascript -e 'tell application "System Events" to key code 123 using {command down, shift down}' ;;
        sel_word_left) osascript -e 'tell application "System Events" to key code 123 using {option down, shift down}' ;;
        cmd_z)        osascript -e 'tell application "System Events" to keystroke "z" using command down' ;;
        cmd_x)        osascript -e 'tell application "System Events" to keystroke "x" using command down' ;;
        cmd_v)        osascript -e 'tell application "System Events" to keystroke "v" using command down' ;;
    esac 2>/dev/null
    sleep 0.1
}

repeat_key() {
    local key="$1" n="$2"
    for (( i=0; i<n; i++ )); do key_press "$key"; done
}

go_to_end() { key_press cmd_end; }

# Undo N times to restore document state
PUSH_URL="https://127.0.0.1:3000/push-reply"
DOC_MARKER="25.11.2022"
SCRIPT_DIR_ABS="$(cd "$(dirname "$0")" && pwd)"

undo_all() {
    # Delete test text after the original doc marker
    curl -sk -X POST "$PUSH_URL" -d "{\"action\":\"deleteAfter\",\"text\":\"$DOC_MARKER\"}" 2>/dev/null
    sleep 2
    # Delete again to catch any remnants
    curl -sk -X POST "$PUSH_URL" -d "{\"action\":\"deleteAfter\",\"text\":\"$DOC_MARKER\"}" 2>/dev/null
    sleep 2
    # Reload add-in so errors resync with actual doc content
    bash "$SCRIPT_DIR_ABS/reload_addin.sh"
    sleep 5
    check_alignment
}

echo "=== NorskTale Error Detection Test ==="
echo ""
osascript -e 'tell application "Microsoft Word" to activate' 2>/dev/null
sleep 1

# ============================================================
echo "Test 0: Document health — errors match underlines"
ERROR_COUNT=$(curl -sk "$ENDPOINT" | python3 -c "import json,sys; print(len(json.load(sys.stdin)))" 2>/dev/null)
UNDERLINE_COUNT=$(osascript "$(dirname "$0")/scan_underlines.applescript" 2>/dev/null | head -1 | grep -o '^[0-9]*')
UNDERLINE_COUNT=${UNDERLINE_COUNT:-0}
if [ "$ERROR_COUNT" = "$UNDERLINE_COUNT" ]; then
    echo "  PASS: $ERROR_COUNT errors = $UNDERLINE_COUNT underlines"
    PASS=$((PASS + 1))
else
    echo "  FAIL: $ERROR_COUNT errors != $UNDERLINE_COUNT underlines"
    echo "  Errors:"
    curl -sk "$ENDPOINT" | python3 -c "import json,sys; [print(f'    {e[\"word\"]}') for e in json.load(sys.stdin)]" 2>/dev/null
    echo "  Underlines:"
    osascript "$(dirname "$0")/scan_underlines.applescript" 2>/dev/null | tail -n +2
    FAIL=$((FAIL + 1))
    echo ""
    echo "=== ABORTING: document health check failed ==="
    echo "=== Results: $PASS passed, $FAIL failed ==="
    exit 1
fi

# ============================================================
echo ""
echo "Test 1: Spelling error 'somx' → 'som'"
go_to_end; key_press return
type_text "Fotball er en morsom sport somx er veldig morsom."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "somx detected" "somx" "som" "$ERRORS"
undo_all 60

# ============================================================
echo ""
echo "Test 2: Correct text — no false positives"
go_to_end; key_press return
type_text "Fotball er en morsom sport."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "sport not flagged" "sport" "$ERRORS"
check_no_error "morsom not flagged" "morsom" "$ERRORS"
undo_all 40

# ============================================================
echo ""
echo "Test 2b: Correct neuter sentence — no false positives"
go_to_end; key_press return
type_text "Fotball er et morsomt spill."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "spill not flagged" "spill" "$ERRORS"
check_no_error "morsomt not flagged" "morsomt" "$ERRORS"
undo_all 40

# ============================================================
echo ""
echo "Test 3: Multiple errors in one sentence"
go_to_end; key_press return
type_text "Jeg liker aa spise matx og drikkx."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "matx detected" "matx" "" "$ERRORS"
check_error "drikkx detected" "drikkx" "" "$ERRORS"
undo_all 50

# ============================================================
echo ""
echo "Test 4: Type misspelled, fix with backspace"
go_to_end; key_press return
type_text "Jeg liker fotbalx"
sleep 3
key_press backspace
type_text "l. "
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "fotball not flagged after fix" "fotball" "$ERRORS"
undo_all 30

# ============================================================
echo ""
echo "Test 5: Grammar error — gender mismatch"
go_to_end; key_press return
type_text "Fotball er en morsom spor."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_grammar "gender mismatch" "spor" "$ERRORS"
undo_all 40

# ============================================================
echo ""
echo "Test 5b: Grammar error — adj gender mismatch (morsomt with masculine)"
go_to_end; key_press return
type_text "Fotball er en morsomt sport."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_grammar "adj gender mismatch" "morsomt" "$ERRORS"
undo_all 40

# ============================================================
echo ""
echo "Test 6: Delete sentence — stale error gone"
go_to_end; key_press return
type_text "Dette er en feilx i teksten."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "feilx detected" "feilx" "" "$ERRORS"
# Undo the whole line (removes all typed text)
undo_all 40
sleep 3
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "feilx gone after undo" "feilx" "$ERRORS"

# ============================================================
echo ""
echo "Test 7: Edit middle of word with arrows"
go_to_end; key_press return
type_text "Han liker fotboll veldig godt."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "fotboll detected" "fotboll" "" "$ERRORS"
# Navigate to 'o' in fotboll: go to start of line, right 14 times
key_press cmd_left
repeat_key right 14
key_press delete   # delete 'o'
type_text "a"      # insert 'a' → fotball
sleep 8
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "fotboll gone after fix" "fotboll" "$ERRORS"
undo_all 50

# ============================================================
echo ""
echo "Test 8: Split sentence with Enter"
go_to_end; key_press return
type_text "Fotball er morsomt somx er fint."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "somx in single line" "somx" "" "$ERRORS"
# Move to position 19 (after "morsomt ") and split
key_press cmd_left
repeat_key right 19
key_press return
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "somx after split" "somx" "" "$ERRORS"
undo_all 50

# ============================================================
echo ""
echo "Test 9: Replace correct word with misspelled"
go_to_end; key_press return
type_text "Jeg spiller fotball hver dag."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "fotball correct" "fotball" "$ERRORS"
# Select "fotball": go to start, right 12, select 7 chars
key_press cmd_left
repeat_key right 12
repeat_key shift_right 7
type_text "fotboll"
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "fotboll after replace" "fotboll" "" "$ERRORS"
undo_all 50

# ============================================================
echo ""
echo "Test 10: Rapid typing — no crash"
go_to_end; key_press return
DELAY=0.05
type_text "Dette er en rask test med mange ord uten feilx."
DELAY=0.15
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "feilx after rapid" "feilx" "" "$ERRORS"
undo_all 60

# ============================================================
echo ""
echo "Test 11: Fix error then re-introduce same error (stale hash race)"
go_to_end; key_press return
type_text "Han spiller fotboll hver dag."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "fotboll detected first time" "fotboll" "" "$ERRORS"
# Fix: select "fotboll" (7 chars starting at pos 12) and replace with "fotball"
key_press cmd_left
repeat_key right 12
repeat_key shift_right 7
type_text "fotball"
go_to_end
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "fotboll gone after fix" "fotboll" "$ERRORS"
# Re-introduce: select "fotball" (7 chars at pos 12) and replace with "fotboll"
key_press cmd_left
repeat_key right 12
repeat_key shift_right 7
type_text "fotboll"
go_to_end
sleep 8
ERRORS=$(curl -sk "$ENDPOINT")
check_error "fotboll re-detected after reintroduce" "fotboll" "" "$ERRORS"
undo_all 50

# ============================================================
echo ""
echo "Test 12: Correct sentences — no false positives"
go_to_end; key_press return
type_text "Fotball er en morsom sport."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "Fotball not flagged" "Fotball" "$ERRORS"
check_no_error "sport not flagged" "sport" "$ERRORS"
undo_all 40

go_to_end; key_press return
type_text "Fotball er et morsomt spill."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "spill not flagged" "spill" "$ERRORS"
check_no_error "Fotball not flagged (neuter)" "Fotball" "$ERRORS"
undo_all 40

go_to_end; key_press return
type_text "Han liker å spille fotball."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "spille not flagged" "spille" "$ERRORS"
check_no_error "fotball not flagged" "fotball" "$ERRORS"
undo_all 40

# ============================================================
echo ""
echo "Test 5c: Grammar error — er + present verb"
go_to_end; key_press return
type_text "Jeg er spiller fotball."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_grammar "er + present verb" "er" "$ERRORS"
undo_all 40

# ============================================================
echo ""
echo "Test 13: Duplicate sentences both detected"
go_to_end; key_press return
type_text "Han liker fotbollx veldig godt."
key_press return
type_text "Han liker fotbollx veldig godt."
sleep 8
ERRORS=$(curl -sk "$ENDPOINT")
# Count how many fotbollx errors
FOTBOLLX_COUNT=$(echo "$ERRORS" | python3 -c "import json,sys; print(len([e for e in json.load(sys.stdin) if e['word']=='fotbollx']))" 2>/dev/null)
if [ "$FOTBOLLX_COUNT" = "2" ]; then
    echo "  PASS: both duplicate fotbollx detected ($FOTBOLLX_COUNT)"
    PASS=$((PASS + 1))
else
    echo "  FAIL: expected 2 fotbollx errors, got $FOTBOLLX_COUNT"
    FAIL=$((FAIL + 1))
fi
undo_all 50

echo ""
echo "=== Results: $PASS passed, $FAIL failed ==="
