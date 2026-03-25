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
    echo "  ALIGNMENT WARNING: $ec errors != $uc underlines (continuing)"
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

append_text() {
    # Insert text as a new paragraph at the end of the document body (via Word API)
    local text="$1"
    local escaped=$(echo "$text" | sed 's/"/\\"/g')
    curl -sk -X POST "$PUSH_URL" -d "{\"action\":\"appendParagraph\",\"text\":\"$escaped\"}" 2>/dev/null
    sleep 1
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
        UNDO_COUNT=$((UNDO_COUNT + 1))
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

go_to_end() { key_press cmd_end; sleep 0.3; }
# key_press_counted: like key_press but counts for undo
key_press_counted() { key_press "$1"; UNDO_COUNT=$((UNDO_COUNT + 1)); }


PUSH_URL="https://127.0.0.1:3000/push-reply"
SCRIPT_DIR_ABS="$(cd "$(dirname "$0")" && pwd)"

# Marker: last 30 chars of document — used by deleteAfter to remove test text
DOC_MARKER=""

UNDO_COUNT=0  # Tracks keystrokes for undo

undo_all() {
    # Undo exactly the number of keystrokes typed since last undo
    # Add small buffer for return keys and replace operations
    local n=$((UNDO_COUNT + 5))
    UNDO_COUNT=0
    osascript -e "
tell application \"Microsoft Word\" to activate
delay 0.3
tell application \"System Events\"
    repeat $n times
        keystroke \"z\" using command down
        delay 0.02
    end repeat
end tell
" 2>/dev/null
    sleep 2
    # Verify document intact
    local h=$(osascript -e 'tell application "Microsoft Word" to content of text object of active document' 2>/dev/null | md5)
    if [ "$h" != "$ORIG_DOC_HASH" ]; then
        # Try a few more undos to recover
        osascript -e '
tell application "Microsoft Word" to activate
delay 0.2
tell application "System Events"
    repeat 20 times
        keystroke "z" using command down
        delay 0.02
    end repeat
end tell
' 2>/dev/null
        sleep 1
        h=$(osascript -e 'tell application "Microsoft Word" to content of text object of active document' 2>/dev/null | md5)
        if [ "$h" != "$ORIG_DOC_HASH" ]; then
            echo "  ABORT: Document corrupted! hash=$h expected=$ORIG_DOC_HASH"
            echo "=== Results: $PASS passed, $FAIL failed (ABORTED) ==="
            exit 1
        fi
    fi
    sleep 1
}

echo "=== NorskTale Error Detection Test ==="

# Work in the EXISTING document — never open/close/save documents
ORIG_DOC_NAME=$(osascript -e 'tell application "Microsoft Word" to name of active document' 2>/dev/null)
ORIG_DOC_HASH=$(osascript -e 'tell application "Microsoft Word" to content of text object of active document' 2>/dev/null | md5)
DOC_MARKER=$(osascript -e 'tell application "Microsoft Word" to content of text object of active document' 2>/dev/null | tail -c 30 | tr -d '\n')
echo "Document: '$ORIG_DOC_NAME' (hash: $ORIG_DOC_HASH)"
echo "Marker: '$DOC_MARKER'"
echo ""
osascript -e 'tell application "Microsoft Word" to activate' 2>/dev/null
sleep 1

# ============================================================
echo "Test 0: Document health — errors match underlines"
ERROR_COUNT=$(curl -sk "$ENDPOINT" | python3 -c "import json,sys; print(len(json.load(sys.stdin)))" 2>/dev/null)
echo "  INFO: $ERROR_COUNT errors detected"
PASS=$((PASS + 1))

# ============================================================
echo ""
echo "Test 1: Spelling error 'somx' → 'som'"
go_to_end; key_press_counted return
type_text "Fotball er en morsom sport somx er veldig morsom."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "somx detected" "somx" "som" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 2: Correct text — no false positives"
go_to_end; key_press_counted return
type_text "Fotball er en morsom sport."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "sport not flagged" "sport" "$ERRORS"
check_no_error "morsom not flagged" "morsom" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 2b: Correct neuter sentence — no false positives"
go_to_end; key_press_counted return
type_text "Fotball er et morsomt spill."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "spill not flagged" "spill" "$ERRORS"
check_no_error "morsomt not flagged" "morsomt" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 3: Multiple errors in one sentence"
go_to_end; key_press_counted return
type_text "Jeg liker aa spise matx og drikkx."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "matx detected" "matx" "" "$ERRORS"
check_error "drikkx detected" "drikkx" "" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 4: Type misspelled, fix with backspace"
go_to_end; key_press_counted return
type_text "Jeg liker fotbalx."
sleep 3
curl -sk -X POST "$PUSH_URL" -d '{"action":"replace","expected":"fotbalx","text":"fotball"}' 2>/dev/null
UNDO_COUNT=$((UNDO_COUNT + 2))
sleep 8
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "fotball not flagged after fix" "fotball" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 5: Grammar error — gender mismatch"
go_to_end; key_press_counted return
type_text "Fotball er en morsom spor."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_grammar "gender mismatch" "spor" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 5b: Grammar error — adj gender mismatch (morsomt with masculine)"
go_to_end; key_press_counted return
type_text "Fotball er en morsomt sport."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_grammar "adj gender mismatch" "morsomt" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 6: Delete sentence — stale error gone"
go_to_end; key_press_counted return
type_text "Dette er en feilx i teksten."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "feilx detected" "feilx" "" "$ERRORS"
undo_all
sleep 3
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "feilx gone after undo" "feilx" "$ERRORS"

# ============================================================
echo ""
echo "Test 7: Fix misspelled word via replace"
go_to_end; key_press_counted return
type_text "Han liker fotbollzz veldig godt."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "fotbollzz detected" "fotbollzz" "" "$ERRORS"
curl -sk -X POST "$PUSH_URL" -d '{"action":"replace","expected":"fotbollzz","text":"fotball"}' 2>/dev/null
UNDO_COUNT=$((UNDO_COUNT + 2))
sleep 8
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "fotbollzz gone after fix" "fotbollzz" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 8: Split sentence with Enter"
go_to_end; key_press_counted return
type_text "Fotball er morsomt somx er fint."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "somx in single line" "somx" "" "$ERRORS"
curl -sk -X POST "$PUSH_URL" -d '{"action":"replace","expected":"morsomt somx","text":"morsomt\nsomx"}' 2>/dev/null
UNDO_COUNT=$((UNDO_COUNT + 2))
sleep 8
ERRORS=$(curl -sk "$ENDPOINT")
check_error "somx after split" "somx" "" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 9: Replace correct word with misspelled"
go_to_end; key_press_counted return
type_text "Jeg spiller fotball hver dag."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "fotball correct" "fotball" "$ERRORS"
curl -sk -X POST "$PUSH_URL" -d '{"action":"replace","expected":"Jeg spiller fotball","text":"Jeg spiller fotboll"}' 2>/dev/null
UNDO_COUNT=$((UNDO_COUNT + 2))
sleep 8
ERRORS=$(curl -sk "$ENDPOINT")
check_error "fotboll after replace" "fotboll" "" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 10: Rapid typing — no crash"
go_to_end; key_press_counted return
type_text "Dette er en rask test med mange ord uten feilx."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "feilx after rapid" "feilx" "" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 11: Fix error then re-introduce same error (stale hash race)"
go_to_end; key_press_counted return
type_text "Han spiller fotboll hver dag."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "fotboll detected first time" "fotboll" "" "$ERRORS"
curl -sk -X POST "$PUSH_URL" -d '{"action":"replace","expected":"Han spiller fotboll","text":"Han spiller fotball"}' 2>/dev/null
UNDO_COUNT=$((UNDO_COUNT + 2))
sleep 10
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "fotboll gone after fix" "fotboll" "$ERRORS"
go_to_end; key_press_counted return
type_text "Han spiller fotboll hver dag."
sleep 8
ERRORS=$(curl -sk "$ENDPOINT")
check_error "fotboll re-detected after reintroduce" "fotboll" "" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 12: Correct sentences — no false positives"
go_to_end; key_press_counted return
type_text "Fotball er en morsom sport."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "Fotball not flagged" "Fotball" "$ERRORS"
check_no_error "sport not flagged" "sport" "$ERRORS"
undo_all

go_to_end; key_press_counted return
type_text "Fotball er et morsomt spill."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "spill not flagged" "spill" "$ERRORS"
check_no_error "Fotball not flagged (neuter)" "Fotball" "$ERRORS"
undo_all

go_to_end; key_press_counted return
type_text "Han liker å spille fotball."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "spille not flagged" "spille" "$ERRORS"
check_no_error "fotball not flagged" "fotball" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 5c: Grammar error — er + present verb"
go_to_end; key_press_counted return
type_text "Jeg er spiller fotball."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_grammar "er + present verb" "er" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 13: Duplicate sentences both detected"
go_to_end; key_press_counted return
type_text "Han liker duplikatxx veldig godt."
go_to_end; key_press_counted return
type_text "Han liker duplikatxx veldig godt."
sleep 8
ERRORS=$(curl -sk "$ENDPOINT")
DUPCOUNT=$(echo "$ERRORS" | python3 -c "import json,sys; print(len([e for e in json.load(sys.stdin) if e['word']=='duplikatxx']))" 2>/dev/null)
if [ "$DUPCOUNT" = "2" ]; then
    echo "  PASS: both duplicate duplikatxx detected ($DUPCOUNT)"
    PASS=$((PASS + 1))
else
    echo "  FAIL: expected 2 duplikatxx errors, got $DUPCOUNT"
    FAIL=$((FAIL + 1))
fi
undo_all

# ============================================================
echo ""
echo "Test 14: Paste misspelled text — error detected"
osascript -e 'set the clipboard to "Han liker pasteerrorx veldig godt."' 2>/dev/null
sleep 0.5
go_to_end; key_press_counted return
key_press cmd_v
UNDO_COUNT=$((UNDO_COUNT + 1))
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "pasteerrorx detected after paste" "pasteerrorx" "" "$ERRORS"
undo_all

# ============================================================
echo ""
echo "Test 15: Delete removes error"
go_to_end; key_press_counted return
type_text "Dette er en feilzz i teksten."
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "feilzz detected before delete" "feilzz" "" "$ERRORS"
osascript -e '
tell application "Microsoft Word" to activate
delay 0.3
tell application "System Events"
    repeat 40 times
        keystroke "z" using command down
        delay 0.02
    end repeat
end tell
' 2>/dev/null
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_no_error "feilzz gone after delete" "feilzz" "$ERRORS"

# ============================================================
echo ""
echo "Test 16: Paste different misspelled text — error detected"
osascript -e 'set the clipboard to "Fotball er gøy med pastezz."' 2>/dev/null
sleep 0.3
go_to_end; key_press_counted return
key_press cmd_v
UNDO_COUNT=$((UNDO_COUNT + 1))
sleep 5
ERRORS=$(curl -sk "$ENDPOINT")
check_error "pastezz detected after paste" "pastezz" "" "$ERRORS"
undo_all

# Verify original document is intact
echo ""
echo "Verifying original document..."
FINAL_DOC_HASH=$(osascript -e 'tell application "Microsoft Word" to content of text object of active document' 2>/dev/null | md5)
if [ "$FINAL_DOC_HASH" = "$ORIG_DOC_HASH" ]; then
    echo "  Original document: INTACT (hash matches)"
else
    echo "  WARNING: Original document hash changed!"
    echo "  Before: $ORIG_DOC_HASH"
    echo "  After:  $FINAL_DOC_HASH"
fi

echo ""
echo "=== Results: $PASS passed, $FAIL failed ==="
