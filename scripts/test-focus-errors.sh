#!/usr/bin/env bash
# Focus-switching regression test for Spell.
#
# Reproduces the bug where Word errors disappear or BERT suggestions leak
# across windows when the user briefly tabs to Slack/Safari/Terminal.
#
# This test never modifies any document — it observes the live state via
# Spell's HTTPS endpoints. The user is expected to have a Word document
# open with some misspelled words already typed in. We:
#   1. Activate Word, snapshot /errors + /completions.
#   2. Activate Safari, snapshot — assert Word errors are still in /errors.
#   3. Activate Terminal, snapshot — same assertion.
#   4. Activate Word again, snapshot — assert errors still there.
#   5. Compare /completions between Word and Safari — they MUST differ
#      (different surrounding text → different BERT predictions).
#
# Requires:
#   - Spell (`acatts-rust`) running, HTTPS port 3000
#   - Microsoft Word installed and a doc with errors typed in
#   - `jq` for JSON parsing

set -uo pipefail

PORT=3000
ERRORS_URL="https://localhost:${PORT}/errors"
COMPLETIONS_URL="https://localhost:${PORT}/completions"
UI_STATE_URL="https://localhost:${PORT}/ui-state"

CURL="curl -sk --max-time 5"

ui_state() { $CURL "$UI_STATE_URL"; }

fail() { echo "FAIL: $*" >&2; exit 1; }
pass() { echo "PASS: $*"; }
info() { echo "INFO: $*"; }

errors_dump() {
  $CURL "$ERRORS_URL"
}

errors_count() {
  $CURL "$ERRORS_URL" | jq 'length'
}

completions_dump() {
  $CURL "$COMPLETIONS_URL"
}

completions_words() {
  # Concatenate left+right completion words, sorted, joined — used to
  # detect whether the bulb panel is showing different predictions for
  # different foreground apps.
  $CURL "$COMPLETIONS_URL" \
    | jq -r '[.completions[].word, .open_completions[].word] | sort | join(",")'
}

activate() {
  osascript -e "tell application \"$1\" to activate" >/dev/null
  sleep 2  # let foreground poller + bridge react
}

# --- 0. Sanity: server reachable ---
if ! $CURL --max-time 3 "$ERRORS_URL" >/dev/null; then
  fail "Spell HTTPS server not reachable at $ERRORS_URL — is the app running?"
fi
if ! $CURL --max-time 3 "$COMPLETIONS_URL" >/dev/null; then
  fail "Spell /completions not reachable — was the latest build deployed?"
fi
if ! $CURL --max-time 3 "$UI_STATE_URL" >/dev/null; then
  fail "Spell /ui-state not reachable — was the latest build deployed?"
fi
pass "server reachable; /errors, /completions and /ui-state respond"

# --- 1. Word foreground ---
activate "Microsoft Word"
sleep 2  # give pipeline a beat
err_word=$(errors_dump)
comp_word=$(completions_dump)
n_word=$(echo "$err_word" | jq 'length')
info "Word /errors (n=$n_word): $err_word"
info "Word /completions: $comp_word"
if [[ "$n_word" == "0" ]]; then
  fail "step 1 — Word foreground: /errors is empty. Did you type misspellings into the doc before running this test?"
fi
pass "step 1 — Word foreground: $n_word errors detected"

# 1c — /ui-state matches what the user sees
ui_word=$(ui_state)
info "Word /ui-state: $ui_word"
pencil_word=$(echo "$ui_word" | jq -r '.pencil_visible')
tips_word=$(echo "$ui_word" | jq -r '.tips_count')
fg_word=$(echo "$ui_word" | jq -r '.fg_app')
if [[ "$pencil_word" != "true" ]]; then
  fail "step 1c — Word foreground: pencil_visible=$pencil_word (expected true)"
fi
if [[ "$tips_word" != "$n_word" ]]; then
  fail "step 1c — Word foreground: tips_count=$tips_word but /errors has $n_word entries"
fi
if [[ "$fg_word" != "microsoft word" ]]; then
  fail "step 1c — Word foreground: fg_app='$fg_word' (expected 'microsoft word')"
fi
pass "step 1c — Word /ui-state matches: pencil_visible=true tips_count=$tips_word"

# Snapshot the Word-side error set (word field) for later comparison.
word_set=$(echo "$err_word" | jq -r '[.[].word] | sort | join(",")')
word_comp_set=$(echo "$comp_word" | completions_words)
info "Word error words: $word_set"

# --- 1b. No duplicates ---
# Each (word, sentence) pair must appear at most once. The pencil panel
# showing 4 cards for 2 underlines is the regression we're guarding here.
dup_count=$(echo "$err_word" | jq '[.[] | "\(.word)|\(.sentence)"] | (length - (unique | length))')
if [[ "$dup_count" != "0" ]]; then
  dups=$(echo "$err_word" | jq -c '[.[] | "\(.word)|\(.sentence)"] | group_by(.) | map(select(length>1) | .[0])')
  fail "step 1b — /errors has $dup_count duplicate entries: $dups"
fi
pass "step 1b — no duplicate errors in /errors"

# --- 2. Safari foreground ---
activate "Safari"
err_safari=$(errors_dump)
comp_safari=$(completions_dump)
n_safari=$(echo "$err_safari" | jq 'length')
info "Safari /errors (n=$n_safari): $err_safari"
info "Safari /completions: $comp_safari"
if [[ "$n_safari" == "0" ]]; then
  fail "step 2 — Safari foreground: Word errors disappeared from /errors"
fi
# Verify the same word set is still there
safari_set=$(echo "$err_safari" | jq -r '[.[].word] | sort | join(",")')
missing_after_safari=$(comm -23 <(echo "$word_set" | tr ',' '\n' | sort -u) <(echo "$safari_set" | tr ',' '\n' | sort -u) | head -1)
if [[ -n "$missing_after_safari" ]]; then
  fail "step 2 — Safari foreground: words missing from /errors (e.g. '$missing_after_safari'; was '$word_set', now '$safari_set')"
fi
pass "step 2 — Safari foreground: $n_safari errors retained"

# 2b — pencil panel hidden in Safari
ui_safari=$(ui_state)
pencil_safari=$(echo "$ui_safari" | jq -r '.pencil_visible')
tips_safari=$(echo "$ui_safari" | jq -r '.tips_count')
info "Safari /ui-state: $ui_safari"
if [[ "$pencil_safari" != "false" ]]; then
  fail "step 2b — Safari foreground: pencil_visible=$pencil_safari (expected false)"
fi
if [[ "$tips_safari" != "0" ]]; then
  fail "step 2b — Safari foreground: tips_count=$tips_safari (expected 0)"
fi
pass "step 2b — Safari /ui-state: pencil hidden, tips=0 (no Word leak)"

# --- 3a. Slack foreground (if running) ---
if osascript -e 'tell application "System Events" to (name of every application process)' 2>/dev/null | grep -qi '\bslack\b'; then
  activate "Slack"
  err_slack=$(errors_dump)
  comp_slack=$(completions_dump)
  n_slack=$(echo "$err_slack" | jq 'length')
  info "Slack /errors (n=$n_slack): $err_slack"
  info "Slack /completions: $comp_slack"
  if [[ "$n_slack" == "0" ]]; then
    fail "step 3a — Slack foreground: Word errors disappeared from /errors"
  fi
  slack_set=$(echo "$err_slack" | jq -r '[.[].word] | sort | join(",")')
  missing_after_slack=$(comm -23 <(echo "$word_set" | tr ',' '\n' | sort -u) <(echo "$slack_set" | tr ',' '\n' | sort -u) | head -1)
  if [[ -n "$missing_after_slack" ]]; then
    fail "step 3a — Slack foreground: word missing (e.g. '$missing_after_slack')"
  fi
  pass "step 3a — Slack foreground: $n_slack errors retained"

  # 3a-b — pencil panel hidden in Slack
  ui_slack=$(ui_state)
  pencil_slack=$(echo "$ui_slack" | jq -r '.pencil_visible')
  tips_slack=$(echo "$ui_slack" | jq -r '.tips_count')
  info "Slack /ui-state: $ui_slack"
  if [[ "$pencil_slack" != "false" ]]; then
    fail "step 3a-b — Slack foreground: pencil_visible=$pencil_slack (expected false; Word errors must NOT show in Slack)"
  fi
  if [[ "$tips_slack" != "0" ]]; then
    fail "step 3a-b — Slack foreground: tips_count=$tips_slack (expected 0)"
  fi
  pass "step 3a-b — Slack /ui-state: pencil hidden, tips=0 (no Word leak)"
else
  info "step 3a — Slack not running, skipped"
fi

# --- 3. Terminal foreground ---
activate "Terminal"
err_term=$(errors_dump)
comp_term=$(completions_dump)
n_term=$(echo "$err_term" | jq 'length')
info "Terminal /errors (n=$n_term): $err_term"
info "Terminal /completions: $comp_term"
if [[ "$n_term" == "0" ]]; then
  fail "step 3 — Terminal foreground: Word errors disappeared from /errors"
fi
term_set=$(echo "$err_term" | jq -r '[.[].word] | sort | join(",")')
missing_after_term=$(comm -23 <(echo "$word_set" | tr ',' '\n' | sort -u) <(echo "$term_set" | tr ',' '\n' | sort -u) | head -1)
if [[ -n "$missing_after_term" ]]; then
  fail "step 3 — Terminal foreground: word missing (e.g. '$missing_after_term')"
fi
pass "step 3 — Terminal foreground: $n_term errors retained"

# 3b — pencil hidden in Terminal
ui_term=$(ui_state)
pencil_term=$(echo "$ui_term" | jq -r '.pencil_visible')
tips_term=$(echo "$ui_term" | jq -r '.tips_count')
info "Terminal /ui-state: $ui_term"
if [[ "$pencil_term" != "false" ]]; then
  fail "step 3b — Terminal foreground: pencil_visible=$pencil_term (expected false)"
fi
if [[ "$tips_term" != "0" ]]; then
  fail "step 3b — Terminal foreground: tips_count=$tips_term (expected 0)"
fi
pass "step 3b — Terminal /ui-state: pencil hidden, tips=0"

# --- 4. Back to Word ---
activate "Microsoft Word"
sleep 1
err_word2=$(errors_dump)
comp_word2=$(completions_dump)
n_word2=$(echo "$err_word2" | jq 'length')
info "Word(2) /errors (n=$n_word2): $err_word2"
info "Word(2) /completions: $comp_word2"
if [[ "$n_word2" == "0" ]]; then
  fail "step 4 — Back to Word: errors gone after detour"
fi
word2_set=$(echo "$err_word2" | jq -r '[.[].word] | sort | join(",")')
missing_after_return=$(comm -23 <(echo "$word_set" | tr ',' '\n' | sort -u) <(echo "$word2_set" | tr ',' '\n' | sort -u) | head -1)
if [[ -n "$missing_after_return" ]]; then
  fail "step 4 — Back to Word: word missing (e.g. '$missing_after_return'; was '$word_set', now '$word2_set')"
fi
pass "step 4 — Back to Word: $n_word2 errors intact"

# 4b — pencil + badge restored on return
ui_word2=$(ui_state)
pencil_word2=$(echo "$ui_word2" | jq -r '.pencil_visible')
tips_word2=$(echo "$ui_word2" | jq -r '.tips_count')
info "Word(2) /ui-state: $ui_word2"
if [[ "$pencil_word2" != "true" ]]; then
  fail "step 4b — Back to Word: pencil_visible=$pencil_word2 (expected true)"
fi
if [[ "$tips_word2" != "$n_word2" ]]; then
  fail "step 4b — Back to Word: tips_count=$tips_word2 != /errors length $n_word2"
fi
pass "step 4b — Word(2) /ui-state: pencil_visible=true tips_count=$tips_word2"

# --- 5. Completions track foreground ---
# Terminal's completions should differ from both Word readings (Word's
# surrounding text is Norwegian; Terminal's is English/CLI). And the
# Word(2) reading must NOT equal Terminal's reading — that would mean the
# bulb panel is stuck on the previous foreground's predictions.
word_comp_set=$($CURL "$ERRORS_URL" >/dev/null; echo "$comp_word" | jq -r '[.completions[].word, .open_completions[].word] | sort | join(",")')
term_comp_set=$(echo "$comp_term" | jq -r '[.completions[].word, .open_completions[].word] | sort | join(",")')
word2_comp_set=$(echo "$comp_word2" | jq -r '[.completions[].word, .open_completions[].word] | sort | join(",")')
info "Word(1) completion words: '$word_comp_set'"
info "Terminal completion words: '$term_comp_set'"
info "Word(2) completion words: '$word2_comp_set'"

# 5a — when BOTH are non-empty, sanity check that Word and Terminal yield
# DIFFERENT predictions (different surrounding text → different BERT
# output). Empty completions are allowed — BERT only fills on cursor move
# and a pure focus change doesn't trigger one.
if [[ -n "$word_comp_set" && -n "$term_comp_set" && "$word_comp_set" == "$term_comp_set" ]]; then
  fail "step 5a — Word and Terminal /completions identical ('$term_comp_set'); bulb is not tracking foreground"
fi
pass "step 5a — Word vs Terminal completions are not stale-shared (or one was empty)"

# 5b — Word(2) MUST NOT show completions from any of the detour apps.
# This is the regression we care about — BERT predictions from
# Slack/Terminal leaking back into Word's bulb panel on return.
if [[ -n "$term_comp_set" && "$word2_comp_set" == "$term_comp_set" ]]; then
  fail "step 5b — Word(2) /completions identical to Terminal's ('$term_comp_set'); bulb did not refresh on return to Word"
fi
if [[ -n "${comp_slack:-}" ]]; then
  slack_comp_set=$(echo "$comp_slack" | jq -r '[.completions[].word, .open_completions[].word] | sort | join(",")')
  if [[ -n "$slack_comp_set" && "$word2_comp_set" == "$slack_comp_set" ]]; then
    fail "step 5b — Word(2) /completions identical to Slack's ('$slack_comp_set'); bulb leaked Slack predictions back into Word"
  fi
fi
pass "step 5b — Word(2) /completions refreshed (no stale leak from Terminal or Slack)"

echo "ALL PASS"
