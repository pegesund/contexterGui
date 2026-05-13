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

CURL="curl -sk --max-time 5"

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
pass "server reachable; /errors and /completions respond"

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

# Snapshot the Word-side error set (word field) for later comparison.
word_set=$(echo "$err_word" | jq -r '[.[].word] | sort | join(",")')
word_comp_set=$(echo "$comp_word" | completions_words)
info "Word error words: $word_set"

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
if [[ "$safari_set" != "$word_set" ]]; then
  fail "step 2 — Safari foreground: error set changed (was '$word_set', now '$safari_set')"
fi
pass "step 2 — Safari foreground: $n_safari errors retained"

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
  if [[ "$slack_set" != "$word_set" ]]; then
    fail "step 3a — Slack foreground: error set changed"
  fi
  pass "step 3a — Slack foreground: $n_slack errors retained"
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
if [[ "$term_set" != "$word_set" ]]; then
  fail "step 3 — Terminal foreground: error set changed"
fi
pass "step 3 — Terminal foreground: $n_term errors retained"

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
if [[ "$word2_set" != "$word_set" ]]; then
  fail "step 4 — Back to Word: error set changed (was '$word_set', now '$word2_set')"
fi
pass "step 4 — Back to Word: $n_word2 errors intact"

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
