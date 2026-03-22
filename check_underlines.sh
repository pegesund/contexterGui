#!/bin/bash
# Compare red underlines in Word document with /errors endpoint
ENDPOINT="https://127.0.0.1:3000/errors"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "=== Underline vs Errors Check ==="

# 1. Count errors from /errors endpoint
ERRORS_JSON=$(curl -sk "$ENDPOINT" 2>/dev/null)
ERROR_COUNT=$(echo "$ERRORS_JSON" | python3 -c "import json,sys; print(len(json.load(sys.stdin)))" 2>/dev/null || echo "0")
echo "Errors from /errors endpoint: $ERROR_COUNT"
echo "$ERRORS_JSON" | python3 -c "
import json,sys
for e in json.load(sys.stdin):
    print(f\"  [{e['category']}] {e['word'][:30]} | {e['rule']}\")" 2>/dev/null

# 2. Count underlined words in Word
echo ""
echo "Scanning Word for underlined words..."
RESULT=$(osascript "$SCRIPT_DIR/scan_underlines.applescript" 2>/dev/null)
echo "$RESULT"

# 3. Extract counts and compare
UNDERLINE_COUNT=$(echo "$RESULT" | head -1 | grep -o '^[0-9]*')
echo ""
echo "=== Comparison ==="
echo "  /errors endpoint: $ERROR_COUNT"
echo "  Word underlines:  ${UNDERLINE_COUNT:-0}"
if [ "${UNDERLINE_COUNT:-0}" -gt "$ERROR_COUNT" ] 2>/dev/null; then
    STALE=$((UNDERLINE_COUNT - ERROR_COUNT))
    echo "  MISMATCH: $STALE stale underlines in document!"
elif [ "${UNDERLINE_COUNT:-0}" -eq "$ERROR_COUNT" ] 2>/dev/null; then
    echo "  OK: counts match"
else
    echo "  Fewer underlines than errors (underlines may not have been applied yet)"
fi
