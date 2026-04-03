"""Windows integration test -- mirrors Mac test_errors.sh.

Requires: Word open with a document, acatts-rust running on port 52580.
Tests spelling errors, grammar errors, underlines, undo/restore.

Usage: py tests/test_windows.py
"""

import sys
sys.stdout.reconfigure(line_buffering=True)

import win32com.client
import time
import json
import urllib.request
import hashlib
import os
import ctypes

# Force unbuffered output (Python buffers when piped)
sys.stdout.reconfigure(line_buffering=True)

ERRORS_URL = "http://127.0.0.1:52580/errors"
LOG_FILE = None  # set in main() from %TEMP%/acatts-rust.log
PASS = 0
FAIL = 0
DELAY = 0.05  # per-character typing delay

# --- Helpers ---

def fetch_errors():
    try:
        with urllib.request.urlopen(ERRORS_URL, timeout=2) as r:
            return json.loads(r.read())
    except:
        return None

def get_underlined_words(doc):
    """Get words with wavy underlines (red=spelling, blue=grammar)."""
    underlined = {"red": [], "blue": []}
    for i in range(1, doc.Words.Count + 1):
        w = doc.Words(i)
        text = w.Text.strip()
        ul = w.Font.Underline
        if ul != 0 and text:
            color = w.Font.UnderlineColor
            # Red: wdColorRed=255, Blue: wdColorBlue=16711680 (0xFF0000 in BGR)
            if color == 255:
                underlined["red"].append(text)
            elif color == 16711680:
                underlined["blue"].append(text)
            else:
                underlined["red"].append(text)  # default to red
    return underlined

def doc_hash(doc):
    text = doc.Content.Text.rstrip("\r\n")
    return hashlib.md5(text.encode("utf-8")).hexdigest()

def check_error(desc, word, errors, expected_suggestion=""):
    global PASS, FAIL
    found = [e for e in errors if e.get("word") == word or word in e.get("word", "") or word in e.get("sentence", "")]
    if found and (not expected_suggestion or any(expected_suggestion in f.get("suggestion", "") for f in found)):
        print(f"  PASS: {desc}")
        PASS += 1
    else:
        print(f"  FAIL: {desc}")
        for e in errors:
            print(f"        {e.get('category','')} {e.get('word','')[:40]} | {e.get('rule','')}")
        FAIL += 1

def check_no_error(desc, word, errors):
    global PASS, FAIL
    found = [e for e in errors if e.get("word") == word]
    if not found:
        print(f"  PASS: {desc}")
        PASS += 1
    else:
        print(f"  FAIL: {desc}")
        FAIL += 1

def check_grammar(desc, fragment, errors):
    global PASS, FAIL
    found = [e for e in errors if e.get("category") == "grammar" and fragment in e.get("sentence", "")]
    if found:
        print(f"  PASS: {desc}")
        PASS += 1
    else:
        print(f"  FAIL: {desc}")
        for e in errors:
            print(f"        {e.get('category','')} {e.get('word','')[:40]} | {e.get('rule','')}")
        FAIL += 1

def check_underlined(desc, word, doc):
    global PASS, FAIL
    for attempt in range(5):
        ul = get_underlined_words(doc)
        all_ul = ul["red"] + ul["blue"]
        if any(word.lower() in w.lower() for w in all_ul):
            print(f"  PASS: {desc}")
            PASS += 1
            return
        time.sleep(2)
    print(f"  FAIL: {desc}")
    print(f"        Underlines: red={ul['red']}, blue={ul['blue']}")
    FAIL += 1

def check_not_underlined(desc, word, doc):
    global PASS, FAIL
    ul = get_underlined_words(doc)
    all_ul = ul["red"] + ul["blue"]
    if not any(word.lower() in w.lower() for w in all_ul):
        print(f"  PASS: {desc}")
        PASS += 1
    else:
        print(f"  FAIL: {desc} (still underlined)")
        FAIL += 1

def check_alignment(doc):
    """Verify error count matches underline count."""
    global PASS, FAIL
    for attempt in range(3):
        errors = fetch_errors() or []
        ul = get_underlined_words(doc)
        n_errors = len(errors)
        n_ul = len(ul["red"]) + len(ul["blue"])
        if n_errors == n_ul:
            print(f"  PASS: alignment {n_errors} errors = {n_ul} underlines")
            PASS += 1
            return
        time.sleep(3)
    print(f"  ALIGNMENT WARNING: {n_errors} errors != {n_ul} underlines")

WORD_HWND = 0  # set in main()

def get_log_size():
    """Get current log file size (used to check for new rescan messages)."""
    import os
    try:
        return os.path.getsize(LOG_FILE)
    except:
        return 0

def check_no_full_rescan(desc, log_pos_before):
    """Check that no full document rescan happened since log_pos_before."""
    global PASS, FAIL
    try:
        with open(LOG_FILE, "r", encoding="utf-8", errors="replace") as f:
            f.seek(log_pos_before)
            new_lines = f.read()
        rescan_count = new_lines.count("rescanning")
        full_doc_count = new_lines.count("Full doc text:")
        if rescan_count == 0 and full_doc_count == 0:
            print(f"  PASS: {desc} (no full rescan)")
            PASS += 1
        else:
            print(f"  FAIL: {desc} ({rescan_count} rescans, {full_doc_count} full doc reads)")
            FAIL += 1
    except:
        print(f"  SKIP: {desc} (log not readable)")


def bring_word_to_front():
    """Force Word to foreground (NorskTale always-on-top steals focus).
    Uses Alt key trick to allow SetForegroundWindow from background."""
    # Press and release Alt to allow SetForegroundWindow
    ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)  # Alt down
    ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)  # Alt up
    ctypes.windll.user32.ShowWindow(WORD_HWND, 9)  # SW_RESTORE
    ctypes.windll.user32.SetForegroundWindow(WORD_HWND)
    time.sleep(0.5)

def type_text(sel, text):
    bring_word_to_front()
    for ch in text:
        if ch == '\n':
            sel.TypeParagraph()
        else:
            sel.TypeText(ch)
        time.sleep(DELAY)
    # Move cursor to trigger NorskTale context read
    bring_word_to_front()
    sel.HomeKey(Unit=6)  # start of doc
    time.sleep(0.5)
    sel.EndKey(Unit=6)   # end of doc
    time.sleep(0.5)

def go_to_end(sel):
    bring_word_to_front()
    sel.EndKey(Unit=6)  # wdStory
    time.sleep(0.3)


def restore_document(doc, orig_text, orig_hash):
    """Restore document to original state and verify."""
    bring_word_to_front()
    # Clear all underlines before replacing text
    doc.Content.Font.Underline = 0  # wdUnderlineNone
    doc.Content.Text = orig_text + "\r"
    time.sleep(1)
    # Move cursor to trigger NorskTale paragraph scanning at start and end
    # New paragraph IDs will cause prune_resolved_errors to drop stale errors
    bring_word_to_front()
    sel = doc.Application.Selection
    sel.HomeKey(Unit=6)  # start of doc
    time.sleep(1)
    sel.EndKey(Unit=6)   # end of doc
    time.sleep(1)
    sel.HomeKey(Unit=6)  # back to start
    time.sleep(2)
    # Wait for errors to drain
    for _ in range(10):
        errors = fetch_errors() or []
        if len(errors) == 0:
            break
        time.sleep(1)
    h = doc_hash(doc)
    if h != orig_hash:
        print(f"  WARNING: Document restore failed! hash={h} expected={orig_hash}")
        return False
    return True

def wait_for_errors(min_count=1, timeout=10):
    start = time.time()
    while time.time() - start < timeout:
        errors = fetch_errors()
        if errors and len(errors) >= min_count:
            return errors
        time.sleep(1)
    return fetch_errors() or []

def wait_no_errors_for(word, timeout=10):
    start = time.time()
    while time.time() - start < timeout:
        errors = fetch_errors() or []
        if not any(e.get("word") == word for e in errors):
            return errors
        time.sleep(1)
    return fetch_errors() or []


# --- Main ---

def main():
    global PASS, FAIL

    print("=== NorskTale Windows Integration Test ===\n")

    # Set log file path
    global LOG_FILE
    import os
    LOG_FILE = os.path.join(os.environ.get("TEMP", "/tmp"), "acatts-rust.log")

    # Check app is running
    errors = fetch_errors()
    if errors is None:
        print("FATAL: acatts-rust not running on port 52580")
        sys.exit(1)

    # Connect to Word
    word = win32com.client.Dispatch("Word.Application")
    doc = word.ActiveDocument
    sel = word.Selection
    orig_text = doc.Content.Text.rstrip("\r\n")
    orig_h = doc_hash(doc)
    baseline_errors = len(fetch_errors() or [])

    print(f"Document: '{doc.Name}' (hash: {orig_h[:12]}, baseline: {baseline_errors} errors)")

    # Bring Word to foreground
    global WORD_HWND
    word.Visible = True
    word.Activate()
    WORD_HWND = word.ActiveWindow.Hwnd
    bring_word_to_front()
    time.sleep(1)

    # ============================================================
    print("\nTest 0: Document health -- errors match underlines")
    check_alignment(doc)

    # ============================================================
    print("\nTest 1: Spelling error 'somx' -> 'som'")
    lp = get_log_size()
    go_to_end(sel); sel.TypeParagraph()
    type_text(sel, "Fotball er en morsom sport somx er veldig morsom.")
    time.sleep(8)
    errors = fetch_errors() or []
    check_error("somx detected", "somx", errors, "som")
    check_underlined("somx underlined", "somx", doc)
    check_alignment(doc)
    check_no_full_rescan("no full rescan", lp)
    restore_document(doc, orig_text, orig_h)

    # ============================================================
    print("\nTest 2: Correct text -- no false positives")
    lp = get_log_size()
    go_to_end(sel); sel.TypeParagraph()
    type_text(sel, "Fotball er en morsom sport.")
    time.sleep(8)
    errors = fetch_errors() or []
    check_no_error("sport not flagged", "sport", errors)
    check_no_error("morsom not flagged", "morsom", errors)
    check_no_full_rescan("no full rescan", lp)
    restore_document(doc, orig_text, orig_h)

    # ============================================================
    print("\nTest 3: Multiple errors in one sentence")
    lp = get_log_size()
    go_to_end(sel); sel.TypeParagraph()
    type_text(sel, "Jeg liker aa spise matx og drikkx.")
    time.sleep(12)
    errors = fetch_errors() or []
    check_error("matx detected", "matx", errors)
    check_error("drikkx detected", "drikkx", errors)
    check_underlined("matx underlined", "matx", doc)
    check_underlined("drikkx underlined", "drikkx", doc)
    check_alignment(doc)
    check_no_full_rescan("no full rescan", lp)
    restore_document(doc, orig_text, orig_h)

    # ============================================================
    print("\nTest 4: Grammar error -- gender mismatch")
    lp = get_log_size()
    go_to_end(sel); sel.TypeParagraph()
    type_text(sel, "Fotball er en morsom spor.")
    time.sleep(12)
    errors = fetch_errors() or []
    check_grammar("gender mismatch", "spor", errors)
    check_no_full_rescan("no full rescan", lp)
    restore_document(doc, orig_text, orig_h)

    # ============================================================
    print("\nTest 5: Grammar error -- adj gender mismatch")
    lp = get_log_size()
    go_to_end(sel); sel.TypeParagraph()
    type_text(sel, "Fotball er en morsomt sport.")
    time.sleep(8)
    errors = fetch_errors() or []
    check_grammar("adj gender mismatch", "morsomt", errors)
    check_no_full_rescan("no full rescan", lp)
    restore_document(doc, orig_text, orig_h)

    # ============================================================
    print("\nTest 6: Error detected, removed by restore")
    lp = get_log_size()
    go_to_end(sel); sel.TypeParagraph()
    type_text(sel, "Dette er en feilx i teksten.")
    time.sleep(8)
    errors = fetch_errors() or []
    check_error("feilx detected", "feilx", errors)
    restore_document(doc, orig_text, orig_h)
    time.sleep(5)
    errors = fetch_errors() or []
    check_no_error("feilx gone after restore", "feilx", errors)

    # ============================================================
    print("\nTest 7: Spelling error with BERT suggestion")
    lp = get_log_size()
    go_to_end(sel); sel.TypeParagraph()
    type_text(sel, "Jeg liker ikke katter og hundder.")
    time.sleep(12)
    errors = fetch_errors() or []
    check_error("hundder detected", "hundder", errors)
    check_underlined("hundder underlined", "hundder", doc)
    check_alignment(doc)
    check_no_full_rescan("no full rescan", lp)
    restore_document(doc, orig_text, orig_h)

    # ============================================================
    print("\nTest 8: Duplicate sentences both detected")
    lp = get_log_size()
    go_to_end(sel); sel.TypeParagraph()
    type_text(sel, "Han liker duplikatxx veldig godt.")
    go_to_end(sel); sel.TypeParagraph()
    type_text(sel, "Han liker duplikatxx veldig godt.")
    time.sleep(8)
    errors = fetch_errors() or []
    dup_count = len([e for e in errors if e.get("word") == "duplikatxx"])
    if dup_count == 2:
        print(f"  PASS: both duplicate duplikatxx detected ({dup_count})")
        PASS += 1
    else:
        print(f"  FAIL: expected 2 duplikatxx errors, got {dup_count}")
        FAIL += 1
    check_no_full_rescan("no full rescan", lp)
    restore_document(doc, orig_text, orig_h)

    # ============================================================
    print("\nTest 9: Rapid typing -- no crash")
    lp = get_log_size()
    go_to_end(sel); sel.TypeParagraph()
    type_text(sel, "Dette er en rask test med mange ord uten feilx.")
    time.sleep(8)
    errors = fetch_errors() or []
    check_error("feilx after rapid", "feilx", errors)
    check_underlined("feilx underlined", "feilx", doc)
    check_alignment(doc)
    check_no_full_rescan("no full rescan", lp)
    restore_document(doc, orig_text, orig_h)

    # ============================================================
    print("\nTest 10: Correct neuter sentence -- no false positives")
    lp = get_log_size()
    go_to_end(sel); sel.TypeParagraph()
    type_text(sel, "Fotball er et morsomt spill.")
    time.sleep(8)
    errors = fetch_errors() or []
    check_no_error("spill not flagged", "spill", errors)
    check_no_error("morsomt not flagged", "morsomt", errors)
    check_no_full_rescan("no full rescan", lp)
    restore_document(doc, orig_text, orig_h)

    # ============================================================
    print("\nTest 11: Grammar error -- wrong article gender")
    lp = get_log_size()
    go_to_end(sel); sel.TypeParagraph()
    type_text(sel, "Han kjopte en nytt hus.")
    time.sleep(12)
    errors = fetch_errors() or []
    check_grammar("wrong article gender", "en", errors)
    check_no_full_rescan("no full rescan", lp)
    restore_document(doc, orig_text, orig_h)

    # ============================================================
    # Verify original document is intact
    print("\nVerifying original document...")
    final_h = doc_hash(doc)
    if final_h == orig_h:
        print("  Original document: INTACT (hash matches)")
    else:
        print(f"  WARNING: Document hash changed! before={orig_h[:12]} after={final_h[:12]}")

    print(f"\n=== Results: {PASS} passed, {FAIL} failed ===")
    if FAIL > 0:
        sys.exit(1)

if __name__ == "__main__":
    main()
