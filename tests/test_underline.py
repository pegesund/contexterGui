"""Test: verify underlines appear in Word document after NorskTale detects errors.

Requires: Word open with a document containing errors, acatts-rust running.

Steps:
1. Activate Word and move cursor
2. Wait for NorskTale to process
3. Check underlines in document
4. Check /errors endpoint
"""

import win32com.client
import time
import json
import urllib.request

ERRORS_URL = "http://127.0.0.1:52580/errors"

def fetch_errors():
    try:
        with urllib.request.urlopen(ERRORS_URL, timeout=2) as r:
            return json.loads(r.read())
    except:
        return None

def get_underlined_words(doc):
    underlined = []
    for i in range(1, doc.Words.Count + 1):
        w = doc.Words(i)
        text = w.Text.strip()
        ul = w.Font.Underline
        if ul != 0 and text:
            underlined.append((text, ul, w.Font.UnderlineColor))
    return underlined

def main():
    print("=== Underline Test ===\n")

    # Connect to Word
    word = win32com.client.Dispatch("Word.Application")
    doc = word.ActiveDocument
    print(f"Document: {doc.Name} ({doc.Content.Text[:50].strip()}...)")

    # Step 1: Activate Word and position cursor
    print("\n1. Activating Word...")
    word.Visible = True
    word.Activate()
    word.WindowState = 1  # wdWindowStateNormal
    # Force Word to foreground using Windows API (AlwaysOnTop NorskTale steals focus)
    import ctypes
    hwnd = word.ActiveWindow.Hwnd
    ctypes.windll.user32.SetForegroundWindow(hwnd)
    time.sleep(1)
    # Click into document area to ensure caret is active
    word.Activate()
    ctypes.windll.user32.SetForegroundWindow(hwnd)
    time.sleep(1)

    # Move cursor around to trigger context read
    sel = word.Selection
    sel.HomeKey(Unit=6)  # wdStory — start
    time.sleep(1)
    sel.EndKey(Unit=6)   # wdStory — end
    time.sleep(1)

    # Type a misspelled word at the end
    print("   Typing 'fiskk' at end of document...")
    sel.TypeParagraph()
    time.sleep(0.3)
    for ch in "Jeg liker fiskk.":
        sel.TypeText(ch)
        time.sleep(0.05)
    time.sleep(1)

    # Move cursor away (so word boundary triggers spelling check)
    sel.HomeKey(Unit=6)
    time.sleep(2)

    # Step 2: Wait for NorskTale to process
    print("\n2. Waiting for error detection...")
    for attempt in range(20):
        errors = fetch_errors()
        if errors is None:
            print("   NorskTale not running!")
            return
        if len(errors) > 0:
            print(f"   {len(errors)} error(s) detected after {attempt + 1}s")
            break
        time.sleep(1)
    else:
        print("   TIMEOUT: no errors detected after 20s")
        errors = []

    print(f"   Errors: {json.dumps(errors, indent=2, ensure_ascii=False)[:500]}")

    # Step 3: Check underlines
    print("\n3. Checking underlines in Word...")
    time.sleep(2)  # Give time for underline sync
    underlined = get_underlined_words(doc)
    print(f"   Underlined words: {len(underlined)}")
    for text, ul, color in underlined:
        ul_name = {11: "wavy", 1: "single", 0: "none"}.get(ul, f"type={ul}")
        print(f"   '{text}' — {ul_name} color={color}")

    # Step 4: Results
    print("\n=== Results ===")
    n_errors = len(errors)
    n_underlines = len(underlined)

    if n_errors > 0 and n_underlines > 0:
        print(f"PASS: {n_errors} errors detected, {n_underlines} underlines in Word")
    elif n_errors > 0 and n_underlines == 0:
        print(f"FAIL: {n_errors} errors detected but 0 underlines in Word")
    elif n_errors == 0:
        print(f"FAIL: 0 errors detected")

    # Cleanup: undo the typed text
    print("\nCleaning up...")
    doc.Undo(3)  # undo TypeText + TypeParagraph + maybe more
    print("Done.")

if __name__ == "__main__":
    main()
