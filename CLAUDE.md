# CLAUDE.md

## Debugging rules

When something doesn't work, the bug is ALWAYS in our code. Never blame:
- Timing ("actor too slow", "need longer wait")
- OS/framework ("macOS focus", "egui behavior", "Word API unstable")
- External systems ("port not ready", "stale state from previous run")

Trace the actual code path. Add targeted logging. Find the exact line.

## Working style

- Never stop to ask questions when the answer is obvious from context
- Make tiny changes, verify each one before the next
- When fixing test failures, test the specific failing case first, not the full suite
- When replacing code, DELETE the old path completely — no fallbacks, no duplicates
- Never rewrite existing working code — search and reuse
- Never rescan the whole document on every keystroke
- Never use timeouts as fixes
- Bash command timeouts: max 15 seconds, preferably 5 seconds. NEVER 10 minutes.

## Running the app

The app is a GUI program. Start from console:
```
./target/debug/acatts-rust &
```
It runs on port 3000 (HTTPS). The Word add-in connects to it.

## Reconnecting the add-in

Two scripts for reconnecting the Word add-in:

1. **`reconnect_addin.sh`** — Use when the add-in lost connection (shows "Prøv på nytt" button).
   Ensures the app is running, then clicks "Prøv på nytt" in the sidebar via AppleScript.
   ```
   ./reconnect_addin.sh
   ```

2. **`reload_addin.sh`** — Use when the add-in is loaded but needs a refresh.
   Right-clicks the taskpane and selects reload.
   ```
   ./reload_addin.sh
   ```

Always try `reconnect_addin.sh` first. Never ask the user to reload — use these scripts.
