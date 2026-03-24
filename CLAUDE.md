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

## Running the app

The app is a GUI program. It can be started from the console:
```
./target/debug/acatts-rust &
```
It runs on port 3000 (HTTPS). The Word add-in connects to it.
Reload the add-in by typing space+backspace in Word — never ask the user to do it.
