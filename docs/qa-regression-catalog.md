# Spell QA Regression Catalog

This catalog turns confirmed QA cases into stable acceptance criteria. It is a
test index, not a replacement for product logic. Runtime behavior remains owned
by the source modules named in each section.

## Coverage levels

- **Unit**: deterministic Rust tests with no desktop application running.
- **Bridge**: tests the contract between `BridgeManager` and a text bridge.
- **GUI**: exercises a real external application on the named operating system.
- **Manual**: required until the row has a reliable automation harness.

Before release, run the unit suite on both Windows and macOS, then execute the
GUI rows for each supported platform. A row marked `Open` is not a release claim.

## Deterministic unit catalog

| ID | Area | Platforms | Acceptance criterion | Owner | Coverage |
|---|---|---|---|---|---|
| UNIT-LANG-001 | Writing-language isolation | All | English does not reuse Norwegian dictionaries; Bokmal and Nynorsk do not reuse English dictionaries after repeated language switches. | `src/main.rs`: `cross_language_codes`, `load_cross_language_analyzers` | Rust tests in `cross_language_barrier_tests` |
| UNIT-SPELL-001 | Case preservation | All | A correction preserves normal initial capitalization and handles all-uppercase input without producing mixed-case output. | `src/main.rs`: `apply_original_initial_case` | Rust tests in `cross_language_barrier_tests` |
| UNIT-SPELL-002 | Duplicate occurrences | All | Two occurrences may produce two location-specific rows; one occurrence must not produce duplicate rows. | `src/main.rs`: writing-error identity and occurrence tracking | Partial Rust coverage; GUI verification required |
| UNIT-SPELL-003 | Norwegian characters | All | With Bokmal selected, `kjore -> kjore` with `o` replaced by `o-slash`, `grot -> grot` with `o` replaced by `o-slash`, `pa -> pa` with `a-ring`, `ar -> ar` with `a-ring`, and `aar` with an initial `a-ring` ranks plain `ar` with `a-ring` above unrelated alternatives. | `src/spelling_scorer.rs`: shared spelling ranking; `src/main.rs`: spelling pipeline | Rust spelling tests; retain exact QA sentences as fixtures |
| UNIT-SPELL-004 | Corrupted characters | All | `fiont` with `o-slash -> fint` and `tael` with `ae -> til` are evaluated in sentence context. | `src/spelling_scorer.rs`; `src/main.rs`: context spelling checks | Rust spelling tests; model-dependent top rank remains a scored assertion |
| UNIT-SPELL-005 | Token forms | All | Supported hyphenated words remain single tokens; email, URL, abbreviation, slash, plus, hash, underscore, and alphanumeric forms are either accepted or deliberately excluded without fragment errors. | `src/main.rs`: token cleanup; language analyzers | Open fixture matrix |
| UNIT-UI-001 | Error-row visibility | All | An error without a correction still has visible explanatory text; it must not render as an empty row of buttons. | `src/main.rs`: `writing_error_has_visible_panel_text` | Rust predicate coverage plus GUI row |

Use ASCII descriptions in source fixtures when practical, but keep the real
Norwegian characters in test input files so encoding behavior is exercised.

## Bridge integration catalog

| ID | Bridge | Platforms | Acceptance criterion | Owner | Coverage |
|---|---|---|---|---|---|
| BRIDGE-BROWSER-001 | Browser extension | Windows, macOS | `Some(empty context)` from an initialized empty editor replaces stale context. `None` means no payload and may use the existing fallback policy. | `src/main.rs`: `BridgeManager::try_browser_context`; `src/bridge/browser.rs`: `BrowserBridge::read_context` | `bridge_manager_tests::empty_browser_context_replaces_stale_context` |
| BRIDGE-BROWSER-002 | Browser full document | Windows, macOS | Deleting all text updates the full document to empty and clears active browser errors on the next poll. | `src/bridge/browser.rs`: `read_full_document`; `src/main.rs`: `try_update_doc_text`, `prune_resolved_errors` | Route test plus GUI verification |
| BRIDGE-UIA-001 | Windows Accessibility | Windows | An empty focused Edit/Document control is authoritative; empty descendants are rejected so a background Notepad tab cannot replace the active empty tab. | `src/bridge/accessibility_win.rs`: `AccessibilityBridge::accept_text_element` | Three tests in `accessibility_win::tests` |
| BRIDGE-AX-001 | macOS Accessibility | macOS | An initialized empty AX editor is authoritative and replaces stale context for Notes, TextEdit, Slack, and other AX writing surfaces. | `src/main.rs`: `BridgeManager::try_macos_ax_context`; `src/bridge/ax_mac.rs`: `AxMacBridge::read_context` | Regression test required before fix |
| BRIDGE-SCOPE-001 | Active bridge | All | The pencil panel, Tips count, `/errors`, and correction actions expose only errors owned by the active bridge. | `src/main.rs`: `paragraph_id_matches_bridge`, `error_paragraph_matches_bridge` | Rust predicate coverage plus focus-switch GUI tests |
| BRIDGE-SWITCH-001 | App switching | All | Word errors recover when returning to Word, while Word rows are not rendered over Slack, browser editors, Notes, or Notepad. | `src/main.rs`: `ContextApp::update` app-switch block | macOS `scripts/test-focus-errors.sh`; Windows manual gap |
| BRIDGE-NOTEPAD-001 | Notepad tabs | Windows | The active tab owns reads and corrections. A misspelling in a background tab must not appear or be corrected from the foreground tab. | `src/bridge/accessibility_win.rs`; `src/main.rs`: `BridgeManager::read_context_windows` | UIA unit policy plus Windows GUI row |
| BRIDGE-EDIT-001 | In-place edit | All | Editing an existing word is rechecked without requiring an extra trailing space after the edit debounce. | `src/main.rs`: context polling and spelling queue | Word automation coverage; non-Word GUI gap |
| BRIDGE-DELETE-001 | Delete/cut | All | Corrected, cut, or deleted text removes its error rows and Tips count within one polling cycle. | `src/main.rs`: `try_update_doc_text`, `prune_resolved_errors` | Word automation; browser/UIA/AX rows below |
| BRIDGE-NAV-001 | Show in document | Word | Navigation selects the exact error through the bridge that owns its paragraph. The action is hidden for bridges without navigation support. | `src/main.rs`: `BridgeManager::can_navigate_to_error`, `select_word_in_paragraph` | GUI verification |

### Verified bridge contracts

`BrowserBridge::read_context` returns an initialized empty context:

```rust
if text.is_empty() {
    return Some(CursorContext {
        cursor_doc_offset: Some(0),
        ..Default::default()
    });
}
```

`AccessibilityBridge::accept_text_element` accepts empty text only for the
focused text-control path:

```rust
let (raw, doc) = Self::try_read_raw(&element, allow_empty)?;
let is_text_control = Self::is_text_control(&element);
if !Self::should_accept_document(&doc, allow_empty, is_text_control) {
    return None;
}
```

`ContextApp::prune_resolved_errors` owns stale-error disposal after document
state has been refreshed:

```rust
if doc_text.trim().is_empty() {
    self.writing_errors.clear();
    self.grammar_queue.clear();
    self.spelling_queue.clear();
    return;
}
```

## Windows GUI catalog

| ID | Application | Acceptance criterion | Automation |
|---|---|---|---|
| WIN-WORD-001 | Word | Type, paste, cut, move, and delete 8-10 lines. Errors follow current text, deleted rows disappear, and unchanged paragraphs are not fully rescanned. | `tests/test_windows.py` plus manual bulk move |
| WIN-WORD-002 | Word | A first click on correction, completion, toolbar, question, and navigation actions executes once without a focus-only click. | Manual |
| WIN-DOCS-001 | Google Docs | Type a misspelling, verify one row, then select all and delete. The row and Tips count clear within one poll. | Manual; browser harness open |
| WIN-NOTEPAD-001 | Notepad | Put different misspellings in two tabs. Only the active tab is visible/actionable; clearing it does not reveal the background tab until that tab is selected. | Manual; UIA harness open |
| WIN-OVERLAY-001 | Word, Docs, Slack, Notepad | Spell stays above the writing app, preserves its caret-relative position during Spell clicks, and can be manually minimized. | Manual |
| WIN-OCR-001 | Snipping Tool | Screenshot prompt appears immediately for English, Bokmal, and Nynorsk; one capture creates one prompt; prompt remains above Spell. | Manual |
| WIN-DOWNLOAD-001 | Initial setup | Slow/interrupted downloads retry or resume; Close cancels active setup instead of continuing in the background. | Manual packaging test |
| WIN-UPDATE-001 | Installed release | Update applies without a raw process/DLL dialog; update toast is fully visible and localized. | Manual installed-build test |
| WIN-IDENTITY-001 | Task Manager/uninstall | Main process is `Spell`; uninstall terminates the native bridge and removes installed runtime files without touching user data unless clean removal is selected. | Manual installed-build test |

## macOS GUI catalog

| ID | Application | Acceptance criterion | Automation |
|---|---|---|---|
| MAC-WORD-001 | Word | Existing errors survive Word -> other app -> Word; correction and navigation work on the first click. | `test_errors.sh`, `scripts/test-focus-errors.sh`, manual click check |
| MAC-AX-001 | Notes, TextEdit, Slack | Type a misspelling and delete all text. Error rows and Tips clear within one poll. | Open until AX empty-context route test and GUI harness land |
| MAC-BROWSER-001 | Google Docs | Browser errors do not inherit Word/AX state; deleting all editor text clears browser rows. | Manual |
| MAC-OVERLAY-001 | Word, Docs, Notes, TextEdit, Slack | Clicking document text must not OS-minimize Spell. Intentional toolbar-only collapse is allowed only when the selected panel has no renderable content. | Open: requires screenshot/video plus `spell.log` to distinguish OS minimize from content-driven resize |
| MAC-OVERLAY-002 | All writing apps | A first click on correction, completion, and toolbar actions executes once and does not move the overlay to an unrelated screen corner. | Manual |
| MAC-STT-001 | Word/other editor | First mic use downloads the selected-language Whisper model if absent; recording transcribes speech or reports a localized actionable error. | Manual installed-build test |
| MAC-WORDADDIN-001 | Word | Installed signed app can install/load the add-in; an unsigned terminal build may fail certificate authorization and must explain the limitation. | Manual signed-build test |

## Cross-platform product catalog

| ID | Area | Acceptance criterion | Owner/automation |
|---|---|---|---|
| PRODUCT-I18N-001 | Localization | App language controls UI strings only; writing language controls spelling, grammar, completion, STT model, voices, and downloads. | Language traits and settings UI; manual screen pass |
| PRODUCT-TTS-001 | Voice | Selected voice remains stable word-to-word and is constrained to the active writing language. | TTS selection code; manual audio pass |
| PRODUCT-STT-001 | STT language | STT uses the active writing-language model; Improve does not silently reinterpret a transcript as another language. | Mic/Whisper flow; manual audio fixtures |
| PRODUCT-PANEL-001 | Pencil rows | Error word, correction/explanation, audio actions, dictionary action, and supported navigation remain visible and clickable. | `src/main.rs` panel renderer; GUI pass |
| PRODUCT-LAYOUT-001 | Borders/icons | Spell icon, process identity, window border, right edge, and taskbar/Dock identity remain consistent. | Main viewport/package metadata; visual pass |

## Release gates

1. Run `cargo test --release --bin spell` on Windows and macOS.
2. Run `py tests/test_windows.py` against a running Windows build and Word.
3. Run `./test_errors.sh` and `./scripts/test-focus-errors.sh` on macOS.
4. Execute all GUI rows affected by the changed owner module.
5. Record OS version, Spell commit/tag, writing language, active bridge from
   Settings, and the relevant log for every failure.

The current GUI scripts are app-assisted integration tests, not hermetic unit
tests. A passing Rust suite therefore does not replace the Windows and macOS
GUI gates.
