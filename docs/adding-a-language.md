# Adding a new language

How to add Nynorsk, English, German or anything else to the NorskTale stack.

## Background

The original NorskTale build had Norwegian Bokmål constants (paths, modal verb
pairs, function words, voice codes, UI strings) scattered across the codebase.
The `multiLanguage` branches refactored every constant behind a trait stack
in the `language-rs` crate so a new language only has to implement the seven
sub-traits — no surgery in the host modules.

The runtime CLI flag `--language nb` (default) selects the language at
startup. Resolution happens once in `main()` and the resulting
`Arc<dyn LanguageBundle>` flows through `ContextApp` to every site that
needs language data.

## Where the trait lives

- Crate: `rustSpell/language-rs` (`pegesund/rustSpell`, `multiLanguage` branch)
- Lib name: `language` (so callers write `language::LanguageProfile`)
- Single source file: `src/lib.rs`
- Bokmål impl: `BokmalLanguage` (a zero-sized struct)
- Registry: `language::resolve_language(code)` returns
  `Result<Arc<dyn LanguageBundle>, UnsupportedLanguage>`

## The trait stack

`LanguageBundle` is a convenience supertrait combining everything; you don't
implement it directly because of the blanket impl.

| Sub-trait | Methods | What it carries |
|---|---|---|
| `LanguageProfile` | `code`, `display_name`, `data_root` | identity |
| `LanguageLexicon` | `mtag_fst_path`, `wordfreq_path`, `wordfreq_common_threshold` | dictionary + frequency corpus |
| `LanguageGrammar` | `prolog_rules_path` | SWI-Prolog grammar rules file |
| `LanguageMlm` | `onnx_path`, `tokenizer_path` | masked-LM (BERT) model + tokenizer |
| `LanguageSpelling` | `modal_confusion_pairs`, `function_words`, `binding_letters`, `binding_s_suffixes`, `free_vowel_swaps` | spell-checker rules + heuristics |
| `LanguageVoice` | `tts_default_voice`, `tts_voice_filters`, `stt_language_code`, `ocr_language_code` | TTS/STT/OCR codes |
| `LanguageUi` | 17 methods (toolbar tooltips, status text, error messages, format-arg variants) | every user-facing string |

All seven sub-traits share `Send + Sync + 'static` bounds (inherited from
`LanguageProfile`) so the implementation is automatically usable behind
`Arc<dyn LanguageBundle>` and across thread boundaries.

## Step-by-step: adding `NynorskLanguage`

The example assumes Norwegian Nynorsk (`nn`). Adapt names for other languages.

### 1. Branch each affected repo

You'll touch `rustSpell` (the trait + impl) and `contexterGui` (registry use
site if you want non-Bokmål to be default-discoverable). Make a branch:

```bash
cd /Users/pegesund/dev/dyslex/rustSpell
git checkout -b add-nynorsk multiLanguage
```

### 2. Define the struct

In `rustSpell/language-rs/src/lib.rs`:

```rust
#[derive(Debug, Default, Clone, Copy)]
pub struct NynorskLanguage;
```

The struct is intentionally zero-sized — language identity is the type, all
data lives behind the trait methods.

### 3. Implement the seven sub-traits

```rust
impl LanguageProfile for NynorskLanguage {
    fn code(&self) -> &'static str { "nn" }
    fn display_name(&self) -> &'static str { "Nynorsk" }
    fn data_root(&self) -> PathBuf { PathBuf::from("nn") }
}

impl LanguageLexicon for NynorskLanguage {
    fn mtag_fst_path(&self) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../mtag-rs/data/fullform_nn.mfst")
    }
    fn wordfreq_path(&self) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../contexter-repo/training-data/wordfreq_nn.tsv")
    }
    fn wordfreq_common_threshold(&self) -> u64 {
        // Calibrate against the new language's corpus.
        40_000
    }
}

impl LanguageGrammar for NynorskLanguage {
    fn prolog_rules_path(&self) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../syntaxer/grammar_rules_nn.pl")
    }
}

impl LanguageMlm for NynorskLanguage {
    fn onnx_path(&self) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../contexter-repo/training-data/onnx/norbert4_nn_int8.onnx")
    }
    fn tokenizer_path(&self) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../contexter-repo/training-data/onnx/tokenizer_nn.json")
    }
}

impl LanguageSpelling for NynorskLanguage {
    fn modal_confusion_pairs(&self) -> &'static [(&'static str, &'static str)] {
        // Linguist call: which BERT-confusable misspellings exist in this language?
        &[]
    }
    fn function_words(&self) -> &'static [&'static str] {
        // Prepositions, conjunctions, articles. Nynorsk shares many with Bokmål.
        &["til", "frå", "med", /* ... */ ]
    }
    fn binding_letters(&self) -> &'static [u8] { b"se" }
    fn binding_s_suffixes(&self) -> &'static [&'static [u8]] {
        &[b"ing", b"het", /* ... */ ]
    }
    fn free_vowel_swaps(&self) -> &'static [(u8, u8)] {
        // Same Norwegian vowel pairs as Bokmål.
        &[(b'e', 0xB8), (b'o', 0xA5), (b'a', 0xA6)]
    }
}

impl LanguageVoice for NynorskLanguage {
    fn tts_default_voice(&self) -> &'static str { "Eir" }       // Nynorsk macOS voice
    fn tts_voice_filters(&self) -> &'static [&'static str] {
        &["nn_NO", "no_NO"]
    }
    fn stt_language_code(&self) -> &'static str { "no" }        // Whisper has no nn
    fn ocr_language_code(&self) -> &'static str { "nn" }        // Windows OCR
}

impl LanguageUi for NynorskLanguage {
    fn ui_pin_cursor_on(&self)  -> &'static str { "Følg markøren (på)" }   // Nynorsk
    fn ui_pin_cursor_off(&self) -> &'static str { "Følg markøren (av)" }
    fn ui_settings(&self)       -> &'static str { "Innstillingar" }
    fn ui_minimize(&self)       -> &'static str { "Minimer" }
    fn ui_no_errors(&self)      -> &'static str { "Ingen feil funne." }
    fn ui_ai_fix_all(&self)     -> &'static str { "✨ AI-rett alle" }
    fn ui_voice_download_help(&self) -> &'static str {
        "Last ned stemmer i Systeminnstillingar > Tilgjenge > Opplest innhald > Systemstemme"
    }
    fn ui_word_not_in_dict(&self, word: &str) -> String {
        format!("«{}» finst ikkje i ordboka.", word)
    }
    fn ui_removed_aa_before(&self, word: &str) -> String {
        format!("Fjerna «å» framfor «{}».", word)
    }
    fn ui_removed_word(&self, word: &str) -> String {
        format!("Fjerna «{}».", word)
    }
    fn ui_ai_correcting_seconds(&self, secs: u64) -> String {
        format!("AI rettar... ({}s)", secs)
    }
    fn ui_no_audio_captured(&self)      -> &'static str { "(ingen lyd fanga)" }
    fn ui_no_speech_recognized(&self)   -> &'static str { "(inga tale attkjend)" }
    fn ui_whisper_model_load_failed(&self) -> &'static str {
        "kunne ikkje laste Whisper-modellen"
    }
    fn ui_whisper_dll_load_failed(&self, error: &str) -> String {
        format!("kunne ikkje laste whisper.dll: {}", error)
    }
    fn ui_ocr_no_text(&self)            -> &'static str { "Ingen tekst funnen i biletet" }
    fn ui_ocr_lang_pack_missing(&self)  -> &'static str {
        "Norwegian Nynorsk OCR language pack not installed."
    }
}
```

You don't have to implement `LanguageBundle` — there's a blanket
`impl<T: ...> LanguageBundle for T` that applies as soon as all seven
sub-traits are implemented.

### 4. Register in `resolve_language`

Same file, near the bottom:

```rust
pub fn resolve_language(code: &str) -> Result<Arc<dyn LanguageBundle>, UnsupportedLanguage> {
    match code {
        "nb" | "no" => Ok(Arc::new(BokmalLanguage)),
        "nn"        => Ok(Arc::new(NynorskLanguage)),    // <-- new arm
        other       => Err(UnsupportedLanguage(other.to_string())),
    }
}
```

### 5. Add tests

For each implementation, add a unit test that asserts the file exists where
the trait method points (so a typo in a path is caught at `cargo test` time,
not at app startup):

```rust
#[test]
fn nynorsk_mtag_fst_exists() {
    let path = NynorskLanguage.mtag_fst_path();
    assert!(path.exists(), "Nynorsk mtag FST not found at {}", path.display());
}

#[test]
fn nynorsk_prolog_rules_exists() {
    let path = NynorskLanguage.prolog_rules_path();
    assert!(path.exists(), "Nynorsk Prolog rules not found at {}", path.display());
}

// ... and so on for wordfreq, onnx, tokenizer
```

```bash
cargo test -p language-rs
```

### 6. Drop in the data files

The trait methods above expect specific file names. Create them:

| Trait method | Expected on disk |
|---|---|
| `mtag_fst_path` | `rustSpell/mtag-rs/data/fullform_nn.mfst` |
| `wordfreq_path` | `contexter-repo/training-data/wordfreq_nn.tsv` |
| `prolog_rules_path` | `syntaxer/grammar_rules_nn.pl` |
| `onnx_path` | `contexter-repo/training-data/onnx/norbert4_nn_int8.onnx` |
| `tokenizer_path` | `contexter-repo/training-data/onnx/tokenizer_nn.json` |

The data files themselves are out of scope for this guide — building a
mtag FST, training a NorBERT4 variant for the new language, and porting
the Prolog grammar rules are each multi-week projects.

If a data file isn't ready yet, point the trait at the Bokmål file as a
temporary fallback so the rest of the stack can be tested:

```rust
fn onnx_path(&self) -> PathBuf {
    BokmalLanguage.onnx_path() // TODO: replace with Nynorsk model when trained
}
```

### 7. Verify end-to-end

```bash
cd /Users/pegesund/dev/dyslex/contexterGui
cargo check --bin acatts-rust
./target/debug/acatts-rust --language nn
```

The startup banner should print `Language: Nynorsk (nn)` and the dictionary
should load. If a trait method points at a missing file, the app crashes
with the file path in the error message — fix the path or drop in the file.

### 8. Commit + PR

Per the project's small-step convention, separate commits per concern:

```bash
git add language-rs/src/lib.rs
git commit -m "Add NynorskLanguage trait implementations"
git push origin add-nynorsk
gh pr create --base multiLanguage --title "Add Nynorsk language support"
```

## Known limitations

The trait covers every place `ContextApp` reads a language constant, but a
few helper modules still hard-code Bokmål via a module-level `BOKMAL`
constant. They're already on the trait API; they just don't yet receive a
runtime language instance because their host structs/functions don't take
an `Arc<dyn LanguageBundle>` parameter:

- `compound_walker.rs` (FST hot path)
- `spelling_scorer.rs::try_split_function_word`
- `tts/macos_impl.rs` (`MacTtsEngine`)
- `stt/mod.rs` and `stt/windows_impl.rs` (`WhisperEngine`)
- `ocr.rs` (`OcrClipboard` impls for both macOS and Windows)
- `score_word_for_completions` free function in `main.rs` (top of file)

When you select `--language nn` today, those modules still use the Bokmål
constants. Bokmål itself is unaffected. To finish the multi-language story
those host modules need to grow an `Arc<dyn LanguageBundle>` parameter
too — a separate refactor commit per module.

The legacy free functions `dict_path()`, `grammar_rules_path()`,
`compound_data_path()`, `syntaxer_dir()` in `main.rs` are also still
defined; only dev/test bins (`test_spelling`, `test_fotball`, etc.)
and the dead `load_swipl_checker` helper still call them. They can be
removed in a sweep once the dev bins are migrated.

## File layout convention (future)

When more than one language is real, the convention is per-language subdirs:

```
rustSpell/mtag-rs/data/
    nb/fullform.mfst
    nn/fullform.mfst
    en/fullform.mfst

syntaxer/
    nb/grammar_rules.pl
    nn/grammar_rules.pl
    en/grammar_rules.pl
```

Today the Bokmål files live at the historical paths (`fullform_bm.mfst`,
`syntaxer/grammar_rules.pl`) and the trait points at them directly. When you
move them into `nb/`, leave a symlink or shim at the old path for the dev
bins that still hard-code the old location.

## Where to look in the source

| Concern | File |
|---|---|
| Trait definitions | `rustSpell/language-rs/src/lib.rs` |
| Bokmål impl | same file, `impl LanguageProfile for BokmalLanguage` and friends |
| Registry | same file, `pub fn resolve_language` near the bottom |
| `--language` flag parsing | `contexterGui/src/main.rs::main()` |
| `ContextApp` field | `contexterGui/src/main.rs`, `struct ContextApp { language: Arc<dyn ...>, ... }` |
| Live trait call sites | grep `self.language.` in `contexterGui/src/main.rs` |
