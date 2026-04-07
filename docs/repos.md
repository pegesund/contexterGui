# NorskTale Repository Overview

NorskTale is a Norwegian spelling and grammar checker for dyslexia users, built across several repositories.

## Repositories

### acatts-rust (this repo)
- **GitHub**: pegesund/contexterGui
- **Path**: `NorskTale/acatts-rust`
- **Purpose**: GUI application (eframe/egui), bridges to Word/Google Docs/accessibility, spelling/grammar error display, word completion UI
- **Key files**: `src/main.rs` (app logic), `src/bert_worker.rs` (BERT inference thread), `src/bridge/` (Word COM, browser, accessibility bridges), `extension/` (Chrome extension for Google Docs)
- **Dependencies**: nostos-cognio (NLP engine), mtag-rs (dictionary)

### nostos-cognio
- **GitHub**: pegesund/contexter (subdir `nostos-cognio/`)
- **Path**: `contexter-repo/nostos-cognio`
- **Purpose**: NLP engine — NorBERT4 masked language model inference, word completion (`complete_word`), spelling suggestions, sentence boundary detection, SWI-Prolog grammar checker integration
- **Key files**: `src/model.rs` (ONNX/OpenVINO inference with auto-detection), `src/complete.rs` (word completion), `src/grammar/swipl_checker.rs` (SWI-Prolog bridge)
- **Models**: `training-data/onnx/norbert4_base_int8.onnx` (ORT INT8), `training-data/openvino_onnx/norbert4_patched_int8.*` (OpenVINO INT8, 2x faster on Intel)

### syntaxer
- **GitHub**: pegesund/syntaxer
- **Path**: `syntaxer`
- **Purpose**: Norwegian grammar rules in Prolog
- **Key files**: `grammar_rules.pl` (all grammar rules), `sentence_split.pl` (sentence boundary splitting), `compound_data.pl` (compound word suggestions)
- **Note**: SWI-Prolog reads these files directly — live editing without recompilation

### mtag-rs
- **GitHub**: pegesund/rustSpell (subdir `mtag-rs/`)
- **Path**: `rustSpell/mtag-rs`
- **Purpose**: Morphological analyzer — FST-based dictionary with 633K Norwegian word forms
- **Key files**: `data/fullform_bm.mfst` (dictionary), `src/lib.rs` (analyze, prefix_lookup, has_word, fuzzy_lookup)

### language-rs
- **GitHub**: pegesund/rustSpell (subdir `language-rs/`)
- **Path**: `rustSpell/language-rs`
- **Purpose**: Language abstraction trait stack — every Bokmål-specific constant (paths, modal pairs, function words, voice/STT/OCR codes, UI strings) lives behind a trait so adding Nynorsk/English/etc. is a single new file
- **Key files**: `src/lib.rs` (LanguageProfile + 6 sub-traits + LanguageBundle supertrait + BokmalLanguage impl + resolve_language registry)
- **See**: [`adding-a-language.md`](adding-a-language.md) for the per-language walkthrough

## External Dependencies

| Dependency | Path (macOS) | Path (Windows) | Purpose |
|-----------|--------------|----------------|---------|
| SWI-Prolog | `/Applications/SWI-Prolog.app/Contents/Frameworks/libswipl.dylib` | `C:\Program Files\swipl\bin\libswipl.dll` | Grammar checking (loaded dynamically at runtime) |
| ONNX Runtime | `/opt/homebrew/lib/libonnxruntime.dylib` | `onnxruntime/` (vendored) | BERT model inference (CPU) |
| OpenVINO | (not used on macOS) | `openvino/` (vendored) | BERT inference (2x faster on Intel CPUs) |

The runtime loads each library dynamically. SWI-Prolog and ONNX Runtime
must be installed before the app starts; OpenVINO is auto-detected
(falls back to ORT if not present).

## Building and running (macOS)

### One-time setup

```bash
# 1. Install runtime dependencies
brew install onnxruntime
brew install --cask swi-prolog
# (rustup must already be installed)

# 2. Clone the four repos as siblings under dev/dyslex/
mkdir -p ~/dev/dyslex && cd ~/dev/dyslex
git clone git@github.com:pegesund/contexterGui.git
git clone git@github.com:pegesund/contexter.git contexter-repo
git clone git@github.com:pegesund/rustSpell.git
git clone git@github.com:pegesund/syntaxer.git

# 3. Symlink the path-dep targets so the Cargo.toml ../../ paths resolve.
#    (contexterGui/Cargo.toml uses ../../rustSpell etc. which is the
#    Windows layout — these symlinks make Mac match.)
ln -s ~/dev/dyslex/rustSpell      ~/dev/rustSpell
ln -s ~/dev/dyslex/contexter-repo ~/dev/contexter-repo
```

### Build

From `contexterGui/`:

```bash
cargo build --bin acatts-rust            # debug
cargo build --release --bin acatts-rust  # release
```

The first build takes 5–10 min (large dependency tree: eframe, ort,
tokenizers, ndarray, rustls, cpal, …). Subsequent builds are
incremental and finish in seconds.

### Run

```bash
# default: Bokmål
./target/debug/acatts-rust &

# pick a language explicitly (only nb / no are registered today)
./target/debug/acatts-rust --language nb &

# disable grammar checking
./target/debug/acatts-rust --no-grammar &

# faster (lower-quality) BERT inference
./target/debug/acatts-rust --quality 0 &

# show the debug tab in the UI
./target/debug/acatts-rust --debug &

# headless spelling test mode (no GUI, runs the same pipeline)
./target/debug/acatts-rust --test-spelling
```

The app opens a small always-on-top window that follows the cursor in
Word / Google Docs / accessibility-enabled apps. It listens on
`https://127.0.0.1:3000` for the Word add-in to connect.

### Verifying the build

Banner on startup should print:

```
Grammar completion: ON
SWI-Prolog engine: ON
Language: Bokmål (nb)
Quality: 1 (Normal)
Loaded dictionary with 633618 entries in ~100ms
Loading NorBERT4 from .../norbert4_base_int8.onnx
Word Add-in HTTPS bridge (native-tls) listening on port 3000
Sentence splitter loaded from .../syntaxer/sentence_split.pl
Grammar actor: SWI-Prolog loaded on actor thread
```

If a path fails to resolve (typically the FST, ONNX model, or Prolog
rules) the app crashes early with the path in the error message.

### Reloading the Word add-in

Two scripts in `contexterGui/`:

```bash
./reconnect_addin.sh   # add-in dropped its connection (shows "Prøv på nytt")
./reload_addin.sh      # add-in is loaded but needs a refresh
```

Always try `reconnect_addin.sh` first.

## Architecture

```
Chrome Extension (Google Docs)
         │ file: norsktale-browser.json
         v
┌─────────────────────────────────────────┐
│  acatts-rust (GUI app)                  │
│  ├── Word COM bridge (Microsoft Word)   │
│  ├── Browser bridge (Google Docs)       │
│  ├── Accessibility bridge (other apps)  │
│  ├── BERT worker thread                 │
│  └── language: Arc<dyn LanguageBundle>  │ ← language-rs
│       │                                 │
│       v                                 │
│  nostos-cognio (NLP engine)             │
│  ├── NorBERT4 (ORT / OpenVINO)          │
│  ├── SWI-Prolog grammar checker ──────────── syntaxer (grammar_rules.pl)
│  └── mtag-rs (dictionary) ────────────────── rustSpell (fullform_bm.mfst)
└─────────────────────────────────────────┘
```

SWI-Prolog is the only grammar engine. The Neo path (and the
neorusticus dependency) was removed in the multiLanguage refactor.

## Chromebook / Android Path

For running on Chromebooks as an Android app:
- Same Rust codebase (nostos-cognio + egui)
- Chrome extension communicates via localhost HTTP (replacing file bridge)
- ONNX Runtime has Android ARM support
- SWI-Prolog: minimal C build for ARM
- See `docs/gpu.md` in nostos-cognio for inference backend benchmarks
