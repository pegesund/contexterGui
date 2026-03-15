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

### neorusticus
- **GitHub**: pegesund/neorusticus
- **Path**: `neorusticus`
- **Purpose**: Pure Rust Prolog engine — zero dependencies, fallback when SWI-Prolog is not available
- **Note**: Grammar rules are embedded via `include_str!` (synced from syntaxer via `sync_grammar_rules.py`)

## External Dependencies

| Dependency | Version | Path | Purpose |
|-----------|---------|------|---------|
| SWI-Prolog | 9.2.9 | `C:\Program Files\swipl\bin\libswipl.dll` | Grammar checking (loaded dynamically) |
| ONNX Runtime | 1.23.0 | `onnxruntime/` | BERT model inference (CPU) |
| OpenVINO | 2025.4.0 | `openvino/` | BERT inference (2x faster on Intel CPUs) |

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
│  └── BERT worker thread                 │
│       │                                 │
│       v                                 │
│  nostos-cognio (NLP engine)             │
│  ├── NorBERT4 (ORT/OpenVINO)           │
│  ├── SWI-Prolog grammar checker ──────────── syntaxer (grammar_rules.pl)
│  ├── mtag-rs (dictionary) ────────────────── rustSpell (fullform_bm.mfst)
│  └── neorusticus (Prolog fallback)      │
└─────────────────────────────────────────┘
```

## Chromebook / Android Path

For running on Chromebooks as an Android app:
- Same Rust codebase (nostos-cognio + egui)
- Chrome extension communicates via localhost HTTP (replacing file bridge)
- ONNX Runtime has Android ARM support
- SWI-Prolog: minimal C build for ARM, or fallback to neorusticus
- See `docs/gpu.md` in nostos-cognio for inference backend benchmarks
