// TTS module — trait-based platform dispatch.
// Composite engine: Piper (cross-platform ONNX) + system voice
// (SAPI 5 on Windows, `say` on macOS). User picks via voice id prefix.

use std::path::{Path, PathBuf};
use std::sync::OnceLock;

// --- Voice info ---

#[derive(Clone, Debug)]
pub struct VoiceInfo {
    pub name: String,
    pub language: String,
    pub sample_text: String,
}

// --- Trait definition ---

pub trait TtsEngine: Send + Sync {
    fn speak(&self, text: &str);
    fn is_available(&self) -> bool;
    fn is_speaking(&self) -> bool;
    fn stop(&self);
    fn available_voices(&self) -> Vec<VoiceInfo>;
    fn current_voice(&self) -> String;
    fn set_voice(&self, name: &str);
}

// --- Platform implementations ---

#[cfg(target_os = "windows")]
pub mod system_windows;

#[cfg(target_os = "macos")]
mod macos_impl;

pub mod multi_backend;
pub mod piper_engine;

// --- Global engine instance ---

static ENGINE: OnceLock<Box<dyn TtsEngine>> = OnceLock::new();

/// Initialize the platform TTS engine.
/// Called once at startup. After this, use speak_word/tts_available/etc.
pub fn init_tts(lang: &dyn language::LanguageVoice) {
    let piper_root = piper_data_root();
    let espeak = espeak_binary_path(&piper_root);
    let engine: Box<dyn TtsEngine> = create_engine(&piper_root, &espeak, lang);
    let _ = ENGINE.set(engine);
}

/// Root directory for downloaded Piper assets (models + espeak-ng binary).
/// Layout:
///   <root>/nb-NO/...                     Norwegian model + FST
///   <root>/en_US-lessac-medium/...       Each English voice in its own dir
///   <root>/bin/espeak-ng[.exe]           Subprocess binary
pub fn piper_data_root() -> PathBuf {
    crate::downloader::data_dir().join("piper")
}

#[cfg(target_os = "windows")]
fn espeak_binary_path(piper_root: &Path) -> PathBuf {
    piper_root.join("bin").join("espeak-ng.exe")
}

#[cfg(target_os = "macos")]
fn espeak_binary_path(piper_root: &Path) -> PathBuf {
    piper_root.join("bin").join("espeak-ng")
}

#[cfg(not(any(target_os = "windows", target_os = "macos")))]
fn espeak_binary_path(piper_root: &Path) -> PathBuf {
    piper_root.join("bin").join("espeak-ng")
}

/// Pick the initial voice id used by the engine before any saved-pref override.
///
/// Per Petter's directive (2026-04-29):
///   - macOS: always default to the system voice (Nora / Samantha — high quality
///     voices that ship with the OS). Piper is opt-in.
///   - Windows: default to the Piper voice for the language when its assets are
///     downloaded. SAPI 5's default voice is mediocre, so Piper is preferred when
///     available. Falls back to SAPI default if Piper assets are absent.
fn default_voice_id(piper_root: &Path, lang: &dyn language::LanguageVoice) -> String {
    #[cfg(target_os = "macos")]
    {
        let _ = piper_root;
        return lang.tts_default_voice().to_string();
    }

    #[cfg(target_os = "windows")]
    {
        use language::LanguageProfile as _;
        let preferred_piper = match lang.code() {
            "nb" | "nn" => Some(piper_engine::VOICE_NB_NO),
            "en" => Some(piper_engine::VOICE_EN_US_LESSAC),
            _ => None,
        };
        if let Some(voice_id) = preferred_piper {
            if piper_engine::voice_assets_exist(piper_root, voice_id) {
                return voice_id.to_string();
            }
        }
        return system_windows::SYSTEM_VOICE_ID.to_string();
    }

    #[cfg(not(any(target_os = "windows", target_os = "macos")))]
    {
        let _ = piper_root;
        let _ = lang;
        return String::new();
    }
}

#[cfg(target_os = "macos")]
fn create_engine(
    piper_root: &Path,
    espeak: &Path,
    lang: &dyn language::LanguageVoice,
) -> Box<dyn TtsEngine> {
    let default = default_voice_id(piper_root, lang);
    let piper = piper_engine::PiperTtsEngine::new(
        piper_root.to_path_buf(),
        espeak.to_path_buf(),
        default.clone(),
    );
    let system: Box<dyn TtsEngine> = Box::new(macos_impl::MacTtsEngine::new(lang));
    let engine = multi_backend::MultiBackendTtsEngine::new(piper, system);
    engine.set_voice(&default);
    Box::new(engine)
}

#[cfg(target_os = "windows")]
fn create_engine(
    piper_root: &Path,
    espeak: &Path,
    lang: &dyn language::LanguageVoice,
) -> Box<dyn TtsEngine> {
    let default = default_voice_id(piper_root, lang);
    let piper = piper_engine::PiperTtsEngine::new(
        piper_root.to_path_buf(),
        espeak.to_path_buf(),
        default.clone(),
    );
    let system: Box<dyn TtsEngine> = Box::new(system_windows::WindowsSystemTtsEngine::new());
    let engine = multi_backend::MultiBackendTtsEngine::new(piper, system);
    engine.set_voice(&default);
    Box::new(engine)
}

#[cfg(not(any(target_os = "windows", target_os = "macos")))]
fn create_engine(
    _piper_root: &Path,
    _espeak: &Path,
    _lang: &dyn language::LanguageVoice,
) -> Box<dyn TtsEngine> {
    Box::new(StubTtsEngine)
}

// --- Public free functions (unchanged API for main.rs) ---

pub fn speak_word(word: &str) {
    if let Some(e) = ENGINE.get() { e.speak(word); }
}

pub fn tts_available() -> bool {
    ENGINE.get().map_or(false, |e| e.is_available())
}

pub fn is_speaking() -> bool {
    ENGINE.get().map_or(false, |e| e.is_speaking())
}

pub fn stop_speaking() {
    if let Some(e) = ENGINE.get() { e.stop(); }
}

pub fn available_voices() -> Vec<VoiceInfo> {
    ENGINE.get().map_or(Vec::new(), |e| e.available_voices())
}

pub fn current_voice() -> String {
    ENGINE.get().map_or(String::new(), |e| e.current_voice())
}

pub fn set_voice(name: &str) {
    if let Some(e) = ENGINE.get() { e.set_voice(name); }
}

// --- Stub for unsupported platforms ---

struct StubTtsEngine;

impl TtsEngine for StubTtsEngine {
    fn speak(&self, _text: &str) {}
    fn is_available(&self) -> bool { false }
    fn is_speaking(&self) -> bool { false }
    fn stop(&self) {}
    fn available_voices(&self) -> Vec<VoiceInfo> { Vec::new() }
    fn current_voice(&self) -> String { String::new() }
    fn set_voice(&self, _name: &str) {}
}
