// TTS module — trait-based platform dispatch.
// Windows: Acapela BabTTS. macOS: `say` command. Other: no-op.

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
mod windows_impl;

#[cfg(target_os = "macos")]
mod macos_impl;

// --- Global engine instance ---

static ENGINE: OnceLock<Box<dyn TtsEngine>> = OnceLock::new();

/// Initialize the platform TTS engine.
/// Called once at startup. After this, use speak_word/tts_available/etc.
pub fn init_tts(sdk_dir: &str, voice_name: &str) {
    let engine: Box<dyn TtsEngine> = create_engine(sdk_dir, voice_name);
    let _ = ENGINE.set(engine);
}

#[cfg(target_os = "macos")]
fn create_engine(_sdk_dir: &str, _voice_name: &str) -> Box<dyn TtsEngine> {
    Box::new(macos_impl::MacTtsEngine::new())
}

#[cfg(target_os = "windows")]
fn create_engine(sdk_dir: &str, voice_name: &str) -> Box<dyn TtsEngine> {
    Box::new(windows_impl::WindowsTtsEngine::new(sdk_dir, voice_name))
}

#[cfg(not(any(target_os = "windows", target_os = "macos")))]
fn create_engine(_sdk_dir: &str, _voice_name: &str) -> Box<dyn TtsEngine> {
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
