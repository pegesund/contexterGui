// TTS module — platform dispatch.
// Windows: Acapela BabTTS. macOS: `say` command. Other: no-op.

#[cfg(target_os = "windows")]
mod windows_impl;
#[cfg(target_os = "windows")]
pub use windows_impl::*;

#[cfg(target_os = "macos")]
mod macos_impl {
    use std::sync::atomic::{AtomicBool, Ordering};
    use std::sync::mpsc;

    static TTS_SENDER: std::sync::OnceLock<mpsc::Sender<String>> = std::sync::OnceLock::new();
    static TTS_AVAILABLE: AtomicBool = AtomicBool::new(false);
    static TTS_SPEAKING: AtomicBool = AtomicBool::new(false);
    static TTS_STOP: AtomicBool = AtomicBool::new(false);

    pub fn init_tts(_sdk_dir: &str, _voice_name: &str) {
        let (tx, rx) = mpsc::channel::<String>();
        std::thread::spawn(move || {
            while let Ok(word) = rx.recv() {
                TTS_STOP.store(false, Ordering::Relaxed);
                TTS_SPEAKING.store(true, Ordering::Relaxed);
                let mut child = std::process::Command::new("say")
                    .arg("-v").arg("Nora")
                    .arg(&word)
                    .spawn().ok();
                if let Some(ref mut c) = child { let _ = c.wait(); }
                TTS_SPEAKING.store(false, Ordering::Relaxed);
            }
        });
        TTS_SENDER.get_or_init(|| tx);
        TTS_AVAILABLE.store(true, Ordering::Relaxed);
    }

    pub fn speak_word(word: &str) {
        if let Some(tx) = TTS_SENDER.get() { let _ = tx.send(word.to_string()); }
    }
    pub fn tts_available() -> bool { TTS_AVAILABLE.load(Ordering::Relaxed) }
    pub fn is_speaking() -> bool { TTS_SPEAKING.load(Ordering::Relaxed) }
    pub fn stop_speaking() { TTS_STOP.store(true, Ordering::Release); }
}
#[cfg(target_os = "macos")]
pub use macos_impl::*;

#[cfg(not(any(target_os = "windows", target_os = "macos")))]
pub fn init_tts(_sdk_dir: &str, _voice_name: &str) {}
#[cfg(not(any(target_os = "windows", target_os = "macos")))]
pub fn speak_word(_word: &str) {}
#[cfg(not(any(target_os = "windows", target_os = "macos")))]
pub fn tts_available() -> bool { false }
#[cfg(not(any(target_os = "windows", target_os = "macos")))]
pub fn is_speaking() -> bool { false }
#[cfg(not(any(target_os = "windows", target_os = "macos")))]
pub fn stop_speaking() {}
