use super::{TtsEngine, VoiceInfo};
use language::LanguageVoice as _;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc;
use std::sync::{OnceLock, RwLock};

// Phase 8: TTS voice constants come from the Language trait. Bokmål is
// hard-coded here for now; later phases pipe a runtime language through.
const BOKMAL: language::BokmalLanguage = language::BokmalLanguage;

static TTS_SENDER: OnceLock<mpsc::Sender<String>> = OnceLock::new();
static TTS_AVAILABLE: AtomicBool = AtomicBool::new(false);
static TTS_SPEAKING: AtomicBool = AtomicBool::new(false);
static TTS_STOP: AtomicBool = AtomicBool::new(false);
static CURRENT_VOICE: OnceLock<RwLock<String>> = OnceLock::new();

pub struct MacTtsEngine;

impl MacTtsEngine {
    pub fn new() -> Self {
        CURRENT_VOICE.get_or_init(|| RwLock::new(BOKMAL.tts_default_voice().to_string()));

        let (tx, rx) = mpsc::channel::<String>();
        std::thread::spawn(move || {
            while let Ok(word) = rx.recv() {
                TTS_STOP.store(false, Ordering::Relaxed);
                TTS_SPEAKING.store(true, Ordering::Relaxed);
                let voice = CURRENT_VOICE.get()
                    .map(|v| v.read().unwrap().clone())
                    .unwrap_or_else(|| BOKMAL.tts_default_voice().to_string());
                let child = std::process::Command::new("say")
                    .arg("-v").arg(&voice)
                    .arg(&word)
                    .spawn().ok();
                if let Some(mut c) = child {
                    // Poll for stop flag while say is running
                    loop {
                        if TTS_STOP.load(Ordering::Relaxed) {
                            let _ = c.kill();
                            let _ = c.wait();
                            break;
                        }
                        match c.try_wait() {
                            Ok(Some(_)) => break,    // process finished
                            Ok(None) => {}           // still running
                            Err(_) => break,
                        }
                        std::thread::sleep(std::time::Duration::from_millis(50));
                    }
                }
                TTS_SPEAKING.store(false, Ordering::Relaxed);
            }
        });
        TTS_SENDER.get_or_init(|| tx);
        TTS_AVAILABLE.store(true, Ordering::Relaxed);
        MacTtsEngine
    }

    /// Query macOS for Norwegian voices by running `say -v '?'`
    fn query_voices() -> Vec<VoiceInfo> {
        let output = std::process::Command::new("say")
            .arg("-v").arg("?")
            .output()
            .ok();

        let output = match output {
            Some(o) => String::from_utf8_lossy(&o.stdout).to_string(),
            None => return Vec::new(),
        };

        let mut voices = Vec::new();
        let voice_filters = BOKMAL.tts_voice_filters();
        for line in output.lines() {
            // Filter set comes from the Language trait (Bokmål returns
            // ["nb_NO", "nn_NO", "no_NO"] so both Bokmål and Nynorsk show up)
            if !voice_filters.iter().any(|f| line.contains(f)) {
                continue;
            }
            // Format: "Name    lang    # Sample text"
            let parts: Vec<&str> = line.splitn(2, '#').collect();
            let sample = parts.get(1).map(|s| s.trim().to_string()).unwrap_or_default();

            let name_lang = parts[0].trim();
            // Name and language are separated by whitespace
            let tokens: Vec<&str> = name_lang.split_whitespace().collect();
            if tokens.len() >= 2 {
                let lang = tokens.last().unwrap().to_string();
                let name = tokens[..tokens.len()-1].join(" ");
                voices.push(VoiceInfo { name, language: lang, sample_text: sample });
            }
        }
        voices
    }
}

impl TtsEngine for MacTtsEngine {
    fn speak(&self, text: &str) {
        if let Some(tx) = TTS_SENDER.get() { let _ = tx.send(text.to_string()); }
    }

    fn is_available(&self) -> bool {
        TTS_AVAILABLE.load(Ordering::Relaxed)
    }

    fn is_speaking(&self) -> bool {
        TTS_SPEAKING.load(Ordering::Relaxed)
    }

    fn stop(&self) {
        TTS_STOP.store(true, Ordering::Release);
    }

    fn available_voices(&self) -> Vec<VoiceInfo> {
        Self::query_voices()
    }

    fn current_voice(&self) -> String {
        CURRENT_VOICE.get()
            .map(|v| v.read().unwrap().clone())
            .unwrap_or_else(|| BOKMAL.tts_default_voice().to_string())
    }

    fn set_voice(&self, name: &str) {
        if let Some(v) = CURRENT_VOICE.get() {
            *v.write().unwrap() = name.to_string();
        }
    }
}
