use super::{TtsEngine, VoiceInfo};
use language::LanguageVoice as _;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc;
use std::sync::{OnceLock, RwLock};

static TTS_SENDER: OnceLock<mpsc::Sender<String>> = OnceLock::new();
static TTS_AVAILABLE: AtomicBool = AtomicBool::new(false);
static TTS_SPEAKING: AtomicBool = AtomicBool::new(false);
static TTS_STOP: AtomicBool = AtomicBool::new(false);
static CURRENT_VOICE: OnceLock<RwLock<String>> = OnceLock::new();
/// Fallback voice name used when CURRENT_VOICE is uninitialised (should not happen).
static FALLBACK_VOICE: OnceLock<String> = OnceLock::new();

pub struct MacTtsEngine {
    voice_filters: &'static [&'static str],
}

impl MacTtsEngine {
    pub fn new(lang: &dyn language::LanguageVoice) -> Self {
        let default_voice = lang.tts_default_voice();
        CURRENT_VOICE.get_or_init(|| RwLock::new(default_voice.to_string()));
        FALLBACK_VOICE.get_or_init(|| default_voice.to_string());

        let (tx, rx) = mpsc::channel::<String>();
        std::thread::spawn(move || {
            while let Ok(word) = rx.recv() {
                TTS_STOP.store(false, Ordering::Relaxed);
                TTS_SPEAKING.store(true, Ordering::Relaxed);
                let voice = CURRENT_VOICE.get()
                    .map(|v| v.read().unwrap().clone())
                    .unwrap_or_else(|| FALLBACK_VOICE.get().cloned().unwrap_or_default());
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
        MacTtsEngine { voice_filters: lang.tts_voice_filters() }
    }

    /// Query macOS for voices matching the language's filter set by running `say -v '?'`
    fn query_voices(voice_filters: &'static [&'static str]) -> Vec<VoiceInfo> {
        let output = std::process::Command::new("say")
            .arg("-v").arg("?")
            .output()
            .ok();

        let output = match output {
            Some(o) => String::from_utf8_lossy(&o.stdout).to_string(),
            None => return Vec::new(),
        };

        let mut voices = Vec::new();
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
        Self::query_voices(self.voice_filters)
    }

    fn current_voice(&self) -> String {
        CURRENT_VOICE.get()
            .map(|v| v.read().unwrap().clone())
            .unwrap_or_else(|| FALLBACK_VOICE.get().cloned().unwrap_or_default())
    }

    fn set_voice(&self, name: &str) {
        if let Some(v) = CURRENT_VOICE.get() {
            *v.write().unwrap() = name.to_string();
        }
    }
}
