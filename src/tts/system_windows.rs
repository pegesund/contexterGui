//! Windows system TTS via SAPI 5 (`ISpVoice`).
//!
//! Uses the system default voice. Works fully offline — SAPI 5 ships with every
//! Windows since XP and never needs network access.

use super::{TtsEngine, VoiceInfo};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc;
use std::sync::{Arc, RwLock};

use windows::Win32::Media::Speech::{
    ISpVoice, SpVoice, SPF_ASYNC, SPF_PURGEBEFORESPEAK, SPVOICESTATUS, SPRS_DONE,
};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, COINIT_MULTITHREADED,
};
use windows::core::HSTRING;

pub const SYSTEM_VOICE_ID: &str = "system:default";

struct SpeakCmd {
    text: String,
    stop_current: bool,
}

pub struct WindowsSystemTtsEngine {
    sender: mpsc::Sender<SpeakCmd>,
    speaking: Arc<AtomicBool>,
    stop: Arc<AtomicBool>,
    available: Arc<AtomicBool>,
    current_voice: RwLock<String>,
}

impl WindowsSystemTtsEngine {
    pub fn new() -> Self {
        let (tx, rx) = mpsc::channel::<SpeakCmd>();
        let speaking = Arc::new(AtomicBool::new(false));
        let stop = Arc::new(AtomicBool::new(false));
        let available = Arc::new(AtomicBool::new(false));

        let speaking_w = speaking.clone();
        let stop_w = stop.clone();
        let available_w = available.clone();

        std::thread::spawn(move || unsafe {
            let coinit = CoInitializeEx(None, COINIT_MULTITHREADED);
            if coinit.is_err() {
                eprintln!("SAPI 5: CoInitializeEx failed: {:?}", coinit);
                return;
            }

            let voice: ISpVoice = match CoCreateInstance(&SpVoice, None, CLSCTX_ALL) {
                Ok(v) => v,
                Err(e) => {
                    eprintln!("SAPI 5: SpVoice CoCreateInstance failed: {:?}", e);
                    CoUninitialize();
                    return;
                }
            };
            available_w.store(true, Ordering::Release);

            while let Ok(cmd) = rx.recv() {
                stop_w.store(false, Ordering::Relaxed);
                let text = HSTRING::from(cmd.text.as_str());
                let mut flags = SPF_ASYNC.0 as u32;
                if cmd.stop_current {
                    flags |= SPF_PURGEBEFORESPEAK.0 as u32;
                }
                if let Err(e) = voice.Speak(&text, flags, None) {
                    eprintln!("SAPI 5: Speak failed: {:?}", e);
                    continue;
                }
                speaking_w.store(true, Ordering::Relaxed);

                loop {
                    if stop_w.load(Ordering::Relaxed) {
                        let empty = HSTRING::new();
                        let _ = voice.Speak(&empty, SPF_PURGEBEFORESPEAK.0 as u32, None);
                        break;
                    }
                    let mut status: SPVOICESTATUS = std::mem::zeroed();
                    if voice.GetStatus(&mut status, std::ptr::null_mut()).is_err() {
                        break;
                    }
                    if status.dwRunningState == SPRS_DONE.0 as u32 {
                        break;
                    }
                    std::thread::sleep(std::time::Duration::from_millis(30));
                }
                speaking_w.store(false, Ordering::Relaxed);
            }

            drop(voice);
            CoUninitialize();
        });

        // Wait briefly for the worker to finish CoCreateInstance so callers
        // get a meaningful is_available() after construction.
        let start = std::time::Instant::now();
        while !available.load(Ordering::Acquire)
            && start.elapsed() < std::time::Duration::from_millis(2000)
        {
            std::thread::sleep(std::time::Duration::from_millis(10));
        }

        WindowsSystemTtsEngine {
            sender: tx,
            speaking,
            stop,
            available,
            current_voice: RwLock::new(SYSTEM_VOICE_ID.to_string()),
        }
    }
}

impl TtsEngine for WindowsSystemTtsEngine {
    fn speak(&self, text: &str) {
        let _ = self.sender.send(SpeakCmd {
            text: text.to_string(),
            stop_current: true,
        });
    }

    fn is_available(&self) -> bool {
        self.available.load(Ordering::Acquire)
    }

    fn is_speaking(&self) -> bool {
        self.speaking.load(Ordering::Relaxed)
    }

    fn stop(&self) {
        self.stop.store(true, Ordering::Release);
    }

    fn available_voices(&self) -> Vec<VoiceInfo> {
        // MVP: expose the system default voice as a single selectable entry.
        // Per-voice enumeration via ISpObjectTokenCategory is a follow-up.
        vec![VoiceInfo {
            name: SYSTEM_VOICE_ID.to_string(),
            language: "system".to_string(),
            sample_text: "Dette er Windows-stemmen.".to_string(),
        }]
    }

    fn current_voice(&self) -> String {
        self.current_voice.read().unwrap().clone()
    }

    fn set_voice(&self, name: &str) {
        *self.current_voice.write().unwrap() = name.to_string();
    }
}
