//! Composite TTS engine routing by voice id prefix.
//!
//! Routing rules:
//!   - voice id starts with `"piper:"` → `PiperTtsEngine`
//!   - everything else → platform system engine (SAPI 5 on Windows, `say` on macOS)
//!
//! Voice list = Piper voices first, then system voices.

use super::piper_engine::PiperTtsEngine;
use super::{TtsEngine, VoiceInfo};

pub struct MultiBackendTtsEngine {
    piper: PiperTtsEngine,
    system: Box<dyn TtsEngine>,
}

impl MultiBackendTtsEngine {
    pub fn new(piper: PiperTtsEngine, system: Box<dyn TtsEngine>) -> Self {
        Self { piper, system }
    }

    fn is_piper_voice(name: &str) -> bool {
        name.starts_with("piper:")
    }

    fn active(&self) -> &dyn TtsEngine {
        if Self::is_piper_voice(&self.piper.current_voice()) {
            &self.piper
        } else {
            self.system.as_ref()
        }
    }
}

impl TtsEngine for MultiBackendTtsEngine {
    fn speak(&self, text: &str) {
        // Stop the inactive backend so a re-route doesn't leave a stale stream playing.
        if Self::is_piper_voice(&self.piper.current_voice()) {
            self.system.stop();
            self.piper.speak(text);
        } else {
            self.piper.stop();
            self.system.speak(text);
        }
    }

    fn is_available(&self) -> bool {
        self.piper.is_available() || self.system.is_available()
    }

    fn is_speaking(&self) -> bool {
        self.piper.is_speaking() || self.system.is_speaking()
    }

    fn stop(&self) {
        self.piper.stop();
        self.system.stop();
    }

    fn available_voices(&self) -> Vec<VoiceInfo> {
        let mut all = self.piper.available_voices();
        all.extend(self.system.available_voices());
        all
    }

    fn current_voice(&self) -> String {
        // Return whichever backend is currently active for the user-selected voice.
        let piper_v = self.piper.current_voice();
        if Self::is_piper_voice(&piper_v) {
            piper_v
        } else {
            self.system.current_voice()
        }
    }

    fn set_voice(&self, name: &str) {
        // Keep both backends in sync so routing stays consistent on the next speak().
        self.piper.set_voice(name);
        self.system.set_voice(name);
        // Switching voice should silence any in-flight utterance from the
        // previous backend. The previous `let _ = self.active();` was a no-op
        // (just discarded the &dyn TtsEngine reference) so a Piper utterance
        // would keep playing after switching to a system voice and vice
        // versa. Stop both — the still-active backend has nothing to say
        // until the next speak() call.
        self.piper.stop();
        self.system.stop();
    }
}
