use super::SttEngine;

/// macOS STT engine — placeholder for Apple SFSpeechRecognizer integration.
/// TODO: Implement using Apple's on-device speech recognition.
pub struct MacSttEngine;

impl MacSttEngine {
    pub fn new() -> Self {
        MacSttEngine
    }
}

impl SttEngine for MacSttEngine {
    fn transcribe(&self, _audio: &[f32]) -> String {
        "(tale-til-tekst ikke implementert for macOS ennå)".into()
    }
}
