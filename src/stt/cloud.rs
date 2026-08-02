//! Cloud STT engine with two backends:
//!
//! - **Proxy** (production): POSTs recorded audio to OUR proxy server, which
//!   holds the actual speech-provider credentials. Client machines never see
//!   a Google token; they only know the proxy URL (+ an app key that can be
//!   rotated server-side). See `server/stt_proxy.py`.
//!   Wire: POST <url>/transcribe, Content-Type audio/wav (16k mono s16le),
//!   X-Spell-Lang, X-Spell-Key → {"text": "..."}
//!
//! - **Google direct** (DEV ONLY): Google Speech-to-Text v1 REST with a plain
//!   API key from settings (`stt_google_api_key`). v1 takes API keys; v2/
//!   chirp needs OAuth, which is exactly why the proxy exists. Anyone with
//!   the machine can extract the key — never ship this to pupils.

use super::{mic_log, SttEngine};

enum Backend {
    Proxy { url: String, app_key: String },
    GoogleDirect { api_key: String },
}

pub struct CloudSttEngine {
    backend: Backend,
    lang: String,
    no_speech_msg: String,
    agent: ureq::Agent,
}

impl CloudSttEngine {
    pub fn new(url: &str, app_key: &str, lang: &str, no_speech_msg: &str) -> Self {
        Self {
            backend: Backend::Proxy {
                url: url.trim_end_matches('/').to_string(),
                app_key: app_key.to_string(),
            },
            lang: lang.to_string(),
            no_speech_msg: no_speech_msg.to_string(),
            agent: ureq::AgentBuilder::new()
                .timeout(std::time::Duration::from_secs(30))
                .build(),
        }
    }

    /// DEV backend: straight to Google STT v1 with an API key.
    pub fn new_google_direct(api_key: &str, lang: &str, no_speech_msg: &str) -> Self {
        Self {
            backend: Backend::GoogleDirect { api_key: api_key.to_string() },
            lang: lang.to_string(),
            no_speech_msg: no_speech_msg.to_string(),
            agent: ureq::AgentBuilder::new()
                .timeout(std::time::Duration::from_secs(30))
                .build(),
        }
    }

    /// Whisper-style 2-letter codes → BCP-47 as Google wants them.
    fn google_lang(&self) -> &'static str {
        match self.lang.as_str() {
            "no" | "nb" => "nb-NO",
            "nn" => "nn-NO",
            "en" => "en-US",
            _ => "nb-NO",
        }
    }

    /// 16 kHz mono f32 → s16le WAV bytes.
    fn to_wav(audio: &[f32]) -> Vec<u8> {
        let n_samples = audio.len() as u32;
        let data_len = n_samples * 2;
        let mut wav = Vec::with_capacity(44 + data_len as usize);
        wav.extend_from_slice(b"RIFF");
        wav.extend_from_slice(&(36 + data_len).to_le_bytes());
        wav.extend_from_slice(b"WAVEfmt ");
        wav.extend_from_slice(&16u32.to_le_bytes()); // fmt chunk size
        wav.extend_from_slice(&1u16.to_le_bytes()); // PCM
        wav.extend_from_slice(&1u16.to_le_bytes()); // mono
        wav.extend_from_slice(&16000u32.to_le_bytes()); // sample rate
        wav.extend_from_slice(&32000u32.to_le_bytes()); // byte rate
        wav.extend_from_slice(&2u16.to_le_bytes()); // block align
        wav.extend_from_slice(&16u16.to_le_bytes()); // bits per sample
        wav.extend_from_slice(b"data");
        wav.extend_from_slice(&data_len.to_le_bytes());
        for s in audio {
            let v = (s.clamp(-1.0, 1.0) * 32767.0) as i16;
            wav.extend_from_slice(&v.to_le_bytes());
        }
        wav
    }
}

/// Minimal base64 (standard alphabet, padded) — avoids a new dependency for
/// the dev-only Google-direct path.
fn base64_encode(data: &[u8]) -> String {
    const TABLE: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity((data.len() + 2) / 3 * 4);
    for chunk in data.chunks(3) {
        let b = [chunk[0], *chunk.get(1).unwrap_or(&0), *chunk.get(2).unwrap_or(&0)];
        let n = (u32::from(b[0]) << 16) | (u32::from(b[1]) << 8) | u32::from(b[2]);
        out.push(TABLE[(n >> 18) as usize & 63] as char);
        out.push(TABLE[(n >> 12) as usize & 63] as char);
        out.push(if chunk.len() > 1 { TABLE[(n >> 6) as usize & 63] as char } else { '=' });
        out.push(if chunk.len() > 2 { TABLE[n as usize & 63] as char } else { '=' });
    }
    out
}

impl CloudSttEngine {
    fn transcribe_proxy(&self, url: &str, app_key: &str, audio: &[f32], audio_secs: f64) -> Result<String, String> {
        let wav = Self::to_wav(audio);
        mic_log(&format!(
            "CloudSTT[proxy]: sending {:.1}s audio ({} KB) to {}...",
            audio_secs, wav.len() / 1024, url));
        let resp = self.agent
            .post(&format!("{}/transcribe", url))
            .set("Content-Type", "audio/wav")
            .set("X-Spell-Lang", &self.lang)
            .set("X-Spell-Key", app_key)
            .send_bytes(&wav)
            .map_err(|e| e.to_string())?;
        let body: serde_json::Value = resp.into_json().map_err(|e| e.to_string())?;
        Ok(body.get("text").and_then(|t| t.as_str()).unwrap_or("").trim().to_string())
    }

    fn transcribe_google(&self, api_key: &str, audio: &[f32], audio_secs: f64) -> Result<String, String> {
        // v1 sync recognize: raw s16le samples (LINEAR16, no WAV header),
        // base64 in JSON. Limit ~60s audio — dictation-length.
        let mut pcm = Vec::with_capacity(audio.len() * 2);
        for s in audio {
            pcm.extend_from_slice(&((s.clamp(-1.0, 1.0) * 32767.0) as i16).to_le_bytes());
        }
        mic_log(&format!(
            "CloudSTT[google-direct]: sending {:.1}s audio ({} KB)...",
            audio_secs, pcm.len() / 1024));
        let body = serde_json::json!({
            "config": {
                "encoding": "LINEAR16",
                "sampleRateHertz": 16000,
                "languageCode": self.google_lang(),
                "enableAutomaticPunctuation": true,
                "model": "latest_long",
            },
            "audio": { "content": base64_encode(&pcm) }
        });
        let url = format!("https://speech.googleapis.com/v1/speech:recognize?key={}", api_key);
        let resp = self.agent
            .post(&url)
            .set("Content-Type", "application/json")
            .send_string(&body.to_string())
            .map_err(|e| match e {
                ureq::Error::Status(code, r) => {
                    let detail = r.into_string().unwrap_or_default();
                    format!("HTTP {}: {}", code, detail.chars().take(200).collect::<String>())
                }
                other => other.to_string(),
            })?;
        let body: serde_json::Value = resp.into_json().map_err(|e| e.to_string())?;
        let text = body.get("results").and_then(|r| r.as_array()).map(|results| {
            results.iter()
                .filter_map(|r| r.get("alternatives")
                    .and_then(|a| a.as_array())
                    .and_then(|a| a.first())
                    .and_then(|a| a.get("transcript"))
                    .and_then(|t| t.as_str()))
                .collect::<Vec<_>>()
                .join(" ")
        }).unwrap_or_default();
        Ok(text.trim().to_string())
    }
}

impl SttEngine for CloudSttEngine {
    fn transcribe(&self, audio: &[f32]) -> String {
        let audio_secs = audio.len() as f64 / 16000.0;
        let start = std::time::Instant::now();

        let result = match &self.backend {
            Backend::Proxy { url, app_key } => self.transcribe_proxy(url, app_key, audio, audio_secs),
            Backend::GoogleDirect { api_key } => self.transcribe_google(api_key, audio, audio_secs),
        };

        match result {
            Ok(text) => {
                mic_log(&format!(
                    "CloudSTT: {:.1}s audio transcribed in {:.1}s ({:.1}x realtime)",
                    audio_secs, start.elapsed().as_secs_f64(),
                    audio_secs / start.elapsed().as_secs_f64().max(0.001)));
                if text.is_empty() { self.no_speech_msg.clone() } else { text }
            }
            Err(e) => {
                mic_log(&format!("CloudSTT: request failed: {}", e));
                format!("Feil: fikk ikke kontakt med taletjeneren ({})",
                    e.chars().take(160).collect::<String>())
            }
        }
    }
}
