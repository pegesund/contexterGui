//! Piper TTS engine — cross-platform (Windows + macOS).
//!
//! Pipeline per utterance:
//!   text → norsk_g2p::synthesis::{NorwegianVoice|EnglishVoice}.synthesize()
//!        → f32 PCM at 22050 Hz
//!        → norsk_g2p::inference::write_wav() to temp file
//!        → OS playback (PlaySoundW on Windows, afplay on macOS)
//!
//! Voice ids:
//!   "piper:nb-NO"
//!   "piper:en_US-lessac-medium"
//!   "piper:en_US-amy-medium"
//!   "piper:en_GB-alba-medium"
//!   "piper:en_GB-northern_english_male-medium"

use super::{TtsEngine, VoiceInfo};
use norsk_g2p::inference::write_wav;
use norsk_g2p::synthesis::{EnglishVoice, NorwegianVoice};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc;
use std::sync::{Arc, RwLock};

pub const VOICE_NB_NO: &str = "piper:nb-NO";
pub const VOICE_EN_US_LESSAC: &str = "piper:en_US-lessac-medium";
pub const VOICE_EN_US_AMY: &str = "piper:en_US-amy-medium";
pub const VOICE_EN_GB_ALBA: &str = "piper:en_GB-alba-medium";
pub const VOICE_EN_GB_NORTHERN: &str = "piper:en_GB-northern_english_male-medium";

const PIPER_VOICES: &[(&str, &str, &str)] = &[
    (VOICE_NB_NO, "nb_NO", "Hei, dette er norsk talesyntese."),
    (VOICE_EN_US_LESSAC, "en_US", "Hello, this is English text to speech."),
    (VOICE_EN_US_AMY, "en_US", "Hello, this is English text to speech."),
    (VOICE_EN_GB_ALBA, "en_GB", "Hello, this is British English text to speech."),
    (VOICE_EN_GB_NORTHERN, "en_GB", "Hello, this is Northern English text to speech."),
];

struct SpeakCmd {
    voice_id: String,
    text: String,
}

enum LoadedVoice {
    Norwegian(NorwegianVoice),
    English(EnglishVoice),
}

pub struct PiperTtsEngine {
    sender: mpsc::Sender<SpeakCmd>,
    speaking: Arc<AtomicBool>,
    stop: Arc<AtomicBool>,
    current_voice: RwLock<String>,
    asset_dir: PathBuf,
}

impl PiperTtsEngine {
    pub fn new(asset_dir: PathBuf, espeak_binary: PathBuf, default_voice: String) -> Self {
        let (tx, rx) = mpsc::channel::<SpeakCmd>();
        let speaking = Arc::new(AtomicBool::new(false));
        let stop = Arc::new(AtomicBool::new(false));

        let asset_dir_w = asset_dir.clone();
        let speaking_w = speaking.clone();
        let stop_w = stop.clone();

        std::thread::spawn(move || {
            let mut loaded: Option<(String, LoadedVoice)> = None;
            while let Ok(mut cmd) = rx.recv() {
                while let Ok(newer) = rx.try_recv() {
                    cmd = newer;
                }
                stop_w.store(false, Ordering::Relaxed);

                let need_load = match &loaded {
                    Some((id, _)) => id != &cmd.voice_id,
                    None => true,
                };
                if need_load {
                    loaded = None;
                    match load_voice(&cmd.voice_id, &asset_dir_w, &espeak_binary) {
                        Ok(v) => loaded = Some((cmd.voice_id.clone(), v)),
                        Err(e) => {
                            eprintln!(
                                "Piper TTS: failed to load voice {}: {}",
                                cmd.voice_id, e
                            );
                            continue;
                        }
                    }
                }

                let synth_result = match &mut loaded {
                    Some((_, LoadedVoice::Norwegian(v))) => v.synthesize(&cmd.text),
                    Some((_, LoadedVoice::English(v))) => v.synthesize(&cmd.text),
                    None => continue,
                };
                let synth = match synth_result {
                    Ok(r) => r,
                    Err(e) => {
                        eprintln!("Piper TTS: synth failed: {}", e);
                        continue;
                    }
                };

                let tmp_path = match temp_wav_path() {
                    Some(p) => p,
                    None => {
                        eprintln!("Piper TTS: no temp dir available");
                        continue;
                    }
                };
                if let Err(e) = write_wav(
                    tmp_path.to_str().unwrap_or(""),
                    &synth.samples,
                    synth.sample_rate,
                ) {
                    eprintln!("Piper TTS: write_wav failed: {}", e);
                    continue;
                }

                speaking_w.store(true, Ordering::Relaxed);
                play_wav(&tmp_path, &synth.samples, synth.sample_rate, &stop_w);
                speaking_w.store(false, Ordering::Relaxed);

                let _ = std::fs::remove_file(&tmp_path);
            }
        });

        PiperTtsEngine {
            sender: tx,
            speaking,
            stop,
            current_voice: RwLock::new(default_voice),
            asset_dir,
        }
    }

    /// True if every required asset for the given voice id exists on disk.
    pub fn is_voice_ready(&self, voice_id: &str) -> bool {
        match voice_paths(voice_id, &self.asset_dir) {
            Some(paths) => paths.iter().all(|p| p.exists()),
            None => false,
        }
    }

    /// All Piper voice ids supported by this build.
    pub fn all_voice_ids() -> &'static [&'static str] {
        const IDS: &[&str] = &[
            VOICE_NB_NO,
            VOICE_EN_US_LESSAC,
            VOICE_EN_US_AMY,
            VOICE_EN_GB_ALBA,
            VOICE_EN_GB_NORTHERN,
        ];
        IDS
    }
}

/// Check if every required asset file for a voice id exists on disk.
/// Used at startup to decide the initial default voice without constructing
/// a full engine.
pub fn voice_assets_exist(asset_dir: &Path, voice_id: &str) -> bool {
    voice_paths(voice_id, asset_dir)
        .map(|paths| paths.iter().all(|p| p.exists()))
        .unwrap_or(false)
}

fn voice_paths(voice_id: &str, asset_dir: &Path) -> Option<Vec<PathBuf>> {
    if voice_id == VOICE_NB_NO {
        let base = asset_dir.join("nb-NO");
        return Some(vec![
            base.join("epoch_649_v5.onnx"),
            base.join("epoch_649_v5.onnx.json"),
            base.join("lexicon.fst"),
            base.join("lexicon_values.bin"),
            base.join("lexicon_phonemes.txt"),
            base.join("pronunciation_overrides.tsv"),
        ]);
    }
    if let Some(voice) = voice_id.strip_prefix("piper:") {
        if voice.starts_with("en_") {
            let base = asset_dir.join(voice);
            let mut paths = vec![
                base.join(format!("{}.onnx", voice)),
                base.join(format!("{}.onnx.json", voice)),
            ];
            paths.push(asset_dir.join("bin").join(if cfg!(target_os = "windows") {
                "espeak-ng.exe"
            } else {
                "espeak-ng"
            }));
            return Some(paths);
        }
    }
    None
}

fn load_voice(
    voice_id: &str,
    asset_dir: &Path,
    espeak_binary: &Path,
) -> Result<LoadedVoice, Box<dyn std::error::Error>> {
    if voice_id == VOICE_NB_NO {
        let base = asset_dir.join("nb-NO");
        let v = NorwegianVoice::load(
            &base.join("epoch_649_v5.onnx"),
            &base.join("epoch_649_v5.onnx.json"),
            &base.join("lexicon.fst"),
            &base.join("lexicon_values.bin"),
            &base.join("lexicon_phonemes.txt"),
            &base.join("pronunciation_overrides.tsv"),
        )?;
        return Ok(LoadedVoice::Norwegian(v));
    }
    if let Some(voice) = voice_id.strip_prefix("piper:") {
        if voice.starts_with("en_") {
            let base = asset_dir.join(voice);
            let espeak_voice = if voice.starts_with("en_GB") {
                "en-gb"
            } else {
                "en-us"
            };
            // S3 strips file modes; ensure the espeak-ng binary has +x on Unix.
            #[cfg(target_os = "macos")]
            {
                use std::os::unix::fs::PermissionsExt;
                if let Ok(meta) = std::fs::metadata(espeak_binary) {
                    let mut perms = meta.permissions();
                    if perms.mode() & 0o111 == 0 {
                        perms.set_mode(0o755);
                        let _ = std::fs::set_permissions(espeak_binary, perms);
                    }
                }
            }
            let v = EnglishVoice::load(
                &base.join(format!("{}.onnx", voice)),
                &base.join(format!("{}.onnx.json", voice)),
                espeak_binary.to_path_buf(),
                espeak_voice.to_string(),
            )?;
            return Ok(LoadedVoice::English(v));
        }
    }
    Err(format!("Unknown Piper voice id: {}", voice_id).into())
}

fn temp_wav_path() -> Option<PathBuf> {
    let dir = std::env::temp_dir();
    let pid = std::process::id();
    let nanos = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_nanos())
        .unwrap_or(0);
    Some(dir.join(format!("spell-piper-{}-{}.wav", pid, nanos)))
}

#[cfg(target_os = "windows")]
fn play_wav(path: &Path, samples: &[f32], sample_rate: u32, stop: &AtomicBool) {
    use libloading::{Library, Symbol};
    use std::os::windows::ffi::OsStrExt;

    type PlaySoundFn = unsafe extern "system" fn(*const u16, usize, u32) -> i32;
    const SND_FILENAME: u32 = 0x00020000;
    const SND_ASYNC: u32 = 0x00000001;
    const SND_NODEFAULT: u32 = 0x00000002;

    let winmm = match unsafe { Library::new("winmm.dll") } {
        Ok(l) => l,
        Err(e) => {
            eprintln!("Piper TTS: winmm.dll load failed: {}", e);
            return;
        }
    };
    let play: Symbol<PlaySoundFn> = match unsafe { winmm.get(b"PlaySoundW") } {
        Ok(s) => s,
        Err(e) => {
            eprintln!("Piper TTS: PlaySoundW lookup failed: {}", e);
            return;
        }
    };
    let wide: Vec<u16> = path
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();
    unsafe {
        play(wide.as_ptr(), 0, SND_FILENAME | SND_ASYNC | SND_NODEFAULT);
    }

    let duration_ms = (samples.len() as u64 * 1000) / sample_rate.max(1) as u64 + 200;
    let start = std::time::Instant::now();
    while start.elapsed().as_millis() < duration_ms as u128 {
        if stop.load(Ordering::Relaxed) {
            unsafe {
                play(std::ptr::null(), 0, 0);
            }
            break;
        }
        std::thread::sleep(std::time::Duration::from_millis(50));
    }
}

#[cfg(target_os = "macos")]
fn play_wav(path: &Path, _samples: &[f32], _sample_rate: u32, stop: &AtomicBool) {
    let child = std::process::Command::new("afplay").arg(path).spawn().ok();
    if let Some(mut c) = child {
        loop {
            if stop.load(Ordering::Relaxed) {
                let _ = c.kill();
                let _ = c.wait();
                break;
            }
            match c.try_wait() {
                Ok(Some(_)) => break,
                Ok(None) => std::thread::sleep(std::time::Duration::from_millis(50)),
                Err(_) => break,
            }
        }
    }
}

#[cfg(not(any(target_os = "windows", target_os = "macos")))]
fn play_wav(_path: &Path, _samples: &[f32], _sample_rate: u32, _stop: &AtomicBool) {}

impl TtsEngine for PiperTtsEngine {
    fn speak(&self, text: &str) {
        let voice_id = self.current_voice.read().unwrap().clone();
        self.stop.store(true, Ordering::Release);
        let _ = self.sender.send(SpeakCmd {
            voice_id,
            text: text.to_string(),
        });
    }

    fn is_available(&self) -> bool {
        true
    }

    fn is_speaking(&self) -> bool {
        self.speaking.load(Ordering::Relaxed)
    }

    fn stop(&self) {
        self.stop.store(true, Ordering::Release);
    }

    fn available_voices(&self) -> Vec<VoiceInfo> {
        PIPER_VOICES
            .iter()
            .map(|(id, lang, sample)| VoiceInfo {
                name: id.to_string(),
                language: lang.to_string(),
                sample_text: sample.to_string(),
            })
            .collect()
    }

    fn current_voice(&self) -> String {
        self.current_voice.read().unwrap().clone()
    }

    fn set_voice(&self, name: &str) {
        *self.current_voice.write().unwrap() = name.to_string();
    }
}
