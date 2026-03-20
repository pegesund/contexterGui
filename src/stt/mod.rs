// STT module — trait-based platform dispatch.
// Windows: Whisper (C API via DLL). macOS: Apple SFSpeechRecognizer (future).

use std::sync::mpsc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};

static MIC_RECORDING: AtomicBool = AtomicBool::new(false);

/// Trait for speech-to-text engines.
/// Implementations must be safe to call from a background thread.
pub trait SttEngine: Send + Sync {
    /// Transcribe 16kHz mono f32 audio to text.
    fn transcribe(&self, audio: &[f32]) -> String;
}

/// Result from transcription
pub struct TranscribeResult {
    pub text: String,
    /// true = partial (more coming), false = final
    pub partial: bool,
}

/// Handle for controlling microphone recording + transcription
pub struct MicHandle {
    stop_flag: Arc<AtomicBool>,
    pub result_rx: mpsc::Receiver<TranscribeResult>,
}

impl MicHandle {
    pub fn stop(&self) {
        self.stop_flag.store(true, Ordering::Release);
    }
}

pub fn is_recording() -> bool {
    MIC_RECORDING.load(Ordering::Relaxed)
}

pub fn force_stop() {
    MIC_RECORDING.store(false, Ordering::Relaxed);
}

// Platform implementations
#[cfg(target_os = "windows")]
mod windows_impl;
#[cfg(target_os = "windows")]
pub use windows_impl::WhisperEngine;

#[cfg(target_os = "macos")]
mod macos_impl;
#[cfg(target_os = "macos")]
pub use macos_impl::MacSttEngine;

/// Start recording from default microphone.
/// Uses `streaming_engine` for fast partial transcription during recording,
/// and `final_engine` for high-quality final transcription after stop.
pub fn start_recording(
    final_engine: Arc<Mutex<Box<dyn SttEngine>>>,
    streaming_engine: Arc<Mutex<Box<dyn SttEngine>>>,
) -> Result<MicHandle, String> {
    if MIC_RECORDING.load(Ordering::Relaxed) {
        return Err("Already recording".into());
    }

    let stop_flag = Arc::new(AtomicBool::new(false));
    let (result_tx, result_rx) = mpsc::channel();

    let stop_clone = stop_flag.clone();

    MIC_RECORDING.store(true, Ordering::Relaxed);

    std::thread::spawn(move || {
        use cpal::traits::{DeviceTrait, HostTrait, StreamTrait};

        let host = cpal::default_host();
        let device = match host.default_input_device() {
            Some(d) => d,
            None => {
                mic_log("Microphone: no input device found");
                MIC_RECORDING.store(false, Ordering::Relaxed);
                let _ = result_tx.send(TranscribeResult { text: "(ingen mikrofon funnet)".into(), partial: false });
                return;
            }
        };
        mic_log(&format!("Microphone: using '{}'", device.name().unwrap_or_default()));

        let config = match device.default_input_config() {
            Ok(c) => c,
            Err(e) => {
                mic_log(&format!("Microphone: no config: {}", e));
                MIC_RECORDING.store(false, Ordering::Relaxed);
                let _ = result_tx.send(TranscribeResult { text: format!("Feil: {}", e), partial: false });
                return;
            }
        };
        let sample_rate = config.sample_rate().0;
        let channels = config.channels() as usize;
        mic_log(&format!("Microphone: {}Hz, {} channels, format {:?}", sample_rate, channels, config.sample_format()));

        let audio_buf: Arc<std::sync::Mutex<Vec<f32>>> = Arc::new(std::sync::Mutex::new(Vec::new()));
        let buf_clone = audio_buf.clone();

        let stream = match config.sample_format() {
            cpal::SampleFormat::F32 => {
                device.build_input_stream(
                    &config.into(),
                    move |data: &[f32], _: &cpal::InputCallbackInfo| {
                        let mut buf = buf_clone.lock().unwrap();
                        if channels == 1 {
                            buf.extend_from_slice(data);
                        } else {
                            for chunk in data.chunks(channels) {
                                let sum: f32 = chunk.iter().sum();
                                buf.push(sum / channels as f32);
                            }
                        }
                    },
                    |err| mic_log(&format!("Microphone error: {}", err)),
                    None,
                )
            }
            cpal::SampleFormat::I16 => {
                let buf_clone2 = audio_buf.clone();
                device.build_input_stream(
                    &config.into(),
                    move |data: &[i16], _: &cpal::InputCallbackInfo| {
                        let mut buf = buf_clone2.lock().unwrap();
                        if channels == 1 {
                            buf.extend(data.iter().map(|&s| s as f32 / 32768.0));
                        } else {
                            for chunk in data.chunks(channels) {
                                let sum: f32 = chunk.iter().map(|&s| s as f32 / 32768.0).sum();
                                buf.push(sum / channels as f32);
                            }
                        }
                    },
                    |err| mic_log(&format!("Microphone error: {}", err)),
                    None,
                )
            }
            fmt => {
                mic_log(&format!("Microphone: unsupported format {:?}", fmt));
                MIC_RECORDING.store(false, Ordering::Relaxed);
                let _ = result_tx.send(TranscribeResult { text: format!("Feil: format {:?}", fmt), partial: false });
                return;
            }
        };

        let stream = match stream {
            Ok(s) => s,
            Err(e) => {
                mic_log(&format!("Microphone: failed to build stream: {}", e));
                MIC_RECORDING.store(false, Ordering::Relaxed);
                let _ = result_tx.send(TranscribeResult { text: format!("Feil: {}", e), partial: false });
                return;
            }
        };

        if let Err(e) = stream.play() {
            mic_log(&format!("Microphone: failed to start: {}", e));
            MIC_RECORDING.store(false, Ordering::Relaxed);
            let _ = result_tx.send(TranscribeResult { text: format!("Feil: {}", e), partial: false });
            return;
        }

        mic_log("Microphone: recording started");

        let first_chunk_samples = sample_rate as usize * 2;
        let chunk_interval_samples = sample_rate as usize * 2;
        let mut last_transcribed_len: usize = 0;
        let mut next_threshold = first_chunk_samples;

        loop {
            if stop_clone.load(Ordering::Acquire) {
                break;
            }
            std::thread::sleep(std::time::Duration::from_millis(100));

            let current_len = {
                let buf = audio_buf.lock().unwrap();
                buf.len()
            };

            if current_len >= next_threshold && current_len > last_transcribed_len {
                let raw_audio = {
                    let buf = audio_buf.lock().unwrap();
                    buf.clone()
                };

                let audio_16k = if sample_rate != 16000 {
                    resample(&raw_audio, sample_rate, 16000)
                } else {
                    raw_audio
                };

                mic_log(&format!("Microphone: streaming transcribe {:.1}s of audio...",
                    audio_16k.len() as f64 / 16000.0));
                let start = std::time::Instant::now();
                let text = {
                    let eng = streaming_engine.lock().unwrap();
                    eng.transcribe(&audio_16k)
                };
                mic_log(&format!("Microphone: partial result in {:.1}s: '{}'",
                    start.elapsed().as_secs_f64(), &text[..text.len().min(80)]));

                let _ = result_tx.send(TranscribeResult { text, partial: true });
                last_transcribed_len = current_len;
                next_threshold = current_len + chunk_interval_samples;
            }
        }

        drop(stream);
        MIC_RECORDING.store(false, Ordering::Relaxed);
        mic_log("Microphone: recording stopped");

        let raw_audio = {
            let buf = audio_buf.lock().unwrap();
            buf.clone()
        };

        if raw_audio.is_empty() {
            let _ = result_tx.send(TranscribeResult { text: "(ingen lyd fanget)".into(), partial: false });
            return;
        }

        mic_log(&format!("Microphone: final transcribe {:.1}s of audio",
            raw_audio.len() as f64 / sample_rate as f64));

        let audio_16k = if sample_rate != 16000 {
            resample(&raw_audio, sample_rate, 16000)
        } else {
            raw_audio
        };

        mic_log(&format!("Microphone: final transcribe {:.1}s...",
            audio_16k.len() as f64 / 16000.0));
        let final_start = std::time::Instant::now();
        let text = {
            let eng = final_engine.lock().unwrap();
            eng.transcribe(&audio_16k)
        };
        mic_log(&format!("STT final result in {:.1}s: '{}'", final_start.elapsed().as_secs_f64(), text));
        let _ = result_tx.send(TranscribeResult { text, partial: false });
    });

    Ok(MicHandle { stop_flag, result_rx })
}

/// Log to the shared log file
fn mic_log(msg: &str) {
    use std::io::Write;
    let path = std::env::temp_dir().join("acatts-rust.log");
    if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true).open(&path) {
        let _ = writeln!(f, "{}", msg);
    }
}

/// Resample audio from one sample rate to another using rubato
fn resample(input: &[f32], from_rate: u32, to_rate: u32) -> Vec<f32> {
    use rubato::{SincFixedIn, SincInterpolationType, SincInterpolationParameters, WindowFunction, Resampler};

    let params = SincInterpolationParameters {
        sinc_len: 256,
        f_cutoff: 0.95,
        interpolation: SincInterpolationType::Linear,
        oversampling_factor: 256,
        window: WindowFunction::BlackmanHarris2,
    };

    let chunk_size = input.len().min(1024 * 1024);
    let mut resampler = match SincFixedIn::<f32>::new(
        to_rate as f64 / from_rate as f64,
        2.0,
        params,
        chunk_size,
        1,
    ) {
        Ok(r) => r,
        Err(e) => {
            mic_log(&format!("Resample error: {}", e));
            return input.to_vec();
        }
    };

    let waves_in = vec![input.to_vec()];
    match resampler.process(&waves_in, None) {
        Ok(waves_out) => waves_out.into_iter().next().unwrap_or_default(),
        Err(e) => {
            mic_log(&format!("Resample error: {}", e));
            input.to_vec()
        }
    }
}
