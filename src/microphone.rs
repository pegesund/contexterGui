use std::sync::mpsc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use libloading::Library;
use std::ffi::{CStr, CString};
use std::os::raw::{c_char, c_int, c_float, c_void};

static MIC_RECORDING: AtomicBool = AtomicBool::new(false);

/// Log to the shared log file (same as main.rs log! macro)
fn mic_log(msg: &str) {
    use std::io::Write;
    let path = std::env::temp_dir().join("acatts-rust.log");
    if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true).open(&path) {
        let _ = writeln!(f, "{}", msg);
    }
}

// Whisper C API types (opaque pointers)
type WhisperContext = c_void;

// whisper_sampling_strategy enum
const WHISPER_SAMPLING_GREEDY: c_int = 0;

// whisper_full_params is 296 bytes — treated as opaque blob
const PARAMS_SIZE: usize = 296;

// Field offsets (from bundled bindings, verified for x64)
const OFF_N_THREADS: usize = 4;
const OFF_PRINT_SPECIAL: usize = 24;
const OFF_PRINT_PROGRESS: usize = 25;
const OFF_PRINT_REALTIME: usize = 26;
const OFF_PRINT_TIMESTAMPS: usize = 27;
const OFF_LANGUAGE: usize = 96;

// Function pointer types for whisper C API
type FnInit = unsafe extern "C" fn(*const c_char) -> *mut WhisperContext;
type FnFree = unsafe extern "C" fn(*mut WhisperContext);
type FnDefaultParams = unsafe extern "C" fn(c_int) -> [u8; PARAMS_SIZE];
type FnFull = unsafe extern "C" fn(*mut WhisperContext, [u8; PARAMS_SIZE], *const c_float, c_int) -> c_int;
type FnSegments = unsafe extern "C" fn(*mut WhisperContext) -> c_int;
type FnSegmentText = unsafe extern "C" fn(*mut WhisperContext, c_int) -> *const c_char;

/// Preloaded Whisper engine — holds the DLL and model context
pub struct WhisperEngine {
    _lib: Library, // must stay alive for function pointers
    ctx: *mut WhisperContext,
    fn_free: FnFree,
    fn_default_params: FnDefaultParams,
    fn_full: FnFull,
    fn_n_segments: FnSegments,
    fn_segment_text: FnSegmentText,
}

// Raw pointers are !Send by default, but we protect with Mutex
unsafe impl Send for WhisperEngine {}
unsafe impl Sync for WhisperEngine {}

impl Drop for WhisperEngine {
    fn drop(&mut self) {
        if !self.ctx.is_null() {
            unsafe { (self.fn_free)(self.ctx); }
            mic_log("Whisper: model freed");
        }
    }
}

impl WhisperEngine {
    /// Load whisper.dll and the GGML model. Call once at startup.
    pub fn load(dll_dir: &str, model_path: &str) -> Result<Self, String> {
        // Add DLL directory to search path so ggml*.dll can be found
        #[cfg(target_os = "windows")]
        unsafe {
            use windows::Win32::System::LibraryLoader::SetDllDirectoryW;
            use windows::core::HSTRING;
            let dir = HSTRING::from(dll_dir);
            let _ = SetDllDirectoryW(&dir);
        }

        let dll_path = format!("{}\\whisper.dll", dll_dir);
        let lib = unsafe { Library::new(&dll_path) }
            .map_err(|e| format!("kunne ikke laste whisper.dll: {}", e))?;

        unsafe {
            let fn_init: FnInit = *lib.get::<FnInit>(b"whisper_init_from_file")
                .map_err(|e| format!("whisper_init_from_file not found: {}", e))?;
            let fn_free: FnFree = *lib.get::<FnFree>(b"whisper_free").unwrap();
            let fn_default_params: FnDefaultParams = *lib.get::<FnDefaultParams>(b"whisper_full_default_params").unwrap();
            let fn_full: FnFull = *lib.get::<FnFull>(b"whisper_full").unwrap();
            let fn_n_segments: FnSegments = *lib.get::<FnSegments>(b"whisper_full_n_segments").unwrap();
            let fn_segment_text: FnSegmentText = *lib.get::<FnSegmentText>(b"whisper_full_get_segment_text").unwrap();

            let model_c = CString::new(model_path)
                .map_err(|_| "ugyldig modellsti".to_string())?;

            mic_log(&format!("Whisper: loading model from {}...", model_path));
            let start = std::time::Instant::now();
            let ctx = fn_init(model_c.as_ptr());
            if ctx.is_null() {
                return Err("kunne ikke laste Whisper-modell".into());
            }
            mic_log(&format!("Whisper: model loaded in {:.1}s", start.elapsed().as_secs_f64()));

            Ok(WhisperEngine {
                _lib: lib,
                ctx,
                fn_free,
                fn_default_params,
                fn_full,
                fn_n_segments,
                fn_segment_text,
            })
        }
    }

    /// Transcribe 16kHz mono f32 audio
    pub fn transcribe(&self, audio: &[f32]) -> String {
        unsafe {
            let mut params = (self.fn_default_params)(WHISPER_SAMPLING_GREEDY);

            // Patch fields
            params[OFF_N_THREADS..OFF_N_THREADS+4].copy_from_slice(&4_i32.to_ne_bytes());
            params[OFF_PRINT_SPECIAL] = 0;
            params[OFF_PRINT_PROGRESS] = 0;
            params[OFF_PRINT_REALTIME] = 0;
            params[OFF_PRINT_TIMESTAMPS] = 0;

            let lang_c = CString::new("no").unwrap();
            let lang_ptr_bytes = (lang_c.as_ptr() as usize).to_ne_bytes();
            params[OFF_LANGUAGE..OFF_LANGUAGE+8].copy_from_slice(&lang_ptr_bytes);

            mic_log(&format!("Whisper: transcribing {} samples ({:.1}s)...", audio.len(), audio.len() as f64 / 16000.0));
            let start = std::time::Instant::now();

            let ret = (self.fn_full)(self.ctx, params, audio.as_ptr(), audio.len() as c_int);
            if ret != 0 {
                return format!("Feil: Whisper-transkribering feilet (kode {})", ret);
            }

            let elapsed = start.elapsed();
            mic_log(&format!("Whisper: transcription took {:.1}s", elapsed.as_secs_f64()));

            let n_segments = (self.fn_n_segments)(self.ctx);
            let mut result = String::new();
            for i in 0..n_segments {
                let text_ptr = (self.fn_segment_text)(self.ctx, i);
                if !text_ptr.is_null() {
                    if let Ok(text) = CStr::from_ptr(text_ptr).to_str() {
                        if !result.is_empty() {
                            result.push(' ');
                        }
                        result.push_str(text.trim());
                    }
                }
            }

            drop(lang_c);

            if result.is_empty() {
                "(ingen tale gjenkjent)".into()
            } else {
                result
            }
        }
    }
}

/// Result from whisper transcription
pub struct TranscribeResult {
    pub text: String,
    /// true = partial (more coming), false = final
    pub partial: bool,
}

/// Handle for controlling microphone recording + whisper transcription
pub struct MicHandle {
    stop_flag: Arc<AtomicBool>,
    pub result_rx: mpsc::Receiver<TranscribeResult>,
}

impl MicHandle {
    /// Stop recording and trigger transcription
    pub fn stop(&self) {
        self.stop_flag.store(true, Ordering::Release);
    }
}

/// Check if currently recording
pub fn is_recording() -> bool {
    MIC_RECORDING.load(Ordering::Relaxed)
}

/// Force-clear the recording flag so the UI can close immediately.
/// The background thread will finish on its own and send to a dropped channel.
pub fn force_stop() {
    MIC_RECORDING.store(false, Ordering::Relaxed);
}

/// Start recording from default microphone.
/// Uses `streaming_engine` (base) for fast partial transcription during recording,
/// and `final_engine` (medium-q5) for high-quality final transcription after stop.
pub fn start_recording(final_engine: Arc<Mutex<WhisperEngine>>, streaming_engine: Arc<Mutex<WhisperEngine>>) -> Result<MicHandle, String> {
    if MIC_RECORDING.load(Ordering::Relaxed) {
        return Err("Already recording".into());
    }

    let stop_flag = Arc::new(AtomicBool::new(false));
    let (result_tx, result_rx) = mpsc::channel();

    let stop_clone = stop_flag.clone();

    MIC_RECORDING.store(true, Ordering::Relaxed);

    // Spawn thread that creates stream (cpal::Stream is !Send), records, then runs whisper
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

        // Streaming transcription: transcribe periodically while recording
        // Uses base model for fast partials (0.24x realtime)
        // First chunk after 2s, then every 2s of new audio
        let first_chunk_samples = sample_rate as usize * 2; // 2 seconds
        let chunk_interval_samples = sample_rate as usize * 2; // 2 seconds
        let mut last_transcribed_len: usize = 0;
        let mut next_threshold = first_chunk_samples;

        loop {
            if stop_clone.load(Ordering::Acquire) {
                break;
            }
            std::thread::sleep(std::time::Duration::from_millis(100));

            // Check if we have enough new audio for a partial transcription
            let current_len = {
                let buf = audio_buf.lock().unwrap();
                buf.len()
            };

            if current_len >= next_threshold && current_len > last_transcribed_len {
                // Grab all audio so far
                let raw_audio = {
                    let buf = audio_buf.lock().unwrap();
                    buf.clone()
                };

                // Resample to 16kHz
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

        // Drop stream to stop recording
        drop(stream);
        MIC_RECORDING.store(false, Ordering::Relaxed);
        mic_log("Microphone: recording stopped");

        // Final transcription of all audio
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

        mic_log(&format!("Microphone: final transcribe {:.1}s (medium-q5)...",
            audio_16k.len() as f64 / 16000.0));
        let final_start = std::time::Instant::now();
        let text = {
            let eng = final_engine.lock().unwrap();
            eng.transcribe(&audio_16k)
        };
        mic_log(&format!("Whisper final result in {:.1}s: '{}'", final_start.elapsed().as_secs_f64(), text));
        let _ = result_tx.send(TranscribeResult { text, partial: false });
    });

    Ok(MicHandle { stop_flag, result_rx })
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
        1, // mono
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
