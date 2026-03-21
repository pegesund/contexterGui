use super::{SttEngine, MicHandle, TranscribeResult, MIC_RECORDING};
use std::sync::{Arc, Mutex, atomic::{AtomicBool, Ordering}};
use std::sync::mpsc;

fn stt_log(msg: &str) {
    use std::io::Write;
    if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true).open("/tmp/acatts-stt.log") {
        let _ = writeln!(f, "{}", msg);
    }
}

/// macOS STT engine — batch transcription from audio buffer (used as fallback)
pub struct MacSttEngine;

impl MacSttEngine {
    pub fn new() -> Self { MacSttEngine }
}

impl SttEngine for MacSttEngine {
    fn transcribe(&self, audio: &[f32]) -> String {
        let wav_path = std::env::temp_dir().join("acatts_stt_input.wav");
        if write_wav(&wav_path, audio, 16000).is_err() { return String::new(); }
        let result = transcribe_file(&wav_path).unwrap_or_default();
        let _ = std::fs::remove_file(&wav_path);
        result
    }
}

/// Start live recording with streaming SFSpeechRecognizer.
/// Audio from cpal is fed directly to SFSpeechAudioBufferRecognitionRequest.
/// Partial results stream to the UI while recording.
pub fn start_recording_live() -> Result<MicHandle, String> {
    if MIC_RECORDING.load(Ordering::Relaxed) {
        return Err("Already recording".into());
    }

    let stop_flag = Arc::new(AtomicBool::new(false));
    let (result_tx, result_rx) = mpsc::channel();
    let stop_clone = stop_flag.clone();

    MIC_RECORDING.store(true, Ordering::Relaxed);

    std::thread::spawn(move || {
        use cpal::traits::{DeviceTrait, HostTrait, StreamTrait};
        use objc::runtime::{Class, Object, BOOL, YES};
        use objc::*;
        use block::ConcreteBlock;
        use std::ffi::CStr;

        let host = cpal::default_host();
        let device = match host.default_input_device() {
            Some(d) => d,
            None => {
                stt_log("No input device");
                MIC_RECORDING.store(false, Ordering::Relaxed);
                let _ = result_tx.send(TranscribeResult { text: "(ingen mikrofon funnet)".into(), partial: false });
                return;
            }
        };

        let config = match device.default_input_config() {
            Ok(c) => c,
            Err(e) => {
                stt_log(&format!("No config: {}", e));
                MIC_RECORDING.store(false, Ordering::Relaxed);
                let _ = result_tx.send(TranscribeResult { text: format!("Feil: {}", e), partial: false });
                return;
            }
        };
        let sample_rate = config.sample_rate().0;
        let channels = config.channels() as usize;
        stt_log(&format!("Mic: {}Hz, {} ch", sample_rate, channels));

        // Set up SFSpeechRecognizer on this thread
        let done = Arc::new(AtomicBool::new(false));

        let request_ptr: Arc<Mutex<*mut Object>> = Arc::new(Mutex::new(std::ptr::null_mut()));

        unsafe {
            let locale = {
                let cls = Class::get("NSLocale").unwrap();
                let ns_id = make_nsstring("nb-NO");
                let loc: *mut Object = msg_send![cls, alloc];
                let loc: *mut Object = msg_send![loc, initWithLocaleIdentifier: ns_id];
                loc
            };

            let cls = Class::get("SFSpeechRecognizer").unwrap();
            let recognizer: *mut Object = msg_send![cls, alloc];
            let recognizer: *mut Object = msg_send![recognizer, initWithLocale: locale];

            let req_cls = Class::get("SFSpeechAudioBufferRecognitionRequest").unwrap();
            let request: *mut Object = msg_send![req_cls, alloc];
            let request: *mut Object = msg_send![request, init];
            let _: () = msg_send![request, setRequiresOnDeviceRecognition: YES];
            let _: () = msg_send![request, setShouldReportPartialResults: YES];

            *request_ptr.lock().unwrap() = request;

            let tx = result_tx.clone();
            let d = done.clone();
            let callback = ConcreteBlock::new(move |result: *mut Object, error: *mut Object| {
                if !error.is_null() {
                    let desc: *mut Object = msg_send![error, localizedDescription];
                    let cstr: *const i8 = msg_send![desc, UTF8String];
                    let s = CStr::from_ptr(cstr).to_string_lossy().to_string();
                    stt_log(&format!("STT error: {}", s));
                    let _ = tx.send(TranscribeResult { text: format!("Feil: {}", s), partial: false });
                    d.store(true, Ordering::Relaxed);
                    return;
                }
                if !result.is_null() {
                    let transcription: *mut Object = msg_send![result, bestTranscription];
                    let formatted: *mut Object = msg_send![transcription, formattedString];
                    let cstr: *const i8 = msg_send![formatted, UTF8String];
                    let s = CStr::from_ptr(cstr).to_string_lossy().to_string();
                    let is_final: BOOL = msg_send![result, isFinal];
                    stt_log(&format!("STT {}: '{}'", if is_final == YES { "final" } else { "partial" }, &s[..s.len().min(60)]));
                    let _ = tx.send(TranscribeResult { text: s, partial: is_final != YES });
                    if is_final == YES {
                        d.store(true, Ordering::Relaxed);
                    }
                }
            });
            let callback = callback.copy();

            let _task: *mut Object = msg_send![recognizer, recognitionTaskWithRequest: request resultHandler: &*callback];
            std::mem::forget(callback);
            stt_log("STT session started");
        }

        // Start cpal audio capture
        let audio_buf: Arc<Mutex<Vec<f32>>> = Arc::new(Mutex::new(Vec::new()));
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
                    |err| stt_log(&format!("Mic error: {}", err)),
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
                    |err| stt_log(&format!("Mic error: {}", err)),
                    None,
                )
            }
            fmt => {
                stt_log(&format!("Unsupported format: {:?}", fmt));
                MIC_RECORDING.store(false, Ordering::Relaxed);
                let _ = result_tx.send(TranscribeResult { text: format!("Feil: format {:?}", fmt), partial: false });
                return;
            }
        };

        let stream = match stream {
            Ok(s) => s,
            Err(e) => {
                stt_log(&format!("Failed to build stream: {}", e));
                MIC_RECORDING.store(false, Ordering::Relaxed);
                let _ = result_tx.send(TranscribeResult { text: format!("Feil: {}", e), partial: false });
                return;
            }
        };

        if let Err(e) = stream.play() {
            stt_log(&format!("Failed to start: {}", e));
            MIC_RECORDING.store(false, Ordering::Relaxed);
            let _ = result_tx.send(TranscribeResult { text: format!("Feil: {}", e), partial: false });
            return;
        }

        stt_log("Recording started");

        // Feed audio to SFSpeechAudioBufferRecognitionRequest every 200ms
        // Pump CFRunLoop so recognition callbacks fire on this thread
        let mut last_sent: usize = 0;
        loop {
            if stop_clone.load(Ordering::Acquire) || done.load(Ordering::Relaxed) {
                break;
            }
            // Pump RunLoop to process recognition callbacks
            unsafe {
                core_foundation::runloop::CFRunLoop::run_in_mode(
                    core_foundation::runloop::kCFRunLoopDefaultMode,
                    std::time::Duration::from_millis(200),
                    true,
                );
            }

            let new_samples = {
                let buf = audio_buf.lock().unwrap();
                if buf.len() > last_sent {
                    let samples = buf[last_sent..].to_vec();
                    last_sent = buf.len();
                    Some(samples)
                } else {
                    None
                }
            };

            if let Some(samples) = new_samples {
                // Resample to 16kHz if needed
                let samples_16k = if sample_rate != 16000 {
                    super::resample(&samples, sample_rate, 16000)
                } else {
                    samples
                };

                // Append to recognition request on main thread
                let req = *request_ptr.lock().unwrap();
                if !req.is_null() && !samples_16k.is_empty() {
                    unsafe {
                        append_audio_buffer(req, &samples_16k, 16000);
                    }
                }
            }
        }

        // Stop recording
        drop(stream);

        // End the recognition request
        let req = *request_ptr.lock().unwrap();
        if !req.is_null() {
            unsafe {
                let _: () = objc::msg_send![req, endAudio];
            }
        }

        // Wait for final result (up to 10 seconds)
        for _ in 0..100 {
            if done.load(Ordering::Relaxed) { break; }
            std::thread::sleep(std::time::Duration::from_millis(100));
        }

        MIC_RECORDING.store(false, Ordering::Relaxed);
        stt_log("Recording stopped");
    });

    Ok(MicHandle { stop_flag, result_rx })
}

/// Append f32 audio samples to an SFSpeechAudioBufferRecognitionRequest
unsafe fn append_audio_buffer(request: *mut objc::runtime::Object, samples: &[f32], sample_rate: u32) {
    use objc::runtime::{Class, Object};
    use objc::*;

    // Create AVAudioFormat (mono, float32, 16kHz)
    let format_cls = Class::get("AVAudioFormat").unwrap();
    let format: *mut Object = msg_send![format_cls, alloc];
    let format: *mut Object = msg_send![format, initWithCommonFormat: 1u64  // AVAudioPCMFormatFloat32 = 1
        sampleRate: sample_rate as f64
        channels: 1u32
        interleaved: false];

    // Create AVAudioPCMBuffer
    let buffer_cls = Class::get("AVAudioPCMBuffer").unwrap();
    let buffer: *mut Object = msg_send![buffer_cls, alloc];
    let buffer: *mut Object = msg_send![buffer, initWithPCMFormat: format frameCapacity: samples.len() as u32];
    let _: () = msg_send![buffer, setFrameLength: samples.len() as u32];

    // Copy samples into buffer
    let float_data: *mut *mut f32 = msg_send![buffer, floatChannelData];
    if !float_data.is_null() {
        let channel_ptr = *float_data;
        std::ptr::copy_nonoverlapping(samples.as_ptr(), channel_ptr, samples.len());
    }

    // Append to request
    let _: () = msg_send![request, appendAudioPCMBuffer: buffer];
}

fn write_wav(path: &std::path::Path, samples: &[f32], sample_rate: u32) -> Result<(), String> {
    use std::io::Write;
    let num_samples = samples.len();
    let data_size = (num_samples * 2) as u32;
    let file_size = 36 + data_size;
    let mut f = std::fs::File::create(path).map_err(|e| e.to_string())?;
    f.write_all(b"RIFF").map_err(|e| e.to_string())?;
    f.write_all(&file_size.to_le_bytes()).map_err(|e| e.to_string())?;
    f.write_all(b"WAVE").map_err(|e| e.to_string())?;
    f.write_all(b"fmt ").map_err(|e| e.to_string())?;
    f.write_all(&16u32.to_le_bytes()).map_err(|e| e.to_string())?;
    f.write_all(&1u16.to_le_bytes()).map_err(|e| e.to_string())?;
    f.write_all(&1u16.to_le_bytes()).map_err(|e| e.to_string())?;
    f.write_all(&sample_rate.to_le_bytes()).map_err(|e| e.to_string())?;
    f.write_all(&(sample_rate * 2).to_le_bytes()).map_err(|e| e.to_string())?;
    f.write_all(&2u16.to_le_bytes()).map_err(|e| e.to_string())?;
    f.write_all(&16u16.to_le_bytes()).map_err(|e| e.to_string())?;
    f.write_all(b"data").map_err(|e| e.to_string())?;
    f.write_all(&data_size.to_le_bytes()).map_err(|e| e.to_string())?;
    for &s in samples {
        let i16_val = (s.clamp(-1.0, 1.0) * 32767.0) as i16;
        f.write_all(&i16_val.to_le_bytes()).map_err(|e| e.to_string())?;
    }
    Ok(())
}

fn transcribe_file(path: &std::path::Path) -> Result<String, String> {
    use objc::runtime::{Class, Object, BOOL, YES};
    use objc::*;
    use block::ConcreteBlock;
    use std::ffi::CStr;

    let path_str = path.to_string_lossy().to_string();
    let done = Arc::new(AtomicBool::new(false));
    let result_text = Arc::new(Mutex::new(String::new()));
    let error_text = Arc::new(Mutex::new(String::new()));

    let done2 = done.clone();
    let result2 = result_text.clone();
    let error2 = error_text.clone();
    let path_clone = path_str.clone();

    unsafe {
        let setup_block = ConcreteBlock::new(move || {
            let url = {
                let cls = Class::get("NSURL").unwrap();
                let ns_path = make_nsstring(&path_clone);
                let url: *mut Object = msg_send![cls, fileURLWithPath: ns_path];
                url
            };
            let locale = {
                let cls = Class::get("NSLocale").unwrap();
                let ns_id = make_nsstring("nb-NO");
                let loc: *mut Object = msg_send![cls, alloc];
                let loc: *mut Object = msg_send![loc, initWithLocaleIdentifier: ns_id];
                loc
            };
            let cls = Class::get("SFSpeechRecognizer").unwrap();
            let recognizer: *mut Object = msg_send![cls, alloc];
            let recognizer: *mut Object = msg_send![recognizer, initWithLocale: locale];

            let req_cls = Class::get("SFSpeechURLRecognitionRequest").unwrap();
            let request: *mut Object = msg_send![req_cls, alloc];
            let request: *mut Object = msg_send![request, initWithURL: url];
            let _: () = msg_send![request, setRequiresOnDeviceRecognition: YES];

            let done3 = done2.clone();
            let result3 = result2.clone();
            let error3 = error2.clone();

            let callback = ConcreteBlock::new(move |result: *mut Object, error: *mut Object| {
                if !error.is_null() {
                    let desc: *mut Object = msg_send![error, localizedDescription];
                    let cstr: *const i8 = msg_send![desc, UTF8String];
                    let s = CStr::from_ptr(cstr).to_string_lossy().to_string();
                    *error3.lock().unwrap() = s;
                    done3.store(true, Ordering::Relaxed);
                    return;
                }
                if !result.is_null() {
                    let is_final: BOOL = msg_send![result, isFinal];
                    if is_final == YES {
                        let transcription: *mut Object = msg_send![result, bestTranscription];
                        let formatted: *mut Object = msg_send![transcription, formattedString];
                        let cstr: *const i8 = msg_send![formatted, UTF8String];
                        let s = CStr::from_ptr(cstr).to_string_lossy().to_string();
                        *result3.lock().unwrap() = s;
                        done3.store(true, Ordering::Relaxed);
                    }
                }
            });
            let callback = callback.copy();
            let _task: *mut Object = msg_send![recognizer, recognitionTaskWithRequest: request resultHandler: &*callback];
            std::mem::forget(callback);
        });
        let setup_block = setup_block.copy();

        unsafe extern "C" {
            static _dispatch_main_q: std::ffi::c_void;
            fn dispatch_async(queue: *const std::ffi::c_void, block: &block::Block<(), ()>);
        }
        unsafe { dispatch_async(&_dispatch_main_q, &*setup_block); }
    }

    for _ in 0..300 {
        if done.load(Ordering::Relaxed) { break; }
        std::thread::sleep(std::time::Duration::from_millis(100));
    }

    let err = error_text.lock().unwrap().clone();
    if !err.is_empty() { return Err(err); }
    Ok(result_text.lock().unwrap().clone())
}

unsafe fn make_nsstring(s: &str) -> *mut objc::runtime::Object {
    use objc::*;
    let cls = objc::runtime::Class::get("NSString").unwrap();
    let bytes = s.as_bytes();
    let ns: *mut objc::runtime::Object = msg_send![cls, alloc];
    msg_send![ns, initWithBytes: bytes.as_ptr() length: bytes.len() encoding: 4usize]
}
