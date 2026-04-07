use super::{SttEngine, mic_log};
use language::LanguageVoice as _;
use libloading::Library;
use std::ffi::{CStr, CString};
use std::os::raw::{c_char, c_int, c_float, c_void};

// Phase 9: STT language code comes from the Language trait. Bokmål is
// hard-coded here for now; later phases pipe a runtime language through.
const BOKMAL: language::BokmalLanguage = language::BokmalLanguage;

type WhisperContext = c_void;

const WHISPER_SAMPLING_GREEDY: c_int = 0;
const PARAMS_SIZE: usize = 296;

// Field offsets (from bundled bindings, verified for x64)
const OFF_N_THREADS: usize = 4;
const OFF_PRINT_SPECIAL: usize = 24;
const OFF_PRINT_PROGRESS: usize = 25;
const OFF_PRINT_REALTIME: usize = 26;
const OFF_PRINT_TIMESTAMPS: usize = 27;
const OFF_LANGUAGE: usize = 96;

type FnInit = unsafe extern "C" fn(*const c_char) -> *mut WhisperContext;
type FnFree = unsafe extern "C" fn(*mut WhisperContext);
type FnDefaultParams = unsafe extern "C" fn(c_int) -> [u8; PARAMS_SIZE];
type FnFull = unsafe extern "C" fn(*mut WhisperContext, [u8; PARAMS_SIZE], *const c_float, c_int) -> c_int;
type FnSegments = unsafe extern "C" fn(*mut WhisperContext) -> c_int;
type FnSegmentText = unsafe extern "C" fn(*mut WhisperContext, c_int) -> *const c_char;

/// Preloaded Whisper engine — holds the DLL and model context
pub struct WhisperEngine {
    _lib: Library,
    ctx: *mut WhisperContext,
    fn_free: FnFree,
    fn_default_params: FnDefaultParams,
    fn_full: FnFull,
    fn_n_segments: FnSegments,
    fn_segment_text: FnSegmentText,
}

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
}

impl SttEngine for WhisperEngine {
    fn transcribe(&self, audio: &[f32]) -> String {
        unsafe {
            let mut params = (self.fn_default_params)(WHISPER_SAMPLING_GREEDY);

            params[OFF_N_THREADS..OFF_N_THREADS+4].copy_from_slice(&4_i32.to_ne_bytes());
            params[OFF_PRINT_SPECIAL] = 0;
            params[OFF_PRINT_PROGRESS] = 0;
            params[OFF_PRINT_REALTIME] = 0;
            params[OFF_PRINT_TIMESTAMPS] = 0;

            let lang_c = CString::new(BOKMAL.stt_language_code()).unwrap();
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
