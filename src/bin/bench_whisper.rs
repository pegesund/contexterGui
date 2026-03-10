/// Benchmark whisper models: tiny, base, small
use libloading::{Library, Symbol};
use std::ffi::{CStr, CString};
use std::os::raw::{c_char, c_int, c_float, c_void};

type WhisperContext = c_void;
const WHISPER_SAMPLING_GREEDY: c_int = 0;
const PARAMS_SIZE: usize = 296;
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

fn main() {
    let manifest = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let dll_dir = manifest.join("../../whisper-build/bin/Release");
    let data_dir = manifest.join("../../contexter-repo/training-data");

    // Set DLL search path
    #[cfg(target_os = "windows")]
    unsafe {
        use windows::Win32::System::LibraryLoader::SetDllDirectoryW;
        use windows::core::HSTRING;
        let dir = HSTRING::from(dll_dir.to_str().unwrap());
        let _ = SetDllDirectoryW(&dir);
    }

    let dll_path = dll_dir.join("whisper.dll");
    let lib = unsafe { Library::new(dll_path.to_str().unwrap()) }.expect("load whisper.dll");

    let fn_init: FnInit = unsafe { *lib.get::<FnInit>(b"whisper_init_from_file").unwrap() };
    let fn_free: FnFree = unsafe { *lib.get::<FnFree>(b"whisper_free").unwrap() };
    let fn_default_params: FnDefaultParams = unsafe { *lib.get::<FnDefaultParams>(b"whisper_full_default_params").unwrap() };
    let fn_full: FnFull = unsafe { *lib.get::<FnFull>(b"whisper_full").unwrap() };
    let fn_n_segments: FnSegments = unsafe { *lib.get::<FnSegments>(b"whisper_full_n_segments").unwrap() };
    let fn_segment_text: FnSegmentText = unsafe { *lib.get::<FnSegmentText>(b"whisper_full_get_segment_text").unwrap() };

    // Generate 5 seconds of 16kHz silence (whisper will output "(ingen tale)" or similar)
    let sample_rate = 16000;
    let duration_secs = 5.0_f32;
    let n_samples = (sample_rate as f32 * duration_secs) as usize;
    let audio: Vec<f32> = vec![0.0; n_samples]; // silence

    println!("Audio: {} samples ({:.1}s at {}Hz)\n", audio.len(), duration_secs, sample_rate);

    let models = [
        ("tiny ", data_dir.join("ggml-nb-whisper-tiny.bin")),
        ("base ", data_dir.join("ggml-nb-whisper-base.bin")),
        ("small", data_dir.join("ggml-nb-whisper-small.bin")),
    ];

    for (name, path) in &models {
        let path_str = path.to_str().unwrap();
        if !path.exists() {
            println!("{}: model not found", name);
            continue;
        }

        let model_c = CString::new(path_str).unwrap();
        let load_start = std::time::Instant::now();
        let ctx = unsafe { fn_init(model_c.as_ptr()) };
        if ctx.is_null() {
            println!("{}: failed to load", name);
            continue;
        }
        println!("{}: loaded in {:.2}s", name, load_start.elapsed().as_secs_f64());

        let lang_c = CString::new("no").unwrap();

        // Transcribe
        let start = std::time::Instant::now();
        unsafe {
            let mut params = fn_default_params(WHISPER_SAMPLING_GREEDY);
            params[OFF_N_THREADS..OFF_N_THREADS+4].copy_from_slice(&4_i32.to_ne_bytes());
            params[OFF_PRINT_SPECIAL] = 0;
            params[OFF_PRINT_PROGRESS] = 0;
            params[OFF_PRINT_REALTIME] = 0;
            params[OFF_PRINT_TIMESTAMPS] = 0;
            let lang_ptr_bytes = (lang_c.as_ptr() as usize).to_ne_bytes();
            params[OFF_LANGUAGE..OFF_LANGUAGE+8].copy_from_slice(&lang_ptr_bytes);

            let ret = fn_full(ctx, params, audio.as_ptr(), audio.len() as c_int);
            let elapsed = start.elapsed().as_secs_f64();
            let realtime_factor = elapsed / duration_secs as f64;

            let mut result = String::new();
            if ret == 0 {
                let n = fn_n_segments(ctx);
                for i in 0..n {
                    let p = fn_segment_text(ctx, i);
                    if !p.is_null() {
                        if let Ok(t) = CStr::from_ptr(p).to_str() {
                            result.push_str(t.trim());
                        }
                    }
                }
            }

            println!("  transcribe: {:.2}s for {:.1}s audio ({:.2}x realtime)",
                elapsed, duration_secs, realtime_factor);
            println!("  result: '{}'", &result[..result.len().min(80)]);
            println!();

            fn_free(ctx);
        }
    }
}
