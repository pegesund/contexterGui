/// Scratch bench: raw speed of the wrapper DLL (spell_whisper_* exports)
/// on the medium-q5 model, isolated from the app. Verifies the _threads
/// export works and measures native-vs-app-contention timing.
use libloading::Library;
use std::ffi::CString;
use std::os::raw::{c_char, c_int, c_float, c_void};

type WhisperContext = c_void;
type FnInit = unsafe extern "C" fn(*const c_char) -> *mut WhisperContext;
type FnFullThreads = unsafe extern "C" fn(*mut WhisperContext, *const c_float, c_int, *const c_char, c_int) -> c_int;

fn main() {
    let dll_dir = "C:/Users/pette/dev/contexter/whisper-build/bin/Release";
    #[cfg(target_os = "windows")]
    unsafe {
        use windows::Win32::System::LibraryLoader::SetDllDirectoryW;
        use windows::core::HSTRING;
        let dir = HSTRING::from(dll_dir);
        let _ = SetDllDirectoryW(&dir);
    }
    let lib = unsafe { Library::new(format!("{}/whisper.dll", dll_dir)) }.expect("load dll");
    let fn_init: FnInit = unsafe { *lib.get(b"spell_whisper_init_from_file").unwrap() };
    let fn_full_threads: FnFullThreads = unsafe { *lib.get(b"spell_whisper_full_threads").expect("threads export missing!") };

    let model = std::env::args().nth(1).unwrap_or_else(||
        "C:/Users/pette/AppData/Roaming/Spell/data/models/whisper/nb/ggml-nb-whisper-medium-q5.bin".into());
    let model_c = CString::new(model.as_str()).unwrap();
    let t0 = std::time::Instant::now();
    let ctx = unsafe { fn_init(model_c.as_ptr()) };
    assert!(!ctx.is_null(), "model load failed");
    println!("model loaded in {:.1}s", t0.elapsed().as_secs_f64());

    // 13.6s of faint noise — encoder cost dominates and is content-independent.
    let n = (16000.0 * 13.6) as usize;
    let audio: Vec<f32> = (0..n).map(|i| ((i as f32 * 0.01).sin()) * 0.001).collect();
    let lang = CString::new("nb").unwrap();

    for threads in [4, 8, 12] {
        let t = std::time::Instant::now();
        let ret = unsafe { fn_full_threads(ctx, audio.as_ptr(), audio.len() as c_int, lang.as_ptr(), threads as c_int) };
        println!("threads={} ret={} took {:.1}s for 13.6s audio", threads, ret, t.elapsed().as_secs_f64());
    }
}
