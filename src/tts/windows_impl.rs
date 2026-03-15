use libloading::{Library, Symbol};
use std::ffi::CString;
use std::sync::mpsc;
use std::sync::atomic::{AtomicBool, Ordering};

type LpBabTts = *mut std::ffi::c_void;
type BabTtsError = i32;
const E_BABTTS_NOERROR: BabTtsError = 0;
const BABTTS_TXT_UTF8: u32 = 0x00000008;
const W_BABTTS_NOMOREDATA: BabTtsError = 3;

type FnBabTtsInit = unsafe extern "system" fn() -> bool;
type FnBabTtsUninit = unsafe extern "system" fn() -> bool;
type FnBabTtsCreate = unsafe extern "system" fn() -> LpBabTts;
type FnBabTtsOpen = unsafe extern "system" fn(LpBabTts, *const i8, u32) -> BabTtsError;
type FnBabTtsClose = unsafe extern "system" fn(LpBabTts) -> BabTtsError;
type FnBabTtsInsertText = unsafe extern "system" fn(LpBabTts, *const i8, u32) -> BabTtsError;
type FnBabTtsReadBuffer = unsafe extern "system" fn(LpBabTts, *mut u8, u32, *mut u32) -> BabTtsError;

type PlaySoundFn = unsafe extern "system" fn(*const u8, usize, u32) -> i32;
const SND_MEMORY: u32 = 0x0004;
const SND_ASYNC: u32 = 0x0001;
const SND_NODEFAULT: u32 = 0x0002;

static TTS_SENDER: std::sync::OnceLock<mpsc::Sender<String>> = std::sync::OnceLock::new();
static TTS_AVAILABLE: AtomicBool = AtomicBool::new(false);
static TTS_SPEAKING: AtomicBool = AtomicBool::new(false);
static TTS_STOP: AtomicBool = AtomicBool::new(false);

fn build_wav(samples: &[u8]) -> Vec<u8> {
    let sample_rate: u32 = 22050;
    let channels: u16 = 1;
    let bits_per_sample: u16 = 16;
    let byte_rate = sample_rate * (channels as u32) * (bits_per_sample as u32) / 8;
    let block_align = channels * bits_per_sample / 8;
    let silence_bytes = 6616;
    let data_size = (silence_bytes + samples.len()) as u32;
    let mut wav = Vec::with_capacity(44 + data_size as usize);
    wav.extend_from_slice(b"RIFF");
    wav.extend_from_slice(&(36 + data_size).to_le_bytes());
    wav.extend_from_slice(b"WAVE");
    wav.extend_from_slice(b"fmt ");
    wav.extend_from_slice(&16u32.to_le_bytes());
    wav.extend_from_slice(&1u16.to_le_bytes());
    wav.extend_from_slice(&channels.to_le_bytes());
    wav.extend_from_slice(&sample_rate.to_le_bytes());
    wav.extend_from_slice(&byte_rate.to_le_bytes());
    wav.extend_from_slice(&block_align.to_le_bytes());
    wav.extend_from_slice(&bits_per_sample.to_le_bytes());
    wav.extend_from_slice(b"data");
    wav.extend_from_slice(&data_size.to_le_bytes());
    wav.extend_from_slice(&vec![0u8; silence_bytes]);
    wav.extend_from_slice(samples);
    wav
}

pub fn init_tts(sdk_dir: &str, voice_name: &str) {
    let sdk_dir = sdk_dir.to_string();
    let voice_name = voice_name.to_string();
    let (tx, rx) = mpsc::channel::<String>();
    let (ready_tx, ready_rx) = mpsc::channel::<bool>();
    std::thread::spawn(move || {
        let dll_path = format!("{}\\AcaTTS.64.dll", sdk_dir);
        let lib = match unsafe { Library::new(&dll_path) } {
            Ok(l) => l,
            Err(e) => { eprintln!("Acapela TTS: failed to load {}: {}", dll_path, e); let _ = ready_tx.send(false); return; }
        };
        let winmm = match unsafe { Library::new("winmm.dll") } {
            Ok(l) => l,
            Err(e) => { eprintln!("Acapela TTS: failed to load winmm.dll: {}", e); let _ = ready_tx.send(false); return; }
        };
        let play_sound: PlaySoundFn = unsafe { let f: Symbol<PlaySoundFn> = winmm.get(b"PlaySoundW").unwrap(); *f };
        unsafe {
            let fn_init: Symbol<FnBabTtsInit> = lib.get(b"BabTTS_Init").unwrap();
            let fn_create: Symbol<FnBabTtsCreate> = lib.get(b"BabTTS_Create").unwrap();
            let fn_open: Symbol<FnBabTtsOpen> = lib.get(b"BabTTS_Open").unwrap();
            let fn_insert: Symbol<FnBabTtsInsertText> = lib.get(b"BabTTS_InsertText").unwrap();
            let fn_read: Symbol<FnBabTtsReadBuffer> = lib.get(b"BabTTS_ReadBuffer").unwrap();
            let fn_close: Symbol<FnBabTtsClose> = lib.get(b"BabTTS_Close").unwrap();
            let fn_uninit: Symbol<FnBabTtsUninit> = lib.get(b"BabTTS_Uninit").unwrap();
            if !fn_init() { eprintln!("Acapela TTS: BabTTS_Init failed"); let _ = ready_tx.send(false); return; }
            let engine = fn_create();
            if engine.is_null() { eprintln!("Acapela TTS: BabTTS_Create returned null"); fn_uninit(); let _ = ready_tx.send(false); return; }
            let voice_c = CString::new(voice_name.as_str()).unwrap();
            let err = fn_open(engine, voice_c.as_ptr(), 0);
            if err != E_BABTTS_NOERROR { eprintln!("Acapela TTS: BabTTS_Open({}) failed: {}", voice_name, err); fn_close(engine); fn_uninit(); let _ = ready_tx.send(false); return; }
            eprintln!("Acapela TTS initialized with voice: {}", voice_name);
            let _ = ready_tx.send(true);
            let mut pcm_buf = [0u8; 4096];
            while let Ok(word) = rx.recv() {
                TTS_STOP.store(false, Ordering::Relaxed);
                let text_c = match CString::new(word.as_str()) { Ok(c) => c, Err(_) => continue };
                let err = fn_insert(engine, text_c.as_ptr(), BABTTS_TXT_UTF8);
                if err != E_BABTTS_NOERROR { eprintln!("TTS InsertText error: {}", err); continue; }
                let mut all_samples: Vec<u8> = Vec::new();
                loop {
                    let mut generated: u32 = 0;
                    let err = fn_read(engine, pcm_buf.as_mut_ptr(), (pcm_buf.len() / 2) as u32, &mut generated);
                    if generated > 0 { all_samples.extend_from_slice(&pcm_buf[..(generated as usize * 2)]); }
                    if err == W_BABTTS_NOMOREDATA || err < 0 { break; }
                }
                if !all_samples.is_empty() && !TTS_STOP.load(Ordering::Relaxed) {
                    let wav = build_wav(&all_samples);
                    TTS_SPEAKING.store(true, Ordering::Relaxed);
                    play_sound(wav.as_ptr(), 0, SND_MEMORY | SND_ASYNC | SND_NODEFAULT);
                    let duration_ms = (all_samples.len() as u64 * 1000) / (22050 * 2) + 300;
                    let start = std::time::Instant::now();
                    while start.elapsed().as_millis() < duration_ms as u128 {
                        if TTS_STOP.load(Ordering::Relaxed) { play_sound(std::ptr::null(), 0, 0); break; }
                        std::thread::sleep(std::time::Duration::from_millis(50));
                    }
                    TTS_SPEAKING.store(false, Ordering::Relaxed);
                }
            }
            fn_close(engine);
            fn_uninit();
        }
    });
    if ready_rx.recv().unwrap_or(false) {
        TTS_SENDER.get_or_init(|| tx);
        TTS_AVAILABLE.store(true, Ordering::Relaxed);
    }
}

pub fn speak_word(word: &str) {
    if let Some(tx) = TTS_SENDER.get() { let _ = tx.send(word.to_string()); }
}

pub fn tts_available() -> bool { TTS_AVAILABLE.load(Ordering::Relaxed) }
pub fn is_speaking() -> bool { TTS_SPEAKING.load(Ordering::Relaxed) }
pub fn stop_speaking() { TTS_STOP.store(true, Ordering::Release); }
