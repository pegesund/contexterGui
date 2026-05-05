//! Piper TTS smoke test.
//!
//! Looks for downloaded Piper assets at the standard data root. If present,
//! synthesizes a short Norwegian sentence and (when `--play` is passed) plays
//! it back. Asserts that the engine returns non-empty PCM.
//!
//! Without assets: prints a hint and exits 0 (so CI can still run this binary
//! without the model files).
//!
//! Usage:
//!   test-tts                             Norwegian smoke test
//!   test-tts --lang en                   English (requires espeak-ng + EN voice)
//!   test-tts --text "Egen tekst her"     Override the sentence
//!   test-tts --play                      Also play the synthesized audio

use std::path::PathBuf;

use acatts_rust::downloader;
use norsk_g2p::synthesis::{EnglishVoice, NorwegianVoice};

/// Mirrors `crate::tts::piper_data_root()` in the GUI binary so the smoke test
/// reads from the same location the app downloads into. The previous
/// `dirs::data_dir().join("NorskTale").join("piper")` path was a leftover from
/// the old branding and pointed at a directory the app never writes to.
fn piper_data_root() -> PathBuf {
    downloader::data_dir().join("piper")
}

#[cfg(target_os = "windows")]
fn espeak_path(root: &std::path::Path) -> PathBuf {
    root.join("bin").join("espeak-ng.exe")
}
#[cfg(not(target_os = "windows"))]
fn espeak_path(root: &std::path::Path) -> PathBuf {
    root.join("bin").join("espeak-ng")
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let mut lang = "nb";
    let mut text: Option<String> = None;
    let mut play = false;

    let mut i = 1;
    while i < args.len() {
        match args[i].as_str() {
            "--lang" => {
                i += 1;
                if i < args.len() {
                    lang = match args[i].as_str() {
                        "nb" | "nn" | "no" => "nb",
                        "en" => "en",
                        other => {
                            eprintln!("Unknown --lang: {}", other);
                            std::process::exit(2);
                        }
                    };
                }
            }
            "--text" => {
                i += 1;
                if i < args.len() {
                    text = Some(args[i].clone());
                }
            }
            "--play" => play = true,
            "-h" | "--help" => {
                println!("Piper TTS smoke test. See file docs.");
                return;
            }
            other => {
                eprintln!("Unknown arg: {}", other);
                std::process::exit(2);
            }
        }
        i += 1;
    }

    let root = piper_data_root();
    eprintln!("Piper data root: {}", root.display());

    match lang {
        "nb" => smoke_norwegian(&root, text.as_deref(), play),
        "en" => smoke_english(&root, text.as_deref(), play),
        _ => unreachable!(),
    }
}

fn smoke_norwegian(root: &std::path::Path, text: Option<&str>, play: bool) {
    let base = root.join("nb-NO");
    let model = base.join("epoch_649_v5.onnx");
    if !model.exists() {
        eprintln!(
            "Norwegian Piper assets missing at {} — run the desktop app, pick \
             Bokmål or Nynorsk, and let the asset downloader complete first.",
            base.display()
        );
        std::process::exit(0);
    }

    let mut voice = match NorwegianVoice::load(
        &model,
        &base.join("epoch_649_v5.onnx.json"),
        &base.join("lexicon.fst"),
        &base.join("lexicon_values.bin"),
        &base.join("lexicon_phonemes.txt"),
        &base.join("pronunciation_overrides.tsv"),
    ) {
        Ok(v) => v,
        Err(e) => {
            eprintln!("Norwegian voice load failed: {}", e);
            std::process::exit(1);
        }
    };

    let utterance = text.unwrap_or("Hei, dette er en test av norsk talesyntese.");
    eprintln!("Synthesizing: {:?}", utterance);
    let result = voice.synthesize(utterance).expect("synthesize");
    assert!(!result.samples.is_empty(), "PCM was empty");
    eprintln!(
        "OK: {} samples ({:.2}s) at {} Hz",
        result.samples.len(),
        result.samples.len() as f32 / result.sample_rate as f32,
        result.sample_rate
    );

    if play {
        play_samples(&result.samples, result.sample_rate);
    }
}

fn smoke_english(root: &std::path::Path, text: Option<&str>, play: bool) {
    let voice_id = "en_US-lessac-medium";
    let base = root.join(voice_id);
    let model = base.join(format!("{}.onnx", voice_id));
    if !model.exists() {
        eprintln!(
            "English Piper assets missing at {} — pick English in the desktop \
             app and let the asset downloader complete first.",
            base.display()
        );
        std::process::exit(0);
    }
    let espeak = espeak_path(root);
    if !espeak.exists() {
        eprintln!(
            "espeak-ng binary missing at {} — should be downloaded alongside \
             English Piper assets.",
            espeak.display()
        );
        std::process::exit(0);
    }

    let mut voice = match EnglishVoice::load(
        &model,
        &base.join(format!("{}.onnx.json", voice_id)),
        espeak.clone(),
        "en-us".to_string(),
    ) {
        Ok(v) => v,
        Err(e) => {
            eprintln!("English voice load failed: {}", e);
            std::process::exit(1);
        }
    };

    let utterance = text.unwrap_or("Hello, this is a test of English speech synthesis.");
    eprintln!("Synthesizing: {:?}", utterance);
    let result = voice.synthesize(utterance).expect("synthesize");
    assert!(!result.samples.is_empty(), "PCM was empty");
    eprintln!(
        "OK: {} samples ({:.2}s) at {} Hz",
        result.samples.len(),
        result.samples.len() as f32 / result.sample_rate as f32,
        result.sample_rate
    );

    if play {
        play_samples(&result.samples, result.sample_rate);
    }
}

fn play_samples(samples: &[f32], sample_rate: u32) {
    use norsk_g2p::inference::write_wav;
    let tmp = std::env::temp_dir().join(format!(
        "test-tts-{}.wav",
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0)
    ));
    if let Err(e) = write_wav(tmp.to_str().unwrap_or(""), samples, sample_rate) {
        eprintln!("write_wav failed: {}", e);
        return;
    }

    #[cfg(target_os = "macos")]
    {
        let _ = std::process::Command::new("afplay").arg(&tmp).status();
    }
    #[cfg(target_os = "windows")]
    {
        use std::os::windows::ffi::OsStrExt;
        type PlaySoundFn = unsafe extern "system" fn(*const u16, usize, u32) -> i32;
        const SND_FILENAME: u32 = 0x00020000;
        const SND_SYNC: u32 = 0x00000000;
        const SND_NODEFAULT: u32 = 0x00000002;
        let lib = unsafe { libloading::Library::new("winmm.dll") }.expect("winmm.dll");
        let play: libloading::Symbol<PlaySoundFn> =
            unsafe { lib.get(b"PlaySoundW") }.expect("PlaySoundW");
        let wide: Vec<u16> = tmp
            .as_os_str()
            .encode_wide()
            .chain(std::iter::once(0))
            .collect();
        unsafe {
            play(wide.as_ptr(), 0, SND_FILENAME | SND_SYNC | SND_NODEFAULT);
        }
    }

    let _ = std::fs::remove_file(&tmp);
}
