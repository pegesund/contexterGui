mod tts;

fn main() {
    // Test: play the same WAV file 5 times using PlaySoundW
    // If this sounds random, it's a Windows audio issue
    let wav_path = format!("{}\\test_1.wav", std::env::temp_dir().display());
    println!("Playing {} five times with PlaySoundW...", wav_path);

    let wav_data = std::fs::read(&wav_path).expect("Could not read test_1.wav");
    println!("WAV size: {} bytes", wav_data.len());

    // Load PlaySoundW
    use std::ffi::OsStr;
    use std::os::windows::ffi::OsStrExt;

    type PlaySoundFn = unsafe extern "system" fn(*const u8, usize, u32) -> i32;
    const SND_MEMORY: u32 = 0x0004;
    const SND_SYNC: u32 = 0x0000;
    const SND_NODEFAULT: u32 = 0x0002;
    const SND_FILENAME: u32 = 0x00020000;

    let winmm = unsafe { libloading::Library::new("winmm.dll") }.expect("winmm.dll");
    let play_sound: PlaySoundFn = unsafe {
        let f: libloading::Symbol<PlaySoundFn> = winmm.get(b"PlaySoundW").unwrap();
        *f
    };

    // Test A: Play from memory (SND_MEMORY)
    println!("\n--- Test A: PlaySoundW with SND_MEMORY ---");
    for i in 1..=5 {
        println!("  Play #{}", i);
        unsafe { play_sound(wav_data.as_ptr(), 0, SND_MEMORY | SND_SYNC | SND_NODEFAULT); }
        std::thread::sleep(std::time::Duration::from_millis(300));
    }

    // Test B: Play from file (SND_FILENAME)
    println!("\n--- Test B: PlaySoundW with SND_FILENAME ---");
    let wide: Vec<u16> = OsStr::new(&wav_path).encode_wide().chain(std::iter::once(0)).collect();
    for i in 1..=5 {
        println!("  Play #{}", i);
        unsafe { play_sound(wide.as_ptr() as *const u8, 0, SND_FILENAME | SND_SYNC | SND_NODEFAULT); }
        std::thread::sleep(std::time::Duration::from_millis(300));
    }

    println!("\nDone!");
}
