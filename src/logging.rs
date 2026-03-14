use std::io::Write;
use std::sync::Mutex;

pub static LOG_FILE: std::sync::LazyLock<Mutex<std::fs::File>> = std::sync::LazyLock::new(|| {
    let path = std::env::temp_dir().join("acatts-rust.log");
    eprintln!("Logging to: {}", path.display());
    let f = std::fs::OpenOptions::new()
        .create(true).write(true).truncate(true)
        .open(&path).expect("failed to open log file");
    Mutex::new(f)
});

#[macro_export]
macro_rules! log {
    ($($arg:tt)*) => {{
        let msg = format!($($arg)*);
        if let Ok(mut f) = $crate::logging::LOG_FILE.lock() {
            use std::io::Write;
            let _ = writeln!(f, "{}", msg);
            let _ = f.flush();
        }
    }};
}
