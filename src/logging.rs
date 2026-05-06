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

/// Bridge that routes the standard `log` crate (used by velopack and other
/// libraries) into our LOG_FILE. Without this, every `log::info!()` /
/// `log::warn!()` / `log::error!()` from a dependency is silently dropped
/// and we have no visibility into update-check failures.
struct FileLogger;

impl log::Log for FileLogger {
    fn enabled(&self, _meta: &log::Metadata) -> bool {
        true
    }

    fn log(&self, record: &log::Record) {
        if let Ok(mut f) = LOG_FILE.lock() {
            let _ = writeln!(
                f,
                "[{} {}] {}",
                record.level(),
                record.target(),
                record.args()
            );
            let _ = f.flush();
        }
    }

    fn flush(&self) {
        if let Ok(mut f) = LOG_FILE.lock() {
            let _ = f.flush();
        }
    }
}

static LOGGER: FileLogger = FileLogger;

/// Install the forwarder. Call once from main() before any code that uses
/// the standard log crate (e.g. velopack). Idempotent — set_logger only
/// succeeds the first time, subsequent calls are a no-op.
pub fn install_log_forwarder() {
    let _ = log::set_logger(&LOGGER);
    log::set_max_level(log::LevelFilter::Info);
}
