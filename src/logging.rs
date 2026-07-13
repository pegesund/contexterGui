use std::io::Write;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant};

static DEBUG_LOGGING: AtomicBool = AtomicBool::new(false);

pub static LOG_FILE: std::sync::LazyLock<Mutex<std::fs::File>> = std::sync::LazyLock::new(|| {
    let path = std::env::temp_dir().join("spell.log");
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

#[macro_export]
macro_rules! debug_log {
    ($($arg:tt)*) => {{
        if $crate::logging::debug_logging_enabled() {
            $crate::log!($($arg)*);
        }
    }};
}

pub fn configure_debug_logging(enabled: bool) {
    DEBUG_LOGGING.store(enabled, Ordering::Relaxed);
    log::set_max_level(configured_level(enabled));
}

pub fn debug_logging_enabled() -> bool {
    DEBUG_LOGGING.load(Ordering::Relaxed)
}

fn configured_level(debug_enabled: bool) -> log::LevelFilter {
    if debug_enabled {
        log::LevelFilter::Debug
    } else {
        log::LevelFilter::Info
    }
}

pub struct LogThrottle {
    last_emit: Mutex<Option<Instant>>,
}

impl LogThrottle {
    pub const fn new() -> Self {
        Self {
            last_emit: Mutex::new(None),
        }
    }

    pub fn should_emit(&self, interval: Duration) -> bool {
        self.should_emit_at(Instant::now(), interval)
    }

    fn should_emit_at(&self, now: Instant, interval: Duration) -> bool {
        let Ok(mut last_emit) = self.last_emit.lock() else {
            return false;
        };
        if last_emit
            .as_ref()
            .is_some_and(|last| now.duration_since(*last) < interval)
        {
            return false;
        }
        *last_emit = Some(now);
        true
    }
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
    log::set_max_level(configured_level(debug_logging_enabled()));
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normal_logging_keeps_info_without_dependency_debug_records() {
        assert_eq!(configured_level(false), log::LevelFilter::Info);
    }

    #[test]
    fn debug_logging_enables_dependency_debug_records() {
        assert_eq!(configured_level(true), log::LevelFilter::Debug);
    }

    #[test]
    fn throttle_emits_first_event_and_then_respects_interval() {
        let throttle = LogThrottle::new();
        let start = Instant::now();
        let interval = Duration::from_secs(2);

        assert!(throttle.should_emit_at(start, interval));
        assert!(!throttle.should_emit_at(start + Duration::from_secs(1), interval));
        assert!(throttle.should_emit_at(start + interval, interval));
    }
}
