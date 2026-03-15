use super::{AppKind, ForegroundApp, PlatformServices};
use std::process::Command;
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

/// macOS platform services.
///
/// Foreground app detection runs on a background thread to avoid
/// blocking the egui UI thread with `osascript` subprocess calls.
pub struct MacPlatform {
    cached_fg: Arc<Mutex<(ForegroundApp, Instant)>>,
    screen: (f32, f32),
}

impl MacPlatform {
    pub fn new() -> Self {
        let cached_fg = Arc::new(Mutex::new((ForegroundApp::default(), Instant::now())));

        // Background thread polls foreground app every 200ms
        let fg_clone = Arc::clone(&cached_fg);
        std::thread::Builder::new()
            .name("fg-poller".into())
            .spawn(move || loop {
                if let Some(app) = query_foreground_app() {
                    if let Ok(mut lock) = fg_clone.lock() {
                        *lock = (app, Instant::now());
                    }
                }
                std::thread::sleep(Duration::from_millis(200));
            })
            .expect("Failed to spawn foreground poller");

        let screen = query_screen_size().unwrap_or((1920.0, 1080.0));

        MacPlatform { cached_fg, screen }
    }
}

impl PlatformServices for MacPlatform {
    fn init_runtime(&self) {}

    fn foreground_app(&self) -> ForegroundApp {
        if let Ok(lock) = self.cached_fg.lock() {
            lock.0.clone()
        } else {
            ForegroundApp::default()
        }
    }

    fn classify_app(&self, app: &ForegroundApp) -> AppKind {
        if app.pid == std::process::id() {
            return AppKind::OurApp;
        }
        let name = app.exe_name.as_str();
        if name == "microsoft word" || (name.contains("word") && app.title.contains(".docx")) {
            return AppKind::Word;
        }
        if matches!(
            name,
            "google chrome" | "microsoft edge" | "safari" | "firefox"
                | "brave browser" | "opera" | "vivaldi" | "arc"
        ) {
            return AppKind::Browser;
        }
        if name == "textedit" {
            return AppKind::Notepad;
        }
        AppKind::Other
    }

    fn screen_size(&self) -> (f32, f32) {
        self.screen
    }

    fn set_foreground(&self, handle: isize) {
        let pid = handle as u32;
        let script = format!(
            r#"tell application "System Events"
                set frontProcess to first application process whose unix id is {}
                set frontmost of frontProcess to true
            end tell"#,
            pid
        );
        let _ = run_applescript(&script);
        std::thread::sleep(Duration::from_millis(100));
    }

    fn check_hotkey_state(&self) -> (bool, bool) {
        // TODO: Use CGEventSourceKeyState via core-graphics crate
        (false, false)
    }

    fn copy_to_clipboard(&self, text: &str) {
        let _ = Command::new("pbcopy")
            .stdin(std::process::Stdio::piped())
            .spawn()
            .and_then(|mut child| {
                use std::io::Write;
                if let Some(stdin) = child.stdin.as_mut() {
                    stdin.write_all(text.as_bytes())?;
                }
                child.wait()
            });
    }

    fn emoji_font_path(&self) -> Option<&str> {
        Some("/System/Library/Fonts/Apple Color Emoji.ttc")
    }

    fn ort_dylib_candidates(&self) -> Vec<String> {
        vec![
            concat!(env!("CARGO_MANIFEST_DIR"), "/../../onnxruntime/lib/libonnxruntime.dylib").to_string(),
            "/usr/local/lib/libonnxruntime.dylib".to_string(),
            "/opt/homebrew/lib/libonnxruntime.dylib".to_string(),
        ]
    }

    fn swipl_path(&self) -> &str {
        "/Applications/SWI-Prolog.app/Contents/Frameworks/libswipl.dylib"
    }

    fn init_tts(&self) {
        // macOS TTS uses `say` command — no init needed, always available
        crate::tts::init_tts("", "");
    }
}

// ── Helpers (run on background thread) ──

fn run_applescript(script: &str) -> Option<String> {
    let output = Command::new("osascript")
        .arg("-e")
        .arg(script)
        .output()
        .ok()?;
    if output.status.success() {
        Some(String::from_utf8_lossy(&output.stdout).trim().to_string())
    } else {
        None
    }
}

fn query_foreground_app() -> Option<ForegroundApp> {
    let script = r#"tell application "System Events"
        set frontApp to first application process whose frontmost is true
        set appName to name of frontApp
        set appPID to unix id of frontApp
        set winTitle to ""
        try
            set winTitle to name of front window of frontApp
        end try
        return appName & "|||" & appPID & "|||" & winTitle
    end tell"#;

    let result = run_applescript(script)?;
    let parts: Vec<&str> = result.splitn(3, "|||").collect();
    if parts.len() >= 2 {
        let app_name = parts[0].to_string();
        let pid: u32 = parts[1].parse().unwrap_or(0);
        let title = parts.get(2).unwrap_or(&"").to_string();
        Some(ForegroundApp {
            handle: pid as isize,
            pid,
            title,
            exe_name: app_name.to_lowercase(),
        })
    } else {
        None
    }
}

fn query_screen_size() -> Option<(f32, f32)> {
    let script = r#"tell application "Finder"
        set screenBounds to bounds of window of desktop
        return (item 3 of screenBounds as text) & "," & (item 4 of screenBounds as text)
    end tell"#;
    let result = run_applescript(script)?;
    let parts: Vec<&str> = result.split(',').collect();
    if parts.len() == 2 {
        if let (Ok(w), Ok(h)) = (parts[0].trim().parse::<f32>(), parts[1].trim().parse::<f32>()) {
            return Some((w, h));
        }
    }
    None
}
