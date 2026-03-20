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
    /// Last foreground app that was NOT our app — used for reading selected text
    last_external_fg: Arc<Mutex<ForegroundApp>>,
    /// Cached selected text from last external app — read while that app is still in focus
    cached_selected_text: Arc<Mutex<Option<String>>>,
    screen: (f32, f32),
}

impl MacPlatform {
    pub fn new() -> Self {
        let cached_fg = Arc::new(Mutex::new((ForegroundApp::default(), Instant::now())));
        let last_external_fg = Arc::new(Mutex::new(ForegroundApp::default()));
        let cached_selected_text = Arc::new(Mutex::new(None));

        // Background thread polls foreground app every 200ms
        let fg_clone = Arc::clone(&cached_fg);
        let ext_clone = Arc::clone(&last_external_fg);
        let sel_clone = Arc::clone(&cached_selected_text);
        std::thread::Builder::new()
            .name("fg-poller".into())
            .spawn(move || {
                loop {
                    if let Some(app) = query_foreground_app() {
                        let is_our_app = app.exe_name == "acatts-rust"
                            || app.exe_name == "norsktale"
                            || app.title == "NorskTale";
                        if !is_our_app {
                            if let Ok(mut lock) = ext_clone.lock() {
                                *lock = app.clone();
                            }
                            // Read selected text while external app still has focus
                            let sel = read_selected_text_system_wide();
                            if let Ok(mut lock) = sel_clone.lock() {
                                *lock = sel;
                            }
                        }
                        if let Ok(mut lock) = fg_clone.lock() {
                            *lock = (app, Instant::now());
                        }
                    }
                    std::thread::sleep(Duration::from_millis(200));
                }
            })
            .expect("Failed to spawn foreground poller");

        let screen = query_screen_size().unwrap_or((1920.0, 1080.0));

        MacPlatform { cached_fg, last_external_fg, cached_selected_text, screen }
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

    fn read_selected_text(&self) -> Option<String> {
        // Return the cached selected text (read by poller while external app had focus)
        if let Ok(lock) = self.cached_selected_text.lock() {
            lock.clone()
        } else {
            None
        }
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

/// Read selected text from a specific application by PID using macOS Accessibility API.
/// If pid is 0, reads from the system-wide focused application.
/// Requires Accessibility permission in System Settings > Privacy > Accessibility.
fn read_selected_text_from_pid(pid: u32) -> Option<String> {
    use accessibility_sys::*;
    use core_foundation::base::{CFRelease, CFTypeRef, TCFType};
    use core_foundation::string::CFString;

    unsafe {
        let app_element = if pid > 0 {
            // Target a specific app by PID
            AXUIElementCreateApplication(pid as i32)
        } else {
            // Fall back to system-wide focused app
            let system_wide = AXUIElementCreateSystemWide();
            let attr_focused_app = CFString::from_static_string("AXFocusedApplication");
            let mut focused_app: CFTypeRef = std::ptr::null();
            let err = AXUIElementCopyAttributeValue(
                system_wide,
                attr_focused_app.as_concrete_TypeRef() as _,
                &mut focused_app,
            );
            CFRelease(system_wide as _);
            if err != 0 || focused_app.is_null() {
                return None;
            }
            focused_app as AXUIElementRef
        };

        // Get the focused UI element within that app
        let attr_focused_elem = CFString::from_static_string("AXFocusedUIElement");
        let mut focused_elem: CFTypeRef = std::ptr::null();
        let err = AXUIElementCopyAttributeValue(
            app_element,
            attr_focused_elem.as_concrete_TypeRef() as _,
            &mut focused_elem,
        );
        CFRelease(app_element as _);
        if err != 0 || focused_elem.is_null() {
            return None;
        }

        // Get the selected text
        let attr_selected = CFString::from_static_string("AXSelectedText");
        let mut selected_text: CFTypeRef = std::ptr::null();
        let err = AXUIElementCopyAttributeValue(
            focused_elem as AXUIElementRef,
            attr_selected.as_concrete_TypeRef() as _,
            &mut selected_text,
        );
        CFRelease(focused_elem);
        if err != 0 || selected_text.is_null() {
            return None;
        }

        // Convert CFStringRef to Rust String
        let cf_str = CFString::wrap_under_create_rule(selected_text as _);
        let result = cf_str.to_string();
        if result.is_empty() {
            None
        } else {
            Some(result)
        }
    }
}

/// Read selected text from the currently focused app via AppleScript + System Events.
/// Called from the poller thread while the external app still has focus.
/// Uses System Events which has its own accessibility permissions.
/// Read selected text from the currently focused app using the Accessibility C API.
/// Supports both apps with AXSelectedText (Word, TextEdit) and apps using
/// text markers (Safari). Called from the poller thread while external app has focus.
fn read_selected_text_system_wide() -> Option<String> {
    use accessibility_sys::*;
    use core_foundation::base::{CFRelease, CFTypeRef, TCFType};
    use core_foundation::string::CFString;

    unsafe {
        let system_wide = AXUIElementCreateSystemWide();

        // Get focused application
        let attr_focused_app = CFString::from_static_string("AXFocusedApplication");
        let mut focused_app: CFTypeRef = std::ptr::null();
        let err = AXUIElementCopyAttributeValue(
            system_wide, attr_focused_app.as_concrete_TypeRef() as _, &mut focused_app,
        );
        CFRelease(system_wide as _);
        if err != 0 || focused_app.is_null() { return None; }

        // Get focused UI element
        let attr_focused_elem = CFString::from_static_string("AXFocusedUIElement");
        let mut focused_elem: CFTypeRef = std::ptr::null();
        let err = AXUIElementCopyAttributeValue(
            focused_app as AXUIElementRef, attr_focused_elem.as_concrete_TypeRef() as _, &mut focused_elem,
        );
        CFRelease(focused_app);
        if err != 0 || focused_elem.is_null() { return None; }

        let elem = focused_elem as AXUIElementRef;

        // Try 1: AXSelectedText (works for Word, TextEdit, etc.)
        let attr_selected = CFString::from_static_string("AXSelectedText");
        let mut selected_text: CFTypeRef = std::ptr::null();
        let err = AXUIElementCopyAttributeValue(
            elem, attr_selected.as_concrete_TypeRef() as _, &mut selected_text,
        );
        if err == 0 && !selected_text.is_null() {
            let cf_str = CFString::wrap_under_create_rule(selected_text as _);
            let result = cf_str.to_string();
            CFRelease(focused_elem);
            if !result.is_empty() { return Some(result); }
            return None;
        }

        // Try 2: AXSelectedTextMarkerRange + AXStringForTextMarkerRange (Safari, web views)
        let attr_marker_range = CFString::from_static_string("AXSelectedTextMarkerRange");
        let mut marker_range: CFTypeRef = std::ptr::null();
        let err = AXUIElementCopyAttributeValue(
            elem, attr_marker_range.as_concrete_TypeRef() as _, &mut marker_range,
        );
        if err == 0 && !marker_range.is_null() {
            let attr_string_for_range = CFString::from_static_string("AXStringForTextMarkerRange");
            let mut result_text: CFTypeRef = std::ptr::null();
            let err = AXUIElementCopyParameterizedAttributeValue(
                elem, attr_string_for_range.as_concrete_TypeRef() as _, marker_range, &mut result_text,
            );
            CFRelease(marker_range);
            CFRelease(focused_elem);
            if err == 0 && !result_text.is_null() {
                let cf_str = CFString::wrap_under_create_rule(result_text as _);
                let result = cf_str.to_string();
                if !result.is_empty() { return Some(result); }
            }
            return None;
        }

        CFRelease(focused_elem);
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
