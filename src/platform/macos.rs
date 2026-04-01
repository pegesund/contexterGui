use super::{AppKind, ForegroundApp, PlatformServices};
use std::process::Command;
use std::sync::{Arc, Mutex};
use std::sync::atomic::{AtomicBool, Ordering};
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
    intercept_tab: Arc<AtomicBool>,
    tab_pressed: Arc<AtomicBool>,
}

impl MacPlatform {
    pub fn new() -> Self {
        // Request accessibility permissions (prompts user if not granted)
        unsafe {
            use core_foundation::base::TCFType;
            use core_foundation::boolean::CFBoolean;
            use core_foundation::string::CFString;
            use core_foundation::dictionary::CFDictionary;
            unsafe extern "C" {
                fn AXIsProcessTrustedWithOptions(options: core_foundation::base::CFTypeRef) -> bool;
            }
            let key = CFString::new("AXTrustedCheckOptionPrompt");
            let dict = CFDictionary::from_CFType_pairs(&[(key.as_CFType(), CFBoolean::true_value().as_CFType())]);
            let trusted = AXIsProcessTrustedWithOptions(dict.as_concrete_TypeRef() as _);
            {
                use std::io::Write;
                if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true).open("/tmp/acatts_accessibility.log") {
                    let _ = writeln!(f, "Accessibility: trusted={}", trusted);
                }
            }
        }

        let cached_fg = Arc::new(Mutex::new((ForegroundApp::default(), Instant::now())));
        let last_external_fg = Arc::new(Mutex::new(ForegroundApp::default()));
        let cached_selected_text = Arc::new(Mutex::new(None));
        let intercept_tab = Arc::new(AtomicBool::new(false));
        let tab_pressed = Arc::new(AtomicBool::new(false));
        { let i = Arc::clone(&intercept_tab); let p = Arc::clone(&tab_pressed);
          std::thread::Builder::new().name("tab-tap".into()).spawn(move || start_tab_event_tap(i, p)).ok(); }

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
                            let sel = read_selected_text_for_app(app.pid, &app.exe_name);
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

        MacPlatform { cached_fg, last_external_fg, cached_selected_text, screen, intercept_tab, tab_pressed }
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

    fn check_hotkey_state(&self) -> (bool, bool) { (false, false) }
    fn set_tab_intercept(&self, active: bool) { self.intercept_tab.store(active, Ordering::Relaxed); }
    fn take_tab_press(&self) -> bool { self.tab_pressed.swap(false, Ordering::Relaxed) }

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

    fn caret_screen_position(&self) -> Option<(i32, i32)> {
        use accessibility_sys::*;
        use core_foundation::base::{CFRelease, TCFType};
        use core_foundation::string::CFString;

        unsafe {
            // Get the system-wide accessibility element
            let sys = AXUIElementCreateSystemWide();
            let key = CFString::new("AXFocusedUIElement");
            let mut focused: core_foundation::base::CFTypeRef = std::ptr::null();
            let err = AXUIElementCopyAttributeValue(sys, key.as_concrete_TypeRef(), &mut focused);
            CFRelease(sys as _);
            if err != 0 || focused.is_null() { return None; }

            // Get AXSelectedTextRange
            let range_key = CFString::new("AXSelectedTextRange");
            let mut range_val: core_foundation::base::CFTypeRef = std::ptr::null();
            let err = AXUIElementCopyAttributeValue(focused as AXUIElementRef, range_key.as_concrete_TypeRef(), &mut range_val);
            if err != 0 || range_val.is_null() {
                CFRelease(focused);
                return None;
            }

            // Get AXBoundsForRange for the selection — gives us screen coordinates
            let bounds_key = CFString::new("AXBoundsForRange");
            let mut bounds_val: core_foundation::base::CFTypeRef = std::ptr::null();
            let err = AXUIElementCopyParameterizedAttributeValue(
                focused as AXUIElementRef,
                bounds_key.as_concrete_TypeRef(),
                range_val,
                &mut bounds_val,
            );
            CFRelease(range_val);
            CFRelease(focused);

            if err != 0 || bounds_val.is_null() { return None; }

            // bounds_val is an AXValue containing a CGRect
            let mut rect = core_graphics::geometry::CGRect::new(
                &core_graphics::geometry::CGPoint::new(0.0, 0.0),
                &core_graphics::geometry::CGSize::new(0.0, 0.0),
            );
            let ok = AXValueGetValue(
                bounds_val as AXValueRef,
                kAXValueTypeCGRect,
                &mut rect as *mut _ as *mut std::ffi::c_void,
            );
            CFRelease(bounds_val);

            if !ok { return None; }

            // rect.origin is top-left of the selection in screen coords
            // Return bottom of selection (where we want the window)
            Some((rect.origin.x as i32, (rect.origin.y + rect.size.height) as i32))
        }
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

        // Get the app's PID for tree walking later
        let mut app_pid: i32 = 0;
        AXUIElementGetPid(focused_app as AXUIElementRef, &mut app_pid);

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

        // Check the role of the focused element — only walk tree for non-web elements
        let mut role: CFTypeRef = std::ptr::null();
        AXUIElementCopyAttributeValue(elem, CFString::from_static_string("AXRole").as_concrete_TypeRef() as _, &mut role);
        let role_str = if !role.is_null() {
            let s = CFString::wrap_under_get_rule(role as _).to_string();
            s
        } else {
            String::new()
        };

        CFRelease(focused_elem);

        // Try 3: Walk tree for apps like Word where focused element isn't the text area
        if app_pid > 0 {
            let app_elem = AXUIElementCreateApplication(app_pid);
            let result = walk_for_selected_text(app_elem, 0);
            CFRelease(app_elem as _);
            return result;
        }
        None
    }
}

/// Walk the accessibility tree to find elements with AXSelectedText.
/// Word exposes it on AXTextArea inside AXLayoutArea → AXSelectedChildren.
unsafe fn walk_for_selected_text(element: accessibility_sys::AXUIElementRef, depth: i32) -> Option<String> {
    use accessibility_sys::*;
    use core_foundation::base::{CFRelease, CFTypeRef, TCFType};
    use core_foundation::string::CFString;

    if depth > 6 { return None; }

    // Check AXSelectedChildren first (Word's AXLayoutArea has this)
    let attr_sel_children = CFString::from_static_string("AXSelectedChildren");
    let mut sel_children: CFTypeRef = std::ptr::null();
    let err = AXUIElementCopyAttributeValue(
        element, attr_sel_children.as_concrete_TypeRef() as _, &mut sel_children,
    );
    if err == 0 && !sel_children.is_null() {
        // sel_children is a CFArray
        let count = core_foundation::array::CFArrayGetCount(sel_children as _);
        for i in 0..count {
            let child = core_foundation::array::CFArrayGetValueAtIndex(sel_children as _, i) as AXUIElementRef;
            // Try AXSelectedText on this child
            let attr_selected = CFString::from_static_string("AXSelectedText");
            let mut selected_text: CFTypeRef = std::ptr::null();
            let err = AXUIElementCopyAttributeValue(
                child, attr_selected.as_concrete_TypeRef() as _, &mut selected_text,
            );
            if err == 0 && !selected_text.is_null() {
                let cf_str = CFString::wrap_under_create_rule(selected_text as _);
                let result = cf_str.to_string();
                CFRelease(sel_children);
                if !result.is_empty() { return Some(result); }
                return None;
            }
        }
        CFRelease(sel_children);
    }

    // Recurse into children
    let attr_children = CFString::from_static_string("AXChildren");
    let mut children: CFTypeRef = std::ptr::null();
    let err = AXUIElementCopyAttributeValue(
        element, attr_children.as_concrete_TypeRef() as _, &mut children,
    );
    if err == 0 && !children.is_null() {
        let count = core_foundation::array::CFArrayGetCount(children as _);
        for i in 0..count {
            let child = core_foundation::array::CFArrayGetValueAtIndex(children as _, i) as AXUIElementRef;
            if let Some(text) = walk_for_selected_text(child, depth + 1) {
                CFRelease(children);
                return Some(text);
            }
        }
        CFRelease(children);
    }

    None
}

/// Read selected text from the given app, trying multiple strategies.
fn read_selected_text_for_app(pid: u32, _app_name: &str) -> Option<String> {
    use accessibility_sys::*;
    use core_foundation::base::{CFRelease, CFTypeRef, TCFType};
    use core_foundation::string::CFString;

    if pid == 0 { return None; }

    unsafe {
        let app_elem = AXUIElementCreateApplication(pid as i32);

        // Enable enhanced UI on the first window (needed for Chrome/Chromium)
        let mut windows: CFTypeRef = std::ptr::null();
        let werr = AXUIElementCopyAttributeValue(
            app_elem, CFString::from_static_string("AXWindows").as_concrete_TypeRef() as _, &mut windows,
        );
        if werr == 0 && !windows.is_null() {
            let count = core_foundation::array::CFArrayGetCount(windows as _);
            if count > 0 {
                let first_win = core_foundation::array::CFArrayGetValueAtIndex(windows as _, 0) as AXUIElementRef;
                // Ignore error — some apps don't support this, but it still triggers Chrome's a11y
                let _ = AXUIElementSetAttributeValue(
                    first_win,
                    CFString::from_static_string("AXEnhancedUserInterface").as_concrete_TypeRef() as _,
                    core_foundation::boolean::CFBoolean::true_value().as_concrete_TypeRef() as CFTypeRef,
                );
            }
            CFRelease(windows);
        }

        // Get focused UI element from this specific app
        let mut focused_elem: CFTypeRef = std::ptr::null();
        let err = AXUIElementCopyAttributeValue(
            app_elem, CFString::from_static_string("AXFocusedUIElement").as_concrete_TypeRef() as _, &mut focused_elem,
        );

        if err == 0 && !focused_elem.is_null() {
            let elem = focused_elem as AXUIElementRef;

            // Try AXSelectedText
            let mut sel: CFTypeRef = std::ptr::null();
            let e = AXUIElementCopyAttributeValue(
                elem, CFString::from_static_string("AXSelectedText").as_concrete_TypeRef() as _, &mut sel,
            );
            if e == 0 && !sel.is_null() {
                let s = CFString::wrap_under_create_rule(sel as _).to_string();
                CFRelease(focused_elem);
                CFRelease(app_elem as _);
                if !s.is_empty() { return Some(s); }
                return None;
            }

            // Try markers (Safari)
            let mut marker: CFTypeRef = std::ptr::null();
            let e = AXUIElementCopyAttributeValue(
                elem, CFString::from_static_string("AXSelectedTextMarkerRange").as_concrete_TypeRef() as _, &mut marker,
            );
            if e == 0 && !marker.is_null() {
                let mut result_text: CFTypeRef = std::ptr::null();
                let e = AXUIElementCopyParameterizedAttributeValue(
                    elem, CFString::from_static_string("AXStringForTextMarkerRange").as_concrete_TypeRef() as _, marker, &mut result_text,
                );
                CFRelease(marker);
                CFRelease(focused_elem);
                CFRelease(app_elem as _);
                if e == 0 && !result_text.is_null() {
                    let s = CFString::wrap_under_create_rule(result_text as _).to_string();
                    if !s.is_empty() { return Some(s); }
                }
                return None;
            }

            CFRelease(focused_elem);
        }

        // Tree walk fallback (Word)
        let result = walk_for_selected_text(app_elem, 0);
        CFRelease(app_elem as _);
        if result.is_some() { return result; }

        // System-wide fallback (Safari — AXFocusedApplication works differently)
        return read_selected_text_system_wide();
    }
}

/// Debug wrapper that logs each step to acatts-speak.log
fn read_selected_text_system_wide_debug(app_name: &str) -> Option<String> {
    use accessibility_sys::*;
    use core_foundation::base::{CFRelease, CFTypeRef, TCFType};
    use core_foundation::string::CFString;

    let log = |msg: &str| {
        if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true)
            .open(std::env::temp_dir().join("acatts-speak.log")) {
            use std::io::Write;
            let _ = writeln!(f, "[{}] {}", app_name, msg);
        }
    };

    unsafe {
        let system_wide = AXUIElementCreateSystemWide();
        let mut focused_app: CFTypeRef = std::ptr::null();
        let err = AXUIElementCopyAttributeValue(
            system_wide, CFString::from_static_string("AXFocusedApplication").as_concrete_TypeRef() as _, &mut focused_app,
        );
        CFRelease(system_wide as _);
        if err != 0 || focused_app.is_null() {
            log(&format!("AXFocusedApplication FAILED err={}", err));
            return None;
        }

        let mut app_pid: i32 = 0;
        AXUIElementGetPid(focused_app as AXUIElementRef, &mut app_pid);

        let mut focused_elem: CFTypeRef = std::ptr::null();
        let err = AXUIElementCopyAttributeValue(
            focused_app as AXUIElementRef, CFString::from_static_string("AXFocusedUIElement").as_concrete_TypeRef() as _, &mut focused_elem,
        );
        CFRelease(focused_app);
        if err != 0 || focused_elem.is_null() {
            log(&format!("AXFocusedUIElement FAILED err={}", err));
            // Still try tree walk
            let app_elem = AXUIElementCreateApplication(app_pid);
            let result = walk_for_selected_text(app_elem, 0);
            CFRelease(app_elem as _);
            if result.is_some() { log("tree walk FOUND text"); }
            return result;
        }

        let elem = focused_elem as AXUIElementRef;

        // Get role for logging
        let mut role: CFTypeRef = std::ptr::null();
        AXUIElementCopyAttributeValue(elem, CFString::from_static_string("AXRole").as_concrete_TypeRef() as _, &mut role);
        let role_str = if !role.is_null() { CFString::wrap_under_get_rule(role as _).to_string() } else { "?".into() };
        log(&format!("focused role={}", role_str));

        // Try 1: AXSelectedText
        let mut selected_text: CFTypeRef = std::ptr::null();
        let err = AXUIElementCopyAttributeValue(
            elem, CFString::from_static_string("AXSelectedText").as_concrete_TypeRef() as _, &mut selected_text,
        );
        if err == 0 && !selected_text.is_null() {
            let cf_str = CFString::wrap_under_create_rule(selected_text as _);
            let result = cf_str.to_string();
            CFRelease(focused_elem);
            if !result.is_empty() {
                log(&format!("AXSelectedText OK len={}", result.len()));
                return Some(result);
            }
            log("AXSelectedText empty");
            return None;
        }
        log(&format!("AXSelectedText FAILED err={}", err));

        // Try 2: markers
        let mut marker_range: CFTypeRef = std::ptr::null();
        let err = AXUIElementCopyAttributeValue(
            elem, CFString::from_static_string("AXSelectedTextMarkerRange").as_concrete_TypeRef() as _, &mut marker_range,
        );
        if err == 0 && !marker_range.is_null() {
            let mut result_text: CFTypeRef = std::ptr::null();
            let err = AXUIElementCopyParameterizedAttributeValue(
                elem, CFString::from_static_string("AXStringForTextMarkerRange").as_concrete_TypeRef() as _, marker_range, &mut result_text,
            );
            CFRelease(marker_range);
            CFRelease(focused_elem);
            if err == 0 && !result_text.is_null() {
                let cf_str = CFString::wrap_under_create_rule(result_text as _);
                let result = cf_str.to_string();
                if !result.is_empty() {
                    log(&format!("markers OK len={}", result.len()));
                    return Some(result);
                }
            }
            log(&format!("markers FAILED err={}", err));
            return None;
        }
        log(&format!("AXSelectedTextMarkerRange FAILED err={}", err));

        CFRelease(focused_elem);

        // Try 3: tree walk
        if app_pid > 0 {
            let app_elem = AXUIElementCreateApplication(app_pid);
            let result = walk_for_selected_text(app_elem, 0);
            CFRelease(app_elem as _);
            if result.is_some() { log("tree walk FOUND text"); } else { log("tree walk NOTHING"); }
            return result;
        }
        None
    }
}

/// AppleScript fallback for reading selected text. Works via System Events
/// which has its own accessibility permissions (doesn't need our app in the list).
fn read_selected_text_applescript(app_name: &str) -> Option<String> {
    // Capitalize app name for AppleScript process name
    let process_name = match app_name {
        "google chrome" => "Google Chrome",
        "safari" => "Safari",
        "microsoft word" => "Microsoft Word",
        "textedit" => "TextEdit",
        _ => return None,
    };
    let script = format!(
        r#"tell application "System Events"
            tell application process "{}"
                try
                    set focEl to value of attribute "AXFocusedUIElement"
                    return value of attribute "AXSelectedText" of focEl
                on error
                    return ""
                end try
            end tell
        end tell"#,
        process_name
    );
    let result = run_applescript(&script)?;
    if result.is_empty() { None } else { Some(result) }
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

fn start_tab_event_tap(intercept: Arc<AtomicBool>, pressed: Arc<AtomicBool>) {
    unsafe {
        unsafe extern "C" {
            fn CGEventTapCreate(t:u32,p:u32,o:u32,m:u64,cb:extern "C" fn(*const std::ffi::c_void,u32,*const std::ffi::c_void,*mut std::ffi::c_void)->*const std::ffi::c_void,i:*mut std::ffi::c_void)->*const std::ffi::c_void;
            fn CFMachPortCreateRunLoopSource(a:*const std::ffi::c_void,p:*const std::ffi::c_void,o:i64)->*const std::ffi::c_void;
            fn CFRunLoopGetCurrent()->*const std::ffi::c_void;
            fn CFRunLoopAddSource(r:*const std::ffi::c_void,s:*const std::ffi::c_void,m:*const std::ffi::c_void);
            fn CFRunLoopRun();
            fn CGEventGetIntegerValueField(e:*const std::ffi::c_void,f:u32)->i64;
            fn CGEventGetFlags(e:*const std::ffi::c_void)->u64;
            static kCFRunLoopCommonModes: *const std::ffi::c_void;
        }
        struct Ctx{i:Arc<AtomicBool>,p:Arc<AtomicBool>}
        let ctx=Box::into_raw(Box::new(Ctx{i:intercept,p:pressed})) as *mut std::ffi::c_void;
        extern "C" fn cb(_:*const std::ffi::c_void,et:u32,ev:*const std::ffi::c_void,ui:*mut std::ffi::c_void)->*const std::ffi::c_void{
            unsafe{
                let c=&*(ui as*const Ctx);
                if et!=10{return ev;}
                if !c.i.load(Ordering::Relaxed){return ev;}
                if CGEventGetIntegerValueField(ev,9)!=48{return ev;}
                // Pass through Tab if any modifier is held (Cmd+Tab, Ctrl+Tab, etc.)
                let flags = CGEventGetFlags(ev);
                let modifiers = flags & 0x1F0000; // Cmd, Shift, Ctrl, Option, Fn
                if modifiers != 0 { return ev; }
                c.p.store(true,Ordering::Relaxed);
                std::ptr::null()
            }
        }
        let tap=CGEventTapCreate(0,0,0,1<<10,cb,ctx);
        if tap.is_null(){return;}
        let src=CFMachPortCreateRunLoopSource(std::ptr::null(),tap,0);
        CFRunLoopAddSource(CFRunLoopGetCurrent(),src,kCFRunLoopCommonModes);
        CFRunLoopRun();
    }
}
