use super::{AppKind, ForegroundApp, PlatformServices};
use std::sync::{Arc, Mutex};

pub struct WindowsPlatform {
    /// Cached selected text — polled via UIA while external app has focus
    cached_selected_text: Arc<Mutex<Option<String>>>,
}

impl WindowsPlatform {
    pub fn new() -> Self {
        let cached_selected_text = Arc::new(Mutex::new(None));
        let sel_clone = Arc::clone(&cached_selected_text);

        // Background thread polls selected text via UIA every 200ms
        // (same pattern as Mac's fg-poller)
        std::thread::Builder::new()
            .name("sel-poller".into())
            .spawn(move || {
                // Initialize COM on this thread
                unsafe {
                    use windows::Win32::System::Com::*;
                    let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok();
                }
                loop {
                    // Only poll when an external app has focus (not our window)
                    let sel = poll_uia_selected_text();
                    if let Ok(mut lock) = sel_clone.lock() {
                        if sel.is_some() {
                            *lock = sel;
                        }
                        // Keep last known selection when our app gets focus
                    }
                    std::thread::sleep(std::time::Duration::from_millis(200));
                }
            })
            .expect("Failed to spawn selection poller");

        Self { cached_selected_text }
    }
}

/// Poll selected text from the focused UIA element using TextPattern.GetSelection.
/// Works for Word, Notepad, Chrome, and any UIA-compliant app.
/// Returns None if no text is selected or our app has focus.
fn poll_uia_selected_text() -> Option<String> {
    unsafe {
        use windows::Win32::System::Com::*;
        use windows::Win32::UI::Accessibility::*;
        use windows::Win32::UI::WindowsAndMessaging::*;
        use windows::core::Interface;

        // Skip if our own window has focus
        let fg = GetForegroundWindow();
        let mut fg_pid = 0u32;
        GetWindowThreadProcessId(fg, Some(&mut fg_pid));
        if fg_pid == std::process::id() {
            return None;
        }

        let uia: IUIAutomation =
            CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER).ok()?;
        let focused = uia.GetFocusedElement().ok()?;

        // Try TextPattern (works for most apps)
        if let Ok(pattern) =
            focused.GetCurrentPatternAs::<IUIAutomationTextPattern>(UIA_TextPatternId)
        {
            let selection = pattern.GetSelection().ok()?;
            let count = selection.Length().ok()?;
            if count > 0 {
                let range: IUIAutomationTextRange = selection.GetElement(0).ok()?;
                let text = range.GetText(-1).ok()?.to_string();
                let trimmed = text.trim().to_string();
                if !trimmed.is_empty() {
                    return Some(trimmed);
                }
            }
        }
        None
    }
}

#[cfg(target_os = "windows")]
impl PlatformServices for WindowsPlatform {
    fn init_runtime(&self) {
        unsafe {
            use windows::Win32::System::Com::*;
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok();
        }
    }

    fn foreground_app(&self) -> ForegroundApp {
        use windows::Win32::UI::WindowsAndMessaging::{
            GetForegroundWindow, GetWindowTextW, GetWindowThreadProcessId,
        };
        use windows::Win32::System::Threading::{
            OpenProcess, QueryFullProcessImageNameW, PROCESS_NAME_FORMAT,
            PROCESS_QUERY_LIMITED_INFORMATION,
        };

        let fg = unsafe { GetForegroundWindow() };
        let mut pid = 0u32;
        unsafe { GetWindowThreadProcessId(fg, Some(&mut pid)); }

        let mut buf = [0u16; 128];
        let len = unsafe { GetWindowTextW(fg, &mut buf) };
        let title = String::from_utf16_lossy(&buf[..len as usize]);

        let exe = if pid > 0 {
            if let Ok(handle) = unsafe {
                OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid)
            } {
                let mut exe_buf = [0u16; 260];
                let mut exe_len = exe_buf.len() as u32;
                if unsafe {
                    QueryFullProcessImageNameW(
                        handle,
                        PROCESS_NAME_FORMAT(0),
                        windows::core::PWSTR(exe_buf.as_mut_ptr()),
                        &mut exe_len,
                    )
                }
                .is_ok()
                {
                    let full = String::from_utf16_lossy(&exe_buf[..exe_len as usize]);
                    full.rsplit('\\').next().unwrap_or("").to_lowercase()
                } else {
                    String::new()
                }
            } else {
                String::new()
            }
        } else {
            String::new()
        };

        ForegroundApp {
            handle: fg.0 as isize,
            pid,
            title,
            exe_name: exe,
        }
    }

    fn classify_app(&self, app: &ForegroundApp) -> AppKind {
        if app.pid == std::process::id() {
            return AppKind::OurApp;
        }
        if app.exe_name == "winword.exe"
            || app.title.contains(".docx")
            || app.title.contains(".doc ")
        {
            return AppKind::Word;
        }
        if matches!(
            app.exe_name.as_str(),
            "chrome.exe" | "msedge.exe" | "firefox.exe" | "brave.exe" | "opera.exe" | "vivaldi.exe"
        ) {
            return AppKind::Browser;
        }
        if app.exe_name == "notepad.exe" {
            return AppKind::Notepad;
        }
        AppKind::Other
    }

    fn screen_size(&self) -> (f32, f32) {
        unsafe {
            use windows::Win32::UI::WindowsAndMessaging::*;
            let w = GetSystemMetrics(SM_CXSCREEN);
            let h = GetSystemMetrics(SM_CYSCREEN);
            (w as f32, h as f32)
        }
    }

    fn set_foreground(&self, handle: isize) {
        unsafe {
            use windows::Win32::Foundation::HWND;
            use windows::Win32::UI::WindowsAndMessaging::SetForegroundWindow;
            let hwnd = HWND(handle as *mut _);
            let _ = SetForegroundWindow(hwnd);
            std::thread::sleep(std::time::Duration::from_millis(100));
        }
    }

    fn check_hotkey_state(&self) -> (bool, bool) {
        use windows::Win32::UI::Input::KeyboardAndMouse::GetAsyncKeyState;
        let ctrl = unsafe { GetAsyncKeyState(0x11) } < 0;
        let space = unsafe { GetAsyncKeyState(0x20) } < 0;
        (ctrl, space)
    }

    fn copy_to_clipboard(&self, text: &str) {
        use windows::Win32::System::DataExchange::*;
        use windows::Win32::System::Memory::*;
        use windows::Win32::Foundation::HANDLE;
        unsafe {
            if OpenClipboard(None).is_ok() {
                let _ = EmptyClipboard();
                let wide: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
                let size = wide.len() * 2;
                let hmem = GlobalAlloc(GMEM_MOVEABLE, size);
                if let Ok(hmem) = hmem {
                    let ptr = GlobalLock(hmem);
                    if !ptr.is_null() {
                        std::ptr::copy_nonoverlapping(
                            wide.as_ptr() as *const u8,
                            ptr as *mut u8,
                            size,
                        );
                        GlobalUnlock(hmem).ok();
                        let _ = SetClipboardData(13, Some(HANDLE(hmem.0))); // CF_UNICODETEXT
                    }
                }
                let _ = CloseClipboard();
            }
        }
    }

    fn emoji_font_path(&self) -> Option<&str> {
        Some("C:/Windows/Fonts/seguiemj.ttf")
    }

    fn ort_dylib_candidates(&self) -> Vec<String> {
        vec![
            concat!(env!("CARGO_MANIFEST_DIR"), "/../../onnxruntime/onnxruntime-win-x64-1.23.0/lib/onnxruntime.dll").to_string(),
            "C:\\Windows\\System32\\onnxruntime.dll".to_string(),
        ]
    }

    fn swipl_path(&self) -> &str {
        "C:/Program Files/swipl/bin/libswipl.dll"
    }

    fn read_selected_text(&self) -> Option<String> {
        if let Ok(lock) = self.cached_selected_text.lock() {
            lock.clone()
        } else {
            None
        }
    }

    fn init_tts(&self) {
        if let Some(home) = std::env::var_os("USERPROFILE") {
            let sdk_dir = std::path::Path::new(&home)
                .join("Downloads/Sdk-Amul-Cogni-TTS-WIN_14-000_AIO");
            if sdk_dir.exists() {
                crate::tts::init_tts(sdk_dir.to_str().unwrap(), "Kari22k_NV");
            } else {
                eprintln!("Acapela SDK not found at {:?}", sdk_dir);
            }
        }
    }
}
