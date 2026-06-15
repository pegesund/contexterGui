use super::{AppKind, ForegroundApp, PlatformServices};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

const ONNX_RUNTIME_VERSION: &str = "1.24.4";

/// Milliseconds since the user's last keyboard or mouse input, system-wide.
///
/// Used to gate the high-frequency UIA / COM polling that otherwise runs
/// at 5–10 Hz forever. When the user leaves their laptop idle for an
/// hour, those polls amount to tens of thousands of cross-process round
/// trips into whatever app is foreground (typically Word) — UIA provider
/// state accumulates in the target process and the whole desktop slows
/// down. Pausing the polls when there's been no input for a while keeps
/// the system fresh.
///
/// `GetLastInputInfo` is itself a cheap kernel call (~µs) that doesn't
/// allocate handles, so calling it on every poll iteration is free.
pub fn idle_millis() -> u32 {
    use windows::Win32::System::SystemInformation::GetTickCount;
    use windows::Win32::UI::Input::KeyboardAndMouse::{GetLastInputInfo, LASTINPUTINFO};
    unsafe {
        let mut info = LASTINPUTINFO {
            cbSize: std::mem::size_of::<LASTINPUTINFO>() as u32,
            dwTime: 0,
        };
        if GetLastInputInfo(&mut info).as_bool() {
            GetTickCount().saturating_sub(info.dwTime)
        } else {
            0
        }
    }
}

/// When running from a packaged Spell-windows-x64 zip extraction, return the
/// absolute path to the bundled `Frameworks` directory. Returns `None`
/// outside a packaged install.
pub fn bundled_frameworks_dir() -> Option<PathBuf> {
    let exe = std::env::current_exe().ok()?;
    let exe_dir = exe.parent()?;
    let dir = exe_dir.join("Frameworks");
    dir.is_dir().then_some(dir)
}

/// Return the absolute path to a DLL bundled at
/// `<Spell.exe>/Frameworks/<name>`. The Frameworks/Resources layout mirrors
/// the Mac .app bundle so nostos-cognio's swipl-home lookup
/// (`<dylib>/../Resources/swipl`) resolves to `Resources/swipl/` on both
/// platforms.
pub fn bundled_dll_path(name: &str) -> Option<PathBuf> {
    let path = bundled_frameworks_dir()?.join(name);
    path.exists().then_some(path)
}

fn bundled_dll(name: &str) -> Option<String> {
    bundled_dll_path(name).map(|path| path_to_swi_string(&path))
}

fn push_existing_path(paths: &mut Vec<String>, path: impl AsRef<Path>) {
    let path = path.as_ref();
    if path.exists() {
        let candidate = path_to_swi_string(path);
        if !paths.iter().any(|p| p.eq_ignore_ascii_case(&candidate)) {
            paths.push(candidate);
        }
    }
}

fn env_file(var: &str) -> Option<String> {
    std::env::var(var).ok().and_then(|value| {
        let path = PathBuf::from(value);
        path.exists().then(|| path_to_swi_string(&path))
    })
}

fn strip_verbatim_prefix(path: PathBuf) -> PathBuf {
    let raw = path.to_string_lossy();
    if let Some(rest) = raw.strip_prefix(r"\\?\UNC\") {
        PathBuf::from(format!(r"\\{}", rest))
    } else if let Some(rest) = raw.strip_prefix(r"\\?\") {
        PathBuf::from(rest)
    } else {
        path
    }
}

fn canonical_for_swi(path: PathBuf) -> PathBuf {
    path.canonicalize()
        .map(strip_verbatim_prefix)
        .unwrap_or(path)
}

fn path_to_swi_string(path: &Path) -> String {
    strip_verbatim_prefix(path.to_path_buf())
        .to_string_lossy()
        .into_owned()
}

fn push_unique_path(paths: &mut Vec<PathBuf>, path: PathBuf) {
    let path = strip_verbatim_prefix(path);
    let candidate = path_to_swi_string(&path)
        .replace('/', "\\")
        .to_ascii_lowercase();
    if !paths.iter().any(|existing| {
        path_to_swi_string(existing)
            .replace('/', "\\")
            .to_ascii_lowercase()
            == candidate
    }) {
        paths.push(path);
    }
}

fn path_entries_with_file(name: &str) -> Vec<PathBuf> {
    std::env::var_os("PATH")
        .map(|path| {
            std::env::split_paths(&path)
                .map(|dir| dir.join(name))
                .filter(|candidate| candidate.exists())
                .collect()
        })
        .unwrap_or_default()
}

fn is_windows_system32(path: &Path) -> bool {
    let lower = path
        .to_string_lossy()
        .replace('/', "\\")
        .to_ascii_lowercase();
    lower.ends_with("\\windows\\system32\\onnxruntime.dll")
}

fn program_files_dirs() -> Vec<PathBuf> {
    ["ProgramFiles", "ProgramFiles(x86)"]
        .iter()
        .filter_map(|var| std::env::var_os(var).map(PathBuf::from))
        .collect()
}

fn program_files_swipl_candidates() -> Vec<PathBuf> {
    let mut candidates = Vec::new();
    for root in program_files_dirs() {
        push_swipl_dir_candidate(&mut candidates, root.join("swipl"));
        push_swipl_dir_candidate(&mut candidates, root.join("SWI-Prolog"));

        if let Ok(entries) = std::fs::read_dir(&root) {
            for entry in entries.flatten() {
                let name = entry.file_name().to_string_lossy().to_ascii_lowercase();
                if name.contains("swipl") || name.contains("swi-prolog") {
                    push_swipl_dir_candidate(&mut candidates, entry.path());
                }
            }
        }
    }
    candidates
}

fn push_swipl_dir_candidate(candidates: &mut Vec<PathBuf>, dir: PathBuf) {
    let dll = dir.join("bin").join("libswipl.dll");
    push_unique_path(candidates, dll);
}

fn swipl_dll_candidates() -> Vec<PathBuf> {
    let mut candidates = Vec::new();

    if let Some(path) = bundled_dll("libswipl.dll") {
        push_unique_path(&mut candidates, PathBuf::from(path));
    }

    for var in ["SPELL_SWIPL_DLL", "SWIPL_DLL"] {
        if let Ok(path) = std::env::var(var) {
            push_unique_path(&mut candidates, PathBuf::from(path));
        }
    }

    for var in ["SWI_HOME_DIR", "SWIPL_HOME_DIR", "SWIPL_HOME", "SWI_HOME"] {
        if let Ok(home) = std::env::var(var) {
            push_swipl_dir_candidate(&mut candidates, PathBuf::from(home));
        }
    }

    let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    push_swipl_dir_candidate(&mut candidates, manifest_dir.join("../../swipl"));
    push_swipl_dir_candidate(&mut candidates, manifest_dir.join("../../swi-prolog"));
    push_swipl_dir_candidate(&mut candidates, manifest_dir.join("../../vendor/swipl"));

    for path in [
        PathBuf::from("C:/Program Files/swipl/bin/libswipl.dll"),
        PathBuf::from("C:/Program Files/SWI-Prolog/bin/libswipl.dll"),
        PathBuf::from("C:/swipl/bin/libswipl.dll"),
        PathBuf::from("C:/tools/swipl/bin/libswipl.dll"),
        PathBuf::from("C:/ProgramData/chocolatey/lib/swi-prolog/tools/swipl/bin/libswipl.dll"),
        PathBuf::from("C:/ProgramData/chocolatey/lib/swipl/tools/swipl/bin/libswipl.dll"),
        PathBuf::from("C:/msys64/mingw64/bin/libswipl.dll"),
        PathBuf::from("C:/msys64/ucrt64/bin/libswipl.dll"),
    ] {
        push_unique_path(&mut candidates, path);
    }

    for path in path_entries_with_file("libswipl.dll") {
        push_unique_path(&mut candidates, path);
    }

    for exe in path_entries_with_file("swipl.exe") {
        if let Some(bin_dir) = exe.parent() {
            push_unique_path(&mut candidates, bin_dir.join("libswipl.dll"));
        }
    }

    for path in program_files_swipl_candidates() {
        push_unique_path(&mut candidates, path);
    }

    candidates
}

fn find_swipl_dll() -> String {
    let candidates = swipl_dll_candidates();
    for path in &candidates {
        if path.exists() {
            return path_to_swi_string(path);
        }
    }

    eprintln!("SWI-Prolog libswipl.dll not found for Windows dev run.");
    eprintln!("Install SWI-Prolog, or set SPELL_SWIPL_DLL to libswipl.dll.");
    eprintln!("Checked SWI-Prolog candidates:");
    for path in candidates.iter().take(24) {
        eprintln!("  {}", path.display());
    }

    "C:/Program Files/swipl/bin/libswipl.dll".to_string()
}

fn swipl_home_for_dll(swipl_path: &str) -> Option<PathBuf> {
    let dll_parent = Path::new(swipl_path).parent()?;

    let app_home = dll_parent.join("../Resources/swipl");
    if app_home.join("boot.prc").exists() {
        return Some(canonical_for_swi(app_home));
    }

    for candidate in [
        dll_parent.parent().map(|p| p.to_path_buf()),
        dll_parent.parent().map(|p| p.join("lib").join("swipl")),
        Some(dll_parent.join("../lib/swipl")),
    ]
    .into_iter()
    .flatten()
    {
        if candidate.join("boot.prc").exists() && candidate.join("library").is_dir() {
            return Some(canonical_for_swi(candidate));
        }
    }

    None
}

pub fn configure_swipl_home_env() {
    if let Ok(home) = std::env::var("SWI_HOME_DIR") {
        let normalized = path_to_swi_string(Path::new(&home));
        if normalized != home {
            unsafe {
                std::env::set_var("SWI_HOME_DIR", normalized);
            }
        }
        return;
    }

    let swipl_path = find_swipl_dll();
    let Some(home) = swipl_home_for_dll(&swipl_path) else {
        return;
    };

    unsafe {
        std::env::set_var("SWI_HOME_DIR", path_to_swi_string(&home));
    }
}

fn windows_ort_candidates() -> Vec<String> {
    let mut paths = Vec::new();
    if let Some(bundled) = bundled_dll("onnxruntime.dll") {
        paths.push(bundled);
    }
    if let Some(path) = env_file("SPELL_ORT_DYLIB") {
        paths.push(path);
    }

    let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    push_existing_path(
        &mut paths,
        manifest_dir.join(format!(
            "../../onnxruntime/onnxruntime-win-x64-{ONNX_RUNTIME_VERSION}/lib/onnxruntime.dll"
        )),
    );
    push_existing_path(
        &mut paths,
        manifest_dir.join("../../onnxruntime/onnxruntime-win-x64-1.23.0/lib/onnxruntime.dll"),
    );
    push_existing_path(
        &mut paths,
        format!("C:/onnxruntime/onnxruntime-win-x64-{ONNX_RUNTIME_VERSION}/lib/onnxruntime.dll"),
    );
    push_existing_path(
        &mut paths,
        "C:/onnxruntime/onnxruntime-win-x64-1.23.0/lib/onnxruntime.dll",
    );

    for path in path_entries_with_file("onnxruntime.dll") {
        if !is_windows_system32(&path) {
            push_existing_path(&mut paths, path);
        }
    }

    paths
}

pub struct WindowsPlatform {
    /// Cached selected text — polled via UIA while external app has focus
    cached_selected_text: Arc<Mutex<Option<String>>>,
    /// Memoized result of `foreground_app()` valid for ~150ms. The egui
    /// repaint loop calls foreground_app on every frame (5–10 Hz);
    /// without this cache each frame did GetForegroundWindow +
    /// OpenProcess + QueryFullProcessImageNameW + CloseHandle. Even
    /// with proper CloseHandle (added in dbe90b3) that's millions of
    /// kernel-handle alloc/free cycles per hour, plus full process-
    /// image-path queries — wasted work because the foreground app
    /// rarely changes within 150ms.
    cached_fg: Mutex<Option<(Instant, ForegroundApp)>>,
    last_space_down: AtomicBool,
}

impl WindowsPlatform {
    pub fn new() -> Self {
        let cached_selected_text = Arc::new(Mutex::new(None));
        let sel_clone = Arc::clone(&cached_selected_text);

        // Background thread polls selected text via UIA every 200ms
        // (same pattern as Mac's fg-poller).
        //
        // Idle gate: when the user hasn't touched the keyboard or mouse
        // for >10s, skip the UIA call entirely. UIA's GetFocusedElement
        // and TextPattern.GetSelection are cross-process round-trips
        // that accumulate provider-side bookkeeping in the target app;
        // leaving them firing 5×/sec for hours while the user is away
        // was the dominant cause of "Spell makes my whole PC slow after
        // an hour idle" reports. Resumes polling on the very next tick
        // after activity, so there's no perceptible latency for the
        // user — the worst case is a 200ms staleness on the selected
        // text the moment they return.
        std::thread::Builder::new()
            .name("sel-poller".into())
            .spawn(move || {
                // Initialize COM on this thread
                unsafe {
                    use windows::Win32::System::Com::*;
                    let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok();
                }
                loop {
                    if idle_millis() < 10_000 {
                        let sel = poll_uia_selected_text();
                        if let Ok(mut lock) = sel_clone.lock() {
                            if sel.is_some() {
                                *lock = sel;
                            }
                            // Keep last known selection when our app gets focus
                        }
                    }
                    std::thread::sleep(Duration::from_millis(200));
                }
            })
            .expect("Failed to spawn selection poller");

        Self {
            cached_selected_text,
            cached_fg: Mutex::new(None),
            last_space_down: AtomicBool::new(false),
        }
    }
}

/// Per-thread cached `IUIAutomation`. Recreating the in-proc COM server on
/// every poll (5 Hz on `sel-poller`, 10 Hz on the egui repaint thread that
/// drives `accessibility_win.rs`) was the dominant cause of "Spell makes
/// my whole PC sluggish after a while" reports. Each `CoCreateInstance` of
/// `CUIAutomation` forces UIA to spin up provider proxies in EVERY target
/// process Spell queries (Word, Chrome, Edge, ...). Those proxies don't
/// drop instantly and accumulate handles in *target* apps — invisible in
/// Spell.exe's Task Manager memory column but lethal to system
/// responsiveness. UIA is a free-threaded interface; one cached instance
/// per thread is the documented Microsoft pattern.
pub fn cached_uia() -> Option<windows::Win32::UI::Accessibility::IUIAutomation> {
    use std::cell::RefCell;
    use windows::Win32::System::Com::{CLSCTX_INPROC_SERVER, CoCreateInstance};
    use windows::Win32::UI::Accessibility::{CUIAutomation, IUIAutomation};
    thread_local! {
        static UIA: RefCell<Option<IUIAutomation>> = const { RefCell::new(None) };
    }
    UIA.with(|cell| {
        if cell.borrow().is_none() {
            let new = unsafe { CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER) }.ok();
            *cell.borrow_mut() = new;
        }
        cell.borrow().clone()
    })
}

/// Poll selected text from the focused UIA element using TextPattern.GetSelection.
/// Works for Word, Notepad, Chrome, and any UIA-compliant app.
/// Returns None if no text is selected or our app has focus.
fn poll_uia_selected_text() -> Option<String> {
    unsafe {
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

        let uia = cached_uia()?;
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

fn read_word_before_cursor_uia() -> Option<String> {
    unsafe {
        use windows::Win32::UI::Accessibility::*;

        let uia = cached_uia()?;
        let focused = uia.GetFocusedElement().ok()?;

        if let Ok(pattern2) =
            focused.GetCurrentPatternAs::<IUIAutomationTextPattern2>(UIA_TextPattern2Id)
        {
            let mut is_active = windows::core::BOOL::default();
            if let Ok(caret_range) = pattern2.GetCaretRange(&mut is_active) {
                let lookback = caret_range.Clone().ok()?;
                let _ = lookback.MoveEndpointByUnit(
                    TextPatternRangeEndpoint_Start,
                    TextUnit_Character,
                    -80,
                );
                let text = lookback.GetText(-1).ok()?.to_string();
                return last_word(&text);
            }
        }

        if let Ok(pattern) =
            focused.GetCurrentPatternAs::<IUIAutomationTextPattern>(UIA_TextPatternId)
        {
            if let Ok(selection) = pattern.GetSelection() {
                if selection.Length().ok()? > 0 {
                    let range: IUIAutomationTextRange = selection.GetElement(0).ok()?;
                    let lookback = range.Clone().ok()?;
                    let _ = lookback.MoveEndpointByUnit(
                        TextPatternRangeEndpoint_Start,
                        TextUnit_Character,
                        -80,
                    );
                    let text = lookback.GetText(-1).ok()?.to_string();
                    return last_word(&text);
                }
            }
        }

        // Do not use ValuePattern.CurrentValue() here. It exposes the whole
        // control value but no caret offset, so a mid-document Space press
        // would speak the final word of the field instead of the word behind
        // the cursor.
        None
    }
}

fn last_word(text: &str) -> Option<String> {
    text.trim_end()
        .rsplit(|c: char| c.is_whitespace())
        .next()
        .map(|word| {
            word.trim_matches(|c: char| {
                !(c.is_alphanumeric() || c == '-' || c == '\'')
            })
            .to_string()
        })
        .filter(|word| !word.is_empty())
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
        use windows::Win32::System::Threading::{
            OpenProcess, PROCESS_NAME_FORMAT, PROCESS_QUERY_LIMITED_INFORMATION,
            QueryFullProcessImageNameW,
        };
        use windows::Win32::UI::WindowsAndMessaging::{
            GetForegroundWindow, GetWindowTextW, GetWindowThreadProcessId,
        };

        // Hot path: serve from cache if the last query was <150ms ago.
        // egui repaints at 5–10 Hz so the cache hits ~95% of the time
        // without harming app-switch latency (one frame ≈ 100ms feels
        // instant). Saves an OpenProcess + QueryFullProcessImageName +
        // CloseHandle round-trip per cached frame; over the course of
        // an idle hour that's tens of thousands of kernel-handle
        // alloc/free cycles eliminated.
        if let Ok(cache) = self.cached_fg.lock() {
            if let Some((when, ref app)) = *cache {
                if when.elapsed() < Duration::from_millis(150) {
                    return app.clone();
                }
            }
        }

        let fg = unsafe { GetForegroundWindow() };
        let mut pid = 0u32;
        unsafe {
            GetWindowThreadProcessId(fg, Some(&mut pid));
        }

        let mut buf = [0u16; 128];
        let len = unsafe { GetWindowTextW(fg, &mut buf) };
        let title = String::from_utf16_lossy(&buf[..len as usize]);

        let exe = if pid > 0 {
            if let Ok(handle) =
                unsafe { OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) }
            {
                let mut exe_buf = [0u16; 260];
                let mut exe_len = exe_buf.len() as u32;
                let result = if unsafe {
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
                };
                // CRITICAL: the `windows` crate's HANDLE has NO Drop impl —
                // OpenProcess returns a kernel handle that must be closed
                // explicitly. Without this CloseHandle, every call to
                // foreground_app() (~10 Hz from egui's repaint loop) leaks
                // one kernel handle. After ~30 minutes of use the handle
                // table gets congested and the whole desktop slows down,
                // even though Spell.exe's memory footprint looks normal.
                // Add the "Handles" column under Task Manager > Details to
                // see this climb monotonically while Spell runs.
                let _ = unsafe { windows::Win32::Foundation::CloseHandle(handle) };
                result
            } else {
                String::new()
            }
        } else {
            String::new()
        };

        let result = ForegroundApp {
            handle: fg.0 as isize,
            pid,
            title,
            exe_name: exe,
        };
        // Memoize for ~150ms — see cache check at top of fn.
        if let Ok(mut cache) = self.cached_fg.lock() {
            *cache = Some((Instant::now(), result.clone()));
        }
        result
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

    fn is_writing_app(&self, app: &ForegroundApp) -> bool {
        if app.pid == std::process::id() {
            return false;
        }
        if matches!(self.classify_app(app), AppKind::Word | AppKind::Browser | AppKind::Notepad) {
            return true;
        }

        !matches!(
            app.exe_name.as_str(),
            "werfault.exe"
                | "snippingtool.exe"
                | "shellexperiencehost.exe"
                | "searchhost.exe"
                | "startmenuexperiencehost.exe"
                | "taskmgr.exe"
                | "explorer.exe"
        )
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

    fn take_space_press(&self) -> bool {
        use windows::Win32::UI::Input::KeyboardAndMouse::GetAsyncKeyState;
        let state = unsafe { GetAsyncKeyState(0x20) };
        let space_down = state < 0;
        let pressed_since_last_query = (state & 1) != 0;
        let was_down = self.last_space_down.swap(space_down, Ordering::Relaxed);
        pressed_since_last_query || (space_down && !was_down)
    }

    fn get_word_before_cursor(&self) -> Option<String> {
        read_word_before_cursor_uia()
    }

    fn copy_to_clipboard(&self, text: &str) {
        use windows::Win32::Foundation::HANDLE;
        use windows::Win32::System::DataExchange::*;
        use windows::Win32::System::Memory::*;
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
        windows_ort_candidates()
    }

    fn swipl_path(&self) -> &str {
        static PATH: std::sync::OnceLock<String> = std::sync::OnceLock::new();
        PATH.get_or_init(find_swipl_dll).as_str()
    }

    // Windows AX bridge already nudges the caret by +4 px in
    // bridge/accessibility_win.rs::get_caret_pos (`gui.rcCaret.bottom + 4`).
    // The default 189 px in the trait was tuned for Word's task pane and put
    // the popup ~190 px below the line — far enough that it dropped well
    // below the user's typing area on narrow editors. Mac sits at 30 (net
    // ~80 px below caret bottom after the +49 platform adjustment); match
    // that net distance here so the popup sticks close to the line you're
    // writing on without overlapping it.
    fn caret_offset_below(&self) -> f32 {
        30.0
    }

    fn read_selected_text(&self) -> Option<String> {
        if let Ok(lock) = self.cached_selected_text.lock() {
            lock.clone()
        } else {
            None
        }
    }

    fn init_tts(&self, lang: &dyn language::LanguageVoice) {
        crate::tts::init_tts(lang);
    }
}
