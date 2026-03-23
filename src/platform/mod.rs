/// Platform abstraction layer.
///
/// Each OS implements `PlatformServices` in its own module.
/// The only `#[cfg]` lives in `create_platform()` — all other code
/// uses the trait, so adding a new platform is just another `impl`.

#[cfg(target_os = "windows")]
pub mod windows;

#[cfg(target_os = "macos")]
pub mod macos;

// ── Types shared by all platforms ──

/// Information about the currently focused application.
#[derive(Debug, Clone, Default)]
pub struct ForegroundApp {
    /// OS-level window/app handle (HWND on Windows, PID on Mac)
    pub handle: isize,
    /// Process ID
    pub pid: u32,
    /// Window title
    pub title: String,
    /// Executable / app name (lowercase, e.g. "winword.exe" or "microsoft word")
    pub exe_name: String,
}

/// Coarse classification of the foreground app.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AppKind {
    OurApp,
    Word,
    Browser,
    Notepad,
    Other,
}

/// Platform-specific services consumed by `BridgeManager` and `ContextApp`.
///
/// Every method that touches the OS goes through this trait so that the
/// rest of the application is platform-agnostic.
///
/// Implementations must be non-blocking on the UI thread. If the underlying
/// OS call is slow (e.g. AppleScript on macOS), the implementation must use
/// a background thread internally and return cached results.
pub trait PlatformServices: Send + Sync {
    /// One-time runtime init (e.g. COM on Windows). Called once at startup.
    fn init_runtime(&self);

    /// Query the currently focused window / application.
    /// Must return quickly — use caching if the OS call is slow.
    fn foreground_app(&self) -> ForegroundApp;

    /// Classify a foreground app into a known category.
    fn classify_app(&self, app: &ForegroundApp) -> AppKind;

    /// PID of our own process.
    fn our_pid(&self) -> u32 {
        std::process::id()
    }

    /// Primary screen dimensions in logical pixels.
    fn screen_size(&self) -> (f32, f32);

    /// Bring a window/app to the foreground by its handle.
    fn set_foreground(&self, handle: isize);

    /// Poll the global hotkey state.
    /// Returns `(ctrl_held, space_held)`.
    fn check_hotkey_state(&self) -> (bool, bool);

    /// Copy text to the system clipboard.
    fn copy_to_clipboard(&self, text: &str);

    /// Path to an emoji font, or None if system default should be used.
    fn emoji_font_path(&self) -> Option<&str>;

    /// Candidate paths for ONNX Runtime dynamic library.
    fn ort_dylib_candidates(&self) -> Vec<String>;

    /// Path to the SWI-Prolog shared library.
    fn swipl_path(&self) -> &str;

    /// Get the screen position of the text cursor (caret) in the focused app.
    /// Returns (x, y) in screen coordinates, or None if unavailable.
    fn caret_screen_position(&self) -> Option<(i32, i32)> { None }

    /// Initialize TTS engine (platform-specific).
    fn init_tts(&self);

    /// Read the currently selected text in the frontmost application.
    /// Returns None if no text is selected or if accessibility access is denied.
    fn read_selected_text(&self) -> Option<String> { None }
}

/// Construct the correct `PlatformServices` for the current OS.
pub fn create_platform() -> Box<dyn PlatformServices> {
    #[cfg(target_os = "windows")]
    { Box::new(windows::WindowsPlatform::new()) }

    #[cfg(target_os = "macos")]
    { Box::new(macos::MacPlatform::new()) }

    #[cfg(not(any(target_os = "windows", target_os = "macos")))]
    { Box::new(StubPlatform) }
}

/// No-op platform for unsupported targets — allows compilation everywhere.
#[cfg(not(any(target_os = "windows", target_os = "macos")))]
struct StubPlatform;

#[cfg(not(any(target_os = "windows", target_os = "macos")))]
impl PlatformServices for StubPlatform {
    fn init_runtime(&self) {}
    fn foreground_app(&self) -> ForegroundApp { ForegroundApp::default() }
    fn classify_app(&self, _app: &ForegroundApp) -> AppKind { AppKind::Other }
    fn screen_size(&self) -> (f32, f32) { (1920.0, 1080.0) }
    fn set_foreground(&self, _handle: isize) {}
    fn check_hotkey_state(&self) -> (bool, bool) { (false, false) }
    fn copy_to_clipboard(&self, _text: &str) {}
    fn emoji_font_path(&self) -> Option<&str> { None }
    fn ort_dylib_candidates(&self) -> Vec<String> { vec![] }
    fn swipl_path(&self) -> &str { "libswipl.so" }
    fn init_tts(&self) {}
}
