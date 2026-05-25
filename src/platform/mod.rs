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

/// Chosen caret position plus its source, for logging/debugging.
#[derive(Debug, Clone, Copy)]
pub struct CaretPositionDecision {
    pub position: (i32, i32),
    pub source: &'static str,
}

/// Platform policy for feeding text into the grammar checker.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct GrammarFeedPolicy {
    /// Use a bridge paragraph read instead of full-document scanning.
    pub use_paragraph_feed: bool,
    /// Suppress the full-document fallback because another event source owns
    /// grammar updates for this bridge.
    pub suppress_full_doc_scan: bool,
    /// Some bridges can synthesize a paragraph from offset 0 when the bridge
    /// has not reported a cursor offset yet.
    pub force_cursor_offset: bool,
}

impl GrammarFeedPolicy {
    pub fn full_document() -> Self {
        Self {
            use_paragraph_feed: false,
            suppress_full_doc_scan: false,
            force_cursor_offset: false,
        }
    }

    pub fn paragraph(force_cursor_offset: bool) -> Self {
        Self {
            use_paragraph_feed: true,
            suppress_full_doc_scan: true,
            force_cursor_offset,
        }
    }

    pub fn external() -> Self {
        Self {
            use_paragraph_feed: false,
            suppress_full_doc_scan: true,
            force_cursor_offset: false,
        }
    }
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

    /// True iff this app is one Spell can plausibly help with (writing prose).
    /// False for code editors, terminals, system utilities — apps where Spell's
    /// popup would only get in the user's way. Default true: when in doubt,
    /// stay active.
    fn is_writing_app(&self, _app: &ForegroundApp) -> bool { true }

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
    fn caret_screen_position(&self) -> Option<(i32, i32)> { None }

    /// Whether this foreground kind should trigger explicit caret polling.
    fn should_poll_caret_position(&self, kind: AppKind) -> bool {
        matches!(kind, AppKind::Word | AppKind::Browser | AppKind::Notepad)
    }

    /// Convert bridge-reported caret coordinates into this platform's
    /// last-caret coordinate space.
    fn normalize_bridge_caret_position(
        &self,
        caret: Option<(i32, i32)>,
        _pixels_per_point: f32,
    ) -> Option<(i32, i32)> {
        caret
    }

    /// Convert platform AX/UIA caret coordinates into this platform's
    /// last-caret coordinate space.
    fn normalize_platform_caret_position(&self, caret: Option<(i32, i32)>) -> Option<(i32, i32)> {
        caret.map(|(x, y)| (x, y + 49))
    }

    /// Choose how the active bridge should feed text into grammar checking.
    fn grammar_feed_policy(
        &self,
        active_bridge_name: &str,
        _app: &ForegroundApp,
        _kind: AppKind,
    ) -> GrammarFeedPolicy {
        if active_bridge_name == "Word COM" {
            GrammarFeedPolicy::paragraph(false)
        } else {
            GrammarFeedPolicy::full_document()
        }
    }

    /// Choose between platform and bridge caret sources for the current app.
    fn choose_caret_position(
        &self,
        kind: AppKind,
        platform_caret: Option<(i32, i32)>,
        bridge_caret: Option<(i32, i32)>,
        pixels_per_point: f32,
    ) -> Option<CaretPositionDecision> {
        let platform_caret = self.normalize_platform_caret_position(platform_caret);
        let bridge_caret = self
            .normalize_bridge_caret_position(bridge_caret, pixels_per_point)
            .filter(|(x, y)| *x != 0 || *y != 0);

        let prefer_bridge = kind == AppKind::Browser;
        let (position, source) = if prefer_bridge {
            match (bridge_caret, platform_caret) {
                (Some(pos), _) => (pos, "bridge"),
                (None, Some(pos)) => (pos, "platform"),
                _ => return None,
            }
        } else {
            match (platform_caret, bridge_caret) {
                (Some(pos), _) => (pos, "platform"),
                (None, Some(pos)) => (pos, "bridge"),
                _ => return None,
            }
        };

        Some(CaretPositionDecision { position, source })
    }

    /// Convert a stored caret position into egui logical points.
    fn caret_position_to_logical(&self, caret: (i32, i32), pixels_per_point: f32) -> (f32, f32) {
        if self.caret_is_physical_pixels() {
            let dpi_scale = pixels_per_point.max(1.0);
            (caret.0 as f32 / dpi_scale, caret.1 as f32 / dpi_scale)
        } else {
            (caret.0 as f32, caret.1 as f32)
        }
    }

    fn set_tab_intercept(&self, _active: bool) {}
    fn take_tab_press(&self) -> bool { false }
    fn take_space_press(&self) -> bool { false }
    fn get_word_before_cursor(&self) -> Option<String> { None }

    /// Vertical offset in logical pixels between the caret and the app window.
    fn caret_offset_below(&self) -> f32 { 189.0 }

    /// Horizontal offset in logical pixels from the caret to the app window.
    fn caret_offset_right(&self) -> f32 { 0.0 }

    /// Whether caret_screen_position returns physical pixels (true on Windows)
    /// or logical points (false on macOS). Controls whether DPI scaling is applied.
    fn caret_is_physical_pixels(&self) -> bool { true }

    /// Initialize TTS engine (platform-specific).
    fn init_tts(&self, lang: &dyn language::LanguageVoice);

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
    fn init_tts(&self, _lang: &dyn language::LanguageVoice) {}
}
