/// macOS Accessibility bridge — reads cursor context from any AX-exposing
/// text element. Covers Teams, Safari inputs, Chrome inputs, TextEdit, etc.
/// Does NOT cover Word (handled by the add-in HTTP bridge).
///
/// Two read paths:
///   1. `AXSelectedTextRange` + `AXStringForRange` — standard character-offset
///      text protocol. Works for Electron apps and most native/web text fields.
///   2. `AXValue` — full element text, assuming cursor at end. Last-resort
///      fallback for rich content editors that don't expose CFRange ranges.
use super::{CursorContext, RawCursorText, TextBridge, build_context};
use accessibility_sys::*;
use core_foundation::base::{CFRelease, CFTypeRef, TCFType};
use core_foundation::string::CFString;
use std::sync::atomic::{AtomicI64, Ordering};
use std::sync::Mutex;

/// Log each distinct trace message at most once per 3s to avoid spam.
fn trace_once(msg: &str) {
    use std::time::Instant;
    static LAST: std::sync::OnceLock<Mutex<(String, Instant)>> = std::sync::OnceLock::new();
    let slot = LAST.get_or_init(|| Mutex::new((String::new(), Instant::now() - std::time::Duration::from_secs(60))));
    let mut g = slot.lock().unwrap();
    if g.0 != msg || g.1.elapsed() > std::time::Duration::from_secs(3) {
        crate::log!("{}", msg);
        g.0 = msg.to_string();
        g.1 = Instant::now();
    }
}

unsafe fn role_of(elem: AXUIElementRef) -> String {
    let mut role: CFTypeRef = std::ptr::null();
    let err = AXUIElementCopyAttributeValue(
        elem,
        CFString::new("AXRole").as_concrete_TypeRef(),
        &mut role,
    );
    if err != 0 || role.is_null() { return "?".to_string(); }
    let s = CFString::wrap_under_create_rule(role as _).to_string();
    s
}

pub struct AxMacBridge {
    /// Target process PID stored via `set_fg_hwnd`. On macOS we reuse the
    /// hwnd plumbing to pass the foreground app's PID.
    target_pid: AtomicI64,
    /// Last word at cursor — cached so `replace_word` knows how many
    /// backspaces to send (Electron doesn't expose AXSetSelectedText).
    last_word: Mutex<String>,
}

impl AxMacBridge {
    pub fn new() -> Self {
        Self {
            target_pid: AtomicI64::new(0),
            last_word: Mutex::new(String::new()),
        }
    }
}

impl TextBridge for AxMacBridge {
    fn name(&self) -> &str { "Accessibility (macOS)" }

    fn is_available(&self) -> bool { true }

    fn set_fg_hwnd(&self, hwnd: isize) {
        self.target_pid.store(hwnd as i64, Ordering::Relaxed);
    }

    fn read_context(&self) -> Option<CursorContext> {
        let pid = self.target_pid.load(Ordering::Relaxed) as i32;
        if pid <= 0 { return None; }
        unsafe {
            let app = AXUIElementCreateApplication(pid);
            let mut focused: CFTypeRef = std::ptr::null();
            let err = AXUIElementCopyAttributeValue(
                app,
                CFString::new("AXFocusedUIElement").as_concrete_TypeRef(),
                &mut focused,
            );
            CFRelease(app as _);
            if err != 0 || focused.is_null() {
                trace_once(&format!("ax_mac pid={} no focused (err={})", pid, err));
                return None;
            }
            let elem = focused as AXUIElementRef;
            let role = role_of(elem);
            let via_range = try_read_via_text_range(elem);
            let via_value = if via_range.is_none() { try_read_via_value(elem) } else { None };
            let via = match (&via_range, &via_value) {
                (Some(_), _) => "range",
                (None, Some(_)) => "value",
                _ => "none",
            };
            CFRelease(focused);
            trace_once(&format!("ax_mac pid={} role={} via={}", pid, role, via));
            let ctx = via_range.or(via_value);
            if let Some(c) = &ctx {
                if let Ok(mut w) = self.last_word.lock() {
                    *w = c.word.clone();
                }
            }
            ctx
        }
    }

    fn replace_word(&self, new_text: &str) -> bool {
        // Completion flow uses `"{prefix}|{word}"` format: prefix = what's
        // already typed before the cursor, word = the completed word the
        // user picked. We just need to insert the missing SUFFIX after the
        // cursor — no deletes required.
        if let Some((prefix, word)) = new_text.split_once('|') {
            let suffix = if word.to_lowercase().starts_with(&prefix.to_lowercase()) {
                &word[prefix.len()..]
            } else {
                // Prefix mismatch — treat as full replacement of the prefix.
                return self.backspace_paste(prefix, word);
            };
            return self.paste_text(suffix);
        }
        // Plain replace — use cached word as the thing to delete.
        let cached_word = self.last_word.lock().ok().map(|w| w.clone()).unwrap_or_default();
        self.backspace_paste(&cached_word, new_text)
    }

    fn find_and_replace(&self, find: &str, replace: &str) -> bool {
        self.backspace_paste(find, replace)
    }

    fn find_and_replace_in_context(&self, find: &str, replace: &str, _context: &str) -> bool {
        self.backspace_paste(find, replace)
    }

    fn find_and_replace_in_context_at(&self, find: &str, replace: &str, _context: &str, _off: usize) -> bool {
        self.backspace_paste(find, replace)
    }

    fn find_and_replace_in_paragraph(&self, find: &str, replace: &str, _p: &str, _c: &str, _o: usize) -> bool {
        self.backspace_paste(find, replace)
    }
}

impl AxMacBridge {
    /// Replace the misspelled word at cursor: ⌫×len(find) + Cmd+V(replace).
    /// Electron / Teams refuses AXSetValue / AXSetSelectedText, so synthesized
    /// keystrokes + clipboard is the only portable path. User's clipboard is
    /// saved and restored asynchronously.
    fn backspace_paste(&self, find: &str, replace: &str) -> bool {
        let pid = self.target_pid.load(Ordering::Relaxed);
        let word_len = find.chars().count();
        crate::log!("ax_mac backspace_paste: pid={} find='{}' len={} replace='{}'",
            pid, find, word_len, replace);
        if pid <= 0 || word_len == 0 { return false; }

        let saved_clip = pbpaste();
        pbcopy(replace);

        let target = pid as u32;
        bring_app_to_front(target);
        send_backspaces_to(target, word_len);
        // Give Teams a beat to process the deletes before the paste event.
        std::thread::sleep(std::time::Duration::from_millis(30));
        send_cmd_v_to(target);

        std::thread::spawn(move || {
            std::thread::sleep(std::time::Duration::from_millis(250));
            pbcopy(&saved_clip);
        });
        true
    }

    /// Paste text at cursor without deleting anything (completion suffix).
    fn paste_text(&self, text: &str) -> bool {
        let pid = self.target_pid.load(Ordering::Relaxed);
        crate::log!("ax_mac paste_text: pid={} text='{}'", pid, text);
        if pid <= 0 || text.is_empty() { return false; }

        let saved_clip = pbpaste();
        pbcopy(text);

        let target = pid as u32;
        bring_app_to_front(target);
        send_cmd_v_to(target);

        std::thread::spawn(move || {
            std::thread::sleep(std::time::Duration::from_millis(250));
            pbcopy(&saved_clip);
        });
        true
    }
}

/// Path 1: CFRange-based read. Works for any element that implements the
/// standard AX text protocol (Electron, Cocoa NSText, most web inputs).
unsafe fn try_read_via_text_range(elem: AXUIElementRef) -> Option<CursorContext> {
    // Cursor position as the start of the (possibly zero-length) selection.
    let mut range_val: CFTypeRef = std::ptr::null();
    let err = AXUIElementCopyAttributeValue(
        elem,
        CFString::new("AXSelectedTextRange").as_concrete_TypeRef(),
        &mut range_val,
    );
    if err != 0 || range_val.is_null() { return None; }
    let mut sel = core_foundation::base::CFRange { location: 0, length: 0 };
    let ok = AXValueGetValue(
        range_val as AXValueRef,
        kAXValueTypeCFRange,
        &mut sel as *mut _ as _,
    );
    CFRelease(range_val);
    if !ok { return None; }

    // Total character count for clamping.
    let mut count_val: CFTypeRef = std::ptr::null();
    let err = AXUIElementCopyAttributeValue(
        elem,
        CFString::new("AXNumberOfCharacters").as_concrete_TypeRef(),
        &mut count_val,
    );
    let total: isize = if err == 0 && !count_val.is_null() {
        let n = core_foundation::number::CFNumber::wrap_under_create_rule(count_val as _);
        n.to_i64().unwrap_or(0) as isize
    } else { 0 };
    if total <= 0 { return None; }

    // 300 chars either side of cursor — enough for sentence context without
    // dragging huge web pages.
    const WIN: isize = 300;
    let cursor = sel.location;
    let win_start = (cursor - WIN).max(0);
    let win_end = (cursor + WIN).min(total);
    let before = read_string_for_range(
        elem,
        core_foundation::base::CFRange { location: win_start, length: (cursor - win_start).max(0) },
    ).unwrap_or_default();
    let after = read_string_for_range(
        elem,
        core_foundation::base::CFRange { location: cursor, length: (win_end - cursor).max(0) },
    ).unwrap_or_default();

    if before.is_empty() && after.is_empty() { return None; }

    // Tier 1: real caret bounds via AXBoundsForRange (Cocoa, Safari inputs).
    // Tier 2: focused-element frame (Electron apps like Teams/Slack/VSCode
    //         that Apple's private AXBoundsForRange APIs intentionally
    //         don't support — see electron/electron#34755).
    // Char-count × glyph-width estimation was tried but fails for line
    // wraps, variable-width fonts, and emoji — not worth the complexity.
    let caret_pos = read_caret_bounds(elem, sel)
        .or_else(|| unsafe { element_frame_bottom_left(elem) });
    Some(build_context(&RawCursorText { before, after }, caret_pos))
}

/// Get caret screen position via `AXBoundsForRange` on a zero-length range
/// at the current cursor. Returns bottom-of-caret coords, x < 50 filtered
/// out as garbage (matches caret_screen_position's filter).
unsafe fn read_caret_bounds(
    elem: AXUIElementRef,
    sel: core_foundation::base::CFRange,
) -> Option<(i32, i32)> {
    // Try AXBoundsForRange first with a zero-length range AT the cursor, then
    // with a length-1 range ENDING at the cursor (some apps refuse empty
    // ranges). Teams/Electron typically fails both with (0, screen_h, 0x0).
    let try_range = |r: core_foundation::base::CFRange| -> Option<(i32, i32)> {
        let ax_range = AXValueCreate(kAXValueTypeCFRange, &r as *const _ as _);
        if ax_range.is_null() { return None; }
        let mut bounds_val: CFTypeRef = std::ptr::null();
        let err = AXUIElementCopyParameterizedAttributeValue(
            elem,
            CFString::new("AXBoundsForRange").as_concrete_TypeRef(),
            ax_range as CFTypeRef,
            &mut bounds_val,
        );
        CFRelease(ax_range as _);
        if err != 0 || bounds_val.is_null() { return None; }
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
        trace_once(&format!("ax_mac bounds: loc={} len={} rect=({},{} {}x{})",
            r.location, r.length,
            rect.origin.x as i32, rect.origin.y as i32,
            rect.size.width as i32, rect.size.height as i32));
        let x = rect.origin.x as i32;
        let y = (rect.origin.y + rect.size.height) as i32;
        if x < 50 { return None; }
        // Zero-size + zero-origin is the macOS "no value" garbage.
        if rect.size.width as i32 == 0 && rect.size.height as i32 == 0
           && rect.origin.x as i32 == 0 { return None; }
        Some((x, y))
    };

    let zero = core_foundation::base::CFRange { location: sel.location, length: 0 };
    if let Some(p) = try_range(zero) { return Some(p); }
    if sel.location > 0 {
        let one_before = core_foundation::base::CFRange { location: sel.location - 1, length: 1 };
        if let Some(p) = try_range(one_before) {
            // Use right edge of the preceding character = caret position.
            // `try_range` returns left edge + full height. Adjust x to right.
            // We can't get the rect.size.width directly here (try_range
            // returned only x,y), so approximate cursor x = left + guess.
            // Leave as-is — the left edge is within a few px of caret.
            return Some(p);
        }
    }

    None
}

unsafe fn element_frame_bottom_left(elem: AXUIElementRef) -> Option<(i32, i32)> {
    let mut pos_val: CFTypeRef = std::ptr::null();
    let perr = AXUIElementCopyAttributeValue(
        elem,
        CFString::new("AXPosition").as_concrete_TypeRef(),
        &mut pos_val,
    );
    if perr != 0 || pos_val.is_null() { return None; }
    let mut size_val: CFTypeRef = std::ptr::null();
    let serr = AXUIElementCopyAttributeValue(
        elem,
        CFString::new("AXSize").as_concrete_TypeRef(),
        &mut size_val,
    );
    if serr != 0 || size_val.is_null() {
        CFRelease(pos_val);
        return None;
    }
    let mut pos = core_graphics::geometry::CGPoint::new(0.0, 0.0);
    let mut sz = core_graphics::geometry::CGSize::new(0.0, 0.0);
    let ok_p = AXValueGetValue(pos_val as AXValueRef, kAXValueTypeCGPoint, &mut pos as *mut _ as *mut _);
    let ok_s = AXValueGetValue(size_val as AXValueRef, kAXValueTypeCGSize, &mut sz as *mut _ as *mut _);
    CFRelease(pos_val);
    CFRelease(size_val);
    if !ok_p || !ok_s { return None; }
    let x = pos.x as i32;
    // Anchor ~40 px below the field's bottom edge so the window sits clearly
    // below the caret line. Main.rs still applies its own caret_offset and
    // may flip above when there's no screen room, but starting lower gives
    // a visible gap from the text field in both directions.
    let y = (pos.y + sz.height) as i32 + 40;
    if x < 50 { return None; }
    trace_once(&format!("ax_mac frame fallback: elem_x={} elem_y={} elem_h={} returning ({},{})",
        pos.x as i32, pos.y as i32, sz.height as i32, x, y));
    Some((x, y))
}

unsafe fn read_string_for_range(
    elem: AXUIElementRef,
    r: core_foundation::base::CFRange,
) -> Option<String> {
    if r.length <= 0 { return Some(String::new()); }
    let ax_range = AXValueCreate(kAXValueTypeCFRange, &r as *const _ as _);
    if ax_range.is_null() { return None; }
    let mut result: CFTypeRef = std::ptr::null();
    let err = AXUIElementCopyParameterizedAttributeValue(
        elem,
        CFString::new("AXStringForRange").as_concrete_TypeRef(),
        ax_range as CFTypeRef,
        &mut result,
    );
    CFRelease(ax_range as _);
    if err != 0 || result.is_null() { return None; }
    Some(CFString::wrap_under_create_rule(result as _).to_string())
}

// ── Keyboard / clipboard helpers ──

/// Backspace key code (kVK_Delete).
const KEY_DELETE: u16 = 0x33;
/// V key code (kVK_ANSI_V).
const KEY_V: u16 = 0x09;
/// Cmd modifier flag (kCGEventFlagMaskCommand).
const FLAG_CMD: u64 = 1 << 20;
/// Post to "session" event tap (kCGSessionEventTap).
const SESSION_TAP: u32 = 1;

#[link(name = "ApplicationServices", kind = "framework")]
unsafe extern "C" {
    fn CGEventCreateKeyboardEvent(source: *mut std::ffi::c_void, keyCode: u16, keyDown: bool)
        -> *mut std::ffi::c_void;
    fn CGEventSetFlags(event: *mut std::ffi::c_void, flags: u64);
    fn CGEventPost(tap: u32, event: *mut std::ffi::c_void);
    fn CGEventPostToPid(pid: u32, event: *mut std::ffi::c_void);
}

fn post_key_to_pid(pid: u32, keycode: u16, down: bool, flags: u64) {
    unsafe {
        let event = CGEventCreateKeyboardEvent(std::ptr::null_mut(), keycode, down);
        if event.is_null() { return; }
        if flags != 0 {
            CGEventSetFlags(event, flags);
        }
        if pid > 0 {
            CGEventPostToPid(pid, event);
        } else {
            CGEventPost(SESSION_TAP, event);
        }
        core_foundation::base::CFRelease(event as _);
    }
}

fn send_backspaces_to(pid: u32, n: usize) {
    for _ in 0..n {
        post_key_to_pid(pid, KEY_DELETE, true, 0);
        post_key_to_pid(pid, KEY_DELETE, false, 0);
        std::thread::sleep(std::time::Duration::from_millis(2));
    }
}

fn send_cmd_v_to(pid: u32) {
    post_key_to_pid(pid, KEY_V, true, FLAG_CMD);
    post_key_to_pid(pid, KEY_V, false, FLAG_CMD);
}

/// Bring an app to frontmost via `osascript` — required so our synthesized
/// keystrokes land on the target app rather than our egui window.
fn bring_app_to_front(pid: u32) {
    let script = format!(
        r#"tell application "System Events"
            set frontProcess to first application process whose unix id is {}
            set frontmost of frontProcess to true
        end tell"#,
        pid
    );
    let _ = std::process::Command::new("osascript")
        .arg("-e")
        .arg(&script)
        .output();
    // Tiny delay so the OS processes the focus change before keys are sent.
    std::thread::sleep(std::time::Duration::from_millis(60));
}

fn pbpaste() -> String {
    std::process::Command::new("pbpaste")
        .output()
        .ok()
        .and_then(|o| String::from_utf8(o.stdout).ok())
        .unwrap_or_default()
}

fn pbcopy(text: &str) {
    use std::io::Write;
    if let Ok(mut child) = std::process::Command::new("pbcopy")
        .stdin(std::process::Stdio::piped())
        .spawn()
    {
        if let Some(stdin) = child.stdin.as_mut() {
            let _ = stdin.write_all(text.as_bytes());
        }
        let _ = child.wait();
    }
}

/// Path 2: full-value read. Assumes the cursor sits at the end of the text
/// (correct for most chat/compose inputs). Used when CFRange APIs fail.
unsafe fn try_read_via_value(elem: AXUIElementRef) -> Option<CursorContext> {
    let mut value: CFTypeRef = std::ptr::null();
    let err = AXUIElementCopyAttributeValue(
        elem,
        CFString::new("AXValue").as_concrete_TypeRef(),
        &mut value,
    );
    if err != 0 || value.is_null() { return None; }
    // AXValue may be a CFString for text elements, or other types for
    // sliders/buttons. Only handle string case.
    let type_id = core_foundation::base::CFGetTypeID(value);
    if type_id != core_foundation::string::CFString::type_id() {
        CFRelease(value);
        return None;
    }
    let s = CFString::wrap_under_create_rule(value as _).to_string();
    if s.trim().is_empty() { return None; }
    Some(build_context(
        &RawCursorText { before: s, after: String::new() },
        None,
    ))
}
