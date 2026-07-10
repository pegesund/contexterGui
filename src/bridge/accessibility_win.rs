use super::{CursorContext, RawCursorText, TextBridge, build_context, extract_word_before_cursor, extract_word_after_cursor};
use std::io::Write;
use std::time::{Duration, Instant};
use windows::core::BOOL;
use windows::Win32::Foundation::*;
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::System::Com::*;
use windows::Win32::UI::Accessibility::*;
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::Win32::UI::Input::KeyboardAndMouse::*;

fn bridge_log(msg: &str) {
    let path = std::env::temp_dir().join("spell-bridge.log");
    if let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true).append(true).open(&path)
    {
        let _ = writeln!(f, "{}", msg);
        let _ = f.flush();
    }
}

pub struct AccessibilityBridge {
    /// Saved HWND of the target app (set externally when good context is read)
    pub target_hwnd: std::cell::Cell<isize>,
    /// Foreground HWND set by BridgeManager before each read_context call
    pub fg_hwnd: std::cell::Cell<isize>,
    /// Cached HWND of the actual edit control inside the target app
    edit_hwnd: std::cell::Cell<isize>,
    /// Cached full document text from last successful read
    cached_doc: std::cell::RefCell<String>,
    /// Cursor offset inside cached_doc. For windowed reads this is local to cached_doc.
    cached_cursor: std::cell::Cell<usize>,
    /// Last known good UIA text element (e.g. the Edge textarea).
    /// Re-read from this when GetFocusedElement() returns something else.
    saved_element: std::cell::RefCell<Option<IUIAutomationElement>>,
    /// PID of the app that owns the saved element
    saved_element_pid: std::cell::Cell<u32>,
    /// UIA can report pre-replace text for a few frames after EM_REPLACESEL.
    replace_old_word: std::cell::RefCell<String>,
    replace_freeze_until: std::cell::Cell<Option<Instant>>,
}

impl AccessibilityBridge {
    pub fn new() -> Self {
        AccessibilityBridge {
            target_hwnd: std::cell::Cell::new(0),
            fg_hwnd: std::cell::Cell::new(0),
            edit_hwnd: std::cell::Cell::new(0),
            cached_doc: std::cell::RefCell::new(String::new()),
            cached_cursor: std::cell::Cell::new(0),
            saved_element: std::cell::RefCell::new(None),
            saved_element_pid: std::cell::Cell::new(0),
            replace_old_word: std::cell::RefCell::new(String::new()),
            replace_freeze_until: std::cell::Cell::new(None),
        }
    }

    fn doc_contains_word(doc: &str, word: &str) -> bool {
        doc.to_lowercase()
            .split(|c: char| !c.is_alphanumeric())
            .any(|w| w == word)
    }

    fn looks_like_slack_window_dump(doc: &str) -> bool {
        if doc.len() < 1000 {
            return false;
        }

        let doc_lower = doc.to_lowercase();
        let markers = [
            "show workspace switcher",
            "back in history",
            "forward in history",
            "chat with slackbot",
            "drafts & sent",
            "jump to date",
            "toggle file",
            "message ready to be sent",
            "processing uploaded file",
        ];
        let marker_hits = markers
            .iter()
            .filter(|marker| doc_lower.contains(**marker))
            .count();

        marker_hits >= 3
    }

    fn should_reject_stale_doc(&self, doc: &str) -> bool {
        let old_word = self.replace_old_word.borrow().clone();
        if old_word.is_empty() {
            return false;
        }
        let still_in_freeze = self.replace_freeze_until.get()
            .map(|until| Instant::now() < until)
            .unwrap_or(false);
        if still_in_freeze && Self::doc_contains_word(doc, &old_word) {
            bridge_log(&format!("Rejecting stale UIA read after replace: still contains '{}'", old_word));
            return true;
        }
        self.replace_old_word.borrow_mut().clear();
        self.replace_freeze_until.set(None);
        false
    }

    fn activate_replace_freeze(&self, old_word: &str) {
        *self.replace_old_word.borrow_mut() = old_word.to_lowercase();
        self.replace_freeze_until.set(Some(Instant::now() + Duration::from_millis(1200)));
    }

    fn accept_text_element(&self, element: IUIAutomationElement) -> Option<(RawCursorText, String, IUIAutomationElement, Option<(i32, i32)>)> {
        if !Self::is_visible_text_element(&element) {
            return None;
        }
        let (raw, doc) = Self::try_read_raw(&element)?;
        if doc.is_empty() || !Self::is_text_field(&doc) || self.should_reject_stale_doc(&doc) {
            return None;
        }
        if Self::looks_like_slack_window_dump(&doc) {
            bridge_log("Rejecting Slack window dump from UIA; waiting for focused composer text");
            return None;
        }
        let caret = Self::estimate_caret_from_element(&element, &raw.before);
        Some((raw, doc, element, caret))
    }

    fn is_visible_text_element(element: &IUIAutomationElement) -> bool {
        unsafe {
            if let Ok(is_offscreen) = element.CurrentIsOffscreen() {
                if is_offscreen.as_bool() {
                    bridge_log("Rejecting offscreen UIA text element");
                    return false;
                }
            }

            if let Ok(rect) = element.CurrentBoundingRectangle() {
                let width = rect.right - rect.left;
                let height = rect.bottom - rect.top;
                if width <= 0 || height <= 0 {
                    bridge_log(&format!(
                        "Rejecting zero-size UIA text element: {}x{}",
                        width, height
                    ));
                    return false;
                }
            }
        }
        true
    }

    fn find_text_element_from_hwnd(&self, uia: &IUIAutomation, hwnd: HWND) -> Option<(RawCursorText, String, IUIAutomationElement, Option<(i32, i32)>)> {
        unsafe {
            let root = uia.ElementFromHandle(hwnd).ok()?;
            if let Some(found) = self.accept_text_element(root.clone()) {
                return Some(found);
            }

            let condition = uia.CreateTrueCondition().ok()?;
            let descendants = root.FindAll(TreeScope_Descendants, &condition).ok()?;
            let count = descendants.Length().unwrap_or(0);
            for i in 0..count.min(100) {
                if let Ok(desc) = descendants.GetElement(i) {
                    if let Some(found) = self.accept_text_element(desc) {
                        return Some(found);
                    }
                }
            }
            None
        }
    }

    /// Try to read text from a UIA element using TextPattern2, TextPattern v1, or ValuePattern.
    fn try_read_raw(element: &IUIAutomationElement) -> Option<(RawCursorText, String)> {
        unsafe {
            // 1. TextPattern2 — best: gives caret position + before/after text (Notepad, Word)
            if let Ok(pattern2) =
                element.GetCurrentPatternAs::<IUIAutomationTextPattern2>(UIA_TextPattern2Id)
            {
                let mut is_active = BOOL::default();
                if let Ok(caret_range) = pattern2.GetCaretRange(&mut is_active) {
                    let before = (|| -> Option<String> {
                        let r = caret_range.Clone().ok()?;
                        let _ = r.MoveEndpointByUnit(
                            TextPatternRangeEndpoint_Start, TextUnit_Character, -2000);
                        Some(r.GetText(-1).ok()?.to_string())
                    })().unwrap_or_default();

                    let after = (|| -> Option<String> {
                        let r = caret_range.Clone().ok()?;
                        let _ = r.MoveEndpointByUnit(
                            TextPatternRangeEndpoint_End, TextUnit_Character, 2000);
                        Some(r.GetText(-1).ok()?.to_string())
                    })().unwrap_or_default();

                    let doc = format!("{}{}", before, after);
                    if !doc.is_empty() {
                        return Some((RawCursorText { before, after }, doc));
                    }
                }
            }

            // 2. TextPattern v1 — Edge textareas: has DocumentRange but no caret position
            if let Ok(tp1) =
                element.GetCurrentPatternAs::<IUIAutomationTextPattern>(UIA_TextPatternId)
            {
                if let Ok(doc_range) = tp1.DocumentRange() {
                    let text = doc_range.GetText(6000).unwrap_or_default().to_string();
                    if !text.is_empty() {
                        // No caret info from v1 — put everything as "before" (cursor at end)
                        return Some((RawCursorText { before: text.clone(), after: String::new() }, text));
                    }
                }
            }

            // 3. ValuePattern — fallback for elements with text value but no TextPattern
            if let Ok(vp) =
                element.GetCurrentPatternAs::<IUIAutomationValuePattern>(UIA_ValuePatternId)
            {
                if let Ok(value) = vp.CurrentValue() {
                    let value_text = value.to_string();
                    let total = value_text.chars().count();
                    let text = if total > 6000 {
                        let start = total.saturating_sub(6000);
                        let start_byte = Self::char_to_byte_offset(&value_text, start);
                        value_text[start_byte..].to_string()
                    } else {
                        value_text
                    };
                    if !text.is_empty() {
                        return Some((RawCursorText { before: text.clone(), after: String::new() }, text));
                    }
                }
            }

            None
        }
    }

    /// Is this element a reasonable text field (not a terminal buffer)?
    fn is_text_field(doc: &str) -> bool {
        // Terminal buffers are huge (>100K). Text fields are usually <50K.
        doc.len() < 100_000
    }

    fn estimate_caret_from_element(element: &IUIAutomationElement, before: &str) -> Option<(i32, i32)> {
        unsafe {
            let rect = element.CurrentBoundingRectangle().ok()?;
            let width = (rect.right - rect.left).max(1) as f32;
            let height = (rect.bottom - rect.top).max(1) as f32;
            if rect.left < 20 || width < 40.0 || height < 12.0 {
                return None;
            }

            let avg_char_w = 7.5f32;
            let line_h = 20.0f32;
            let pad_x = 12.0f32;
            let pad_y = 8.0f32;
            let usable_w = (width - pad_x * 2.0).max(80.0);
            let chars_per_line = (usable_w / avg_char_w).floor().max(1.0) as usize;
            let current_line = before.rsplit('\n').next().unwrap_or(before);
            let col_chars = current_line.chars().count();
            let visual_line = col_chars / chars_per_line;
            let col = col_chars % chars_per_line;
            let x = rect.left as f32 + pad_x + col as f32 * avg_char_w;
            let y = rect.top as f32 + pad_y + (visual_line as f32 + 1.0) * line_h;
            Some((x.round() as i32, y.min(rect.bottom as f32 + line_h).round() as i32))
        }
    }

    fn char_to_byte_offset(text: &str, char_offset: usize) -> usize {
        text.char_indices()
            .nth(char_offset)
            .map(|(byte_idx, _)| byte_idx)
            .unwrap_or(text.len())
    }

    fn paragraph_window_from_cached(doc: &str, cursor: usize) -> Option<(String, String, usize)> {
        if doc.trim().is_empty() {
            return None;
        }
        let cursor = cursor.min(doc.chars().count());
        let cursor_byte = Self::char_to_byte_offset(doc, cursor);
        let para_start_byte = doc[..cursor_byte]
            .rfind('\n')
            .map(|pos| pos + 1)
            .unwrap_or(0);
        let para_end_byte = doc[cursor_byte..]
            .find('\n')
            .map(|pos| cursor_byte + pos)
            .unwrap_or(doc.len());
        let para_start_char = doc[..para_start_byte].chars().count();
        let para_end_char = doc[..para_end_byte].chars().count();
        let max_chars = 6000usize;

        let (start_char, end_char) = if para_end_char.saturating_sub(para_start_char) <= max_chars {
            (para_start_char, para_end_char)
        } else {
            let half = max_chars / 2;
            let mut start = cursor.saturating_sub(half).max(para_start_char);
            let mut end = start.saturating_add(max_chars).min(para_end_char);
            if end.saturating_sub(start) < max_chars {
                start = end.saturating_sub(max_chars).max(para_start_char);
            }
            end = end.max(start);
            (start, end)
        };

        let start_byte = Self::char_to_byte_offset(doc, start_char);
        let end_byte = Self::char_to_byte_offset(doc, end_char);
        let text = doc[start_byte..end_byte].to_string();
        if text.trim().is_empty() {
            return None;
        }
        Some((format!("uia:{}", start_char), text, start_char))
    }

    /// Get raw text from the user's text field. Strategy:
    /// 1. If fg_hwnd PID changed (user switched apps), clear saved element
    /// 2. Try GetFocusedElement() — if it's a text field, save it and read
    /// 3. If focused element is wrong (terminal, etc.), re-read from saved element
    fn get_raw_text(&self) -> Option<(RawCursorText, Option<(i32, i32)>)> {
        unsafe {
            let uia = crate::platform::windows::cached_uia()?;
            let our_pid = std::process::id();

            // Step 0: If foreground app changed, clear saved element so we discover
            // the new app's text field via GetFocusedElement
            let fg_hwnd_val = self.fg_hwnd.get();
            if fg_hwnd_val != 0 {
                let mut fg_pid = 0u32;
                GetWindowThreadProcessId(HWND(fg_hwnd_val as *mut _), Some(&mut fg_pid));
                let saved_pid = self.saved_element_pid.get();
                if saved_pid != 0 && fg_pid != 0 && fg_pid != saved_pid {
                    bridge_log(&format!("App changed: pid {} → {} — clearing saved element", saved_pid, fg_pid));
                    *self.saved_element.borrow_mut() = None;
                    self.saved_element_pid.set(0);
                }
            }

            // Step 1: Try GetFocusedElement — accept only if it's a text field (<100K)
            if let Ok(focused) = uia.GetFocusedElement() {
                let focused_pid = focused.CurrentProcessId().unwrap_or(0) as u32;
                if focused_pid != our_pid && focused_pid != 0 {
                    if let Some((raw, doc, element, caret)) = self.accept_text_element(focused) {
                        bridge_log(&format!("Focused text field: '{}' ({} chars)",
                            {let mut e=60.min(doc.len()); while e>0 && !doc.is_char_boundary(e){e-=1;} &doc[..e]}, doc.len()));
                        self.cached_cursor.set(raw.before.chars().count());
                        *self.cached_doc.borrow_mut() = doc;
                        *self.saved_element.borrow_mut() = Some(element);
                        self.saved_element_pid.set(focused_pid);
                        return Some((raw, caret));
                    }
                }
            }

            // Step 2: Focused element was wrong — discover the edit element
            // under the foreground HWND. Win11 Notepad can initially expose
            // only a container until another UI event nudges UIA.
            if fg_hwnd_val != 0 {
                let mut fg_pid = 0u32;
                GetWindowThreadProcessId(HWND(fg_hwnd_val as *mut _), Some(&mut fg_pid));
                if fg_pid != our_pid && fg_pid != 0 {
                    if let Some((raw, doc, element, caret)) =
                        self.find_text_element_from_hwnd(&uia, HWND(fg_hwnd_val as *mut _))
                    {
                        bridge_log(&format!("HWND text field: '{}' ({} chars)",
                            {let mut e=60.min(doc.len()); while e>0 && !doc.is_char_boundary(e){e-=1;} &doc[..e]}, doc.len()));
                        self.cached_cursor.set(raw.before.chars().count());
                        *self.cached_doc.borrow_mut() = doc;
                        *self.saved_element.borrow_mut() = Some(element);
                        self.saved_element_pid.set(fg_pid);
                        return Some((raw, caret));
                    }
                }
            }

            // Step 3: Focused element was wrong — re-read from saved element
            // This gives us LIVE text even when the terminal has focus
            let saved = self.saved_element.borrow().clone();
            if let Some(ref element) = saved {
                if !Self::is_visible_text_element(element) {
                    bridge_log("Saved element hidden/offscreen - clearing");
                } else if let Some((raw, doc)) = Self::try_read_raw(element) {
                    if !doc.is_empty() && !self.should_reject_stale_doc(&doc) {
                        bridge_log(&format!("Saved element re-read: '{}' ({} chars)",
                            {let mut e=60.min(doc.len()); while e>0 && !doc.is_char_boundary(e){e-=1;} &doc[..e]}, doc.len()));
                        let caret = Self::estimate_caret_from_element(element, &raw.before);
                        self.cached_cursor.set(raw.before.chars().count());
                        *self.cached_doc.borrow_mut() = doc;
                        return Some((raw, caret));
                    }
                }
                // Saved element no longer readable — clear it
                bridge_log("Saved element stale — clearing");
                drop(saved);
                *self.saved_element.borrow_mut() = None;
                self.saved_element_pid.set(0);
            }

            None
        }
    }

    fn get_caret_pos(&self) -> Option<(i32, i32)> {
        unsafe {
            let fg = GetForegroundWindow();
            let thread_id = GetWindowThreadProcessId(fg, None);
            let mut gui = GUITHREADINFO {
                cbSize: std::mem::size_of::<GUITHREADINFO>() as u32,
                ..Default::default()
            };
            if GetGUIThreadInfo(thread_id, &mut gui).is_ok() && gui.hwndCaret != HWND::default() {
                let mut pt = POINT {
                    x: gui.rcCaret.left,
                    y: gui.rcCaret.bottom + 4,
                };
                let _ = ClientToScreen(gui.hwndCaret, &mut pt);
                return Some((pt.x, pt.y));
            }
            None
        }
    }

    fn is_plausible_gui_caret((_, y): (i32, i32)) -> bool {
        y.abs() >= 20
    }

    fn read_context_via_uia(&self) -> Option<String> {
        unsafe {
            let uia = crate::platform::windows::cached_uia()?;
            let focused = uia.GetFocusedElement().ok()?;

            if let Ok(pattern2) =
                focused.GetCurrentPatternAs::<IUIAutomationTextPattern2>(UIA_TextPattern2Id)
            {
                let mut is_active = BOOL::default();
                if let Ok(caret_range) = pattern2.GetCaretRange(&mut is_active) {
                    let context_range = caret_range.Clone().ok()?;
                    let _ = context_range.MoveEndpointByUnit(
                        TextPatternRangeEndpoint_Start,
                        TextUnit_Character,
                        -5000,
                    );
                    let text = context_range.GetText(-1).ok()?.to_string();
                    if !text.is_empty() {
                        return Some(text);
                    }
                }
            }
            None
        }
    }

    /// Restore the target app's HWND to the foreground before doing any UIA
    /// or keyboard-injection work. Without this, when the user clicks a
    /// suggestion in Spell's window, focus shifts TO Spell, and either
    /// `GetFocusedElement()` returns Spell's own UI (so TextPattern2 ops
    /// run on Spell instead of the user's text field) or the keyboard
    /// fallback types into Spell. End-user symptom in Notepad/Sticky Notes:
    /// click a suggestion → nothing happens in the target app. Word COM
    /// is unaffected because COM doesn't depend on focus.
    ///
    /// SetForegroundWindow only succeeds if the calling thread has the
    /// input focus; we just got it from the user's click, so it works.
    /// The 80ms sleep gives Windows time to actually transfer focus
    /// before subsequent UIA reads — empirically necessary on Win11.
    fn restore_target_foreground(&self) {
        let target = self.target_hwnd.get();
        if target == 0 { return; }
        unsafe {
            let hwnd = HWND(target as *mut _);
            let _ = SetForegroundWindow(hwnd);
        }
        std::thread::sleep(std::time::Duration::from_millis(80));
    }

    /// Replace word at cursor — try UIA TextPattern2, fall back to keyboard
    fn replace_word_impl(&self, replace_text: &str) -> bool {
        // The user clicked a suggestion in our window, so we just stole
        // focus from the target app. Hand it back before reading the
        // focused element or sending keystrokes.
        self.restore_target_foreground();
        unsafe {
            let uia = match crate::platform::windows::cached_uia() {
                Some(u) => u,
                None => {
                    select_word_keyboard();
                    send_string(replace_text);
                    return true;
                }
            };
            let focused = match uia.GetFocusedElement() {
                Ok(f) => f,
                Err(_) => {
                    select_word_keyboard();
                    send_string(replace_text);
                    return true;
                }
            };

            if let Ok(pattern2) =
                focused.GetCurrentPatternAs::<IUIAutomationTextPattern2>(UIA_TextPattern2Id)
            {
                let mut is_active = BOOL::default();
                if let Ok(caret_range) = pattern2.GetCaretRange(&mut is_active) {
                    // Scan backwards to find word start
                    let back_range = caret_range.Clone().unwrap();
                    let _ = back_range.MoveEndpointByUnit(
                        TextPatternRangeEndpoint_Start,
                        TextUnit_Character,
                        -50,
                    );
                    let before_text = back_range.GetText(-1).unwrap_or_default().to_string();
                    let word_before = extract_word_before_cursor(&before_text);
                    let chars_before = word_before.chars().count() as i32;

                    // Scan forwards to find word end
                    let fwd_range = caret_range.Clone().unwrap();
                    let _ = fwd_range.MoveEndpointByUnit(
                        TextPatternRangeEndpoint_End,
                        TextUnit_Character,
                        50,
                    );
                    let after_text = fwd_range.GetText(-1).unwrap_or_default().to_string();
                    let word_after = extract_word_after_cursor(&after_text);
                    let chars_after = word_after.chars().count() as i32;

                    // Build a range covering exactly the word
                    let word_range = caret_range.Clone().unwrap();
                    if chars_before > 0 {
                        let _ = word_range.MoveEndpointByUnit(
                            TextPatternRangeEndpoint_Start,
                            TextUnit_Character,
                            -chars_before,
                        );
                    }
                    if chars_after > 0 {
                        let _ = word_range.MoveEndpointByUnit(
                            TextPatternRangeEndpoint_End,
                            TextUnit_Character,
                            chars_after,
                        );
                    }

                    let _ = word_range.Select();
                    std::thread::sleep(std::time::Duration::from_millis(30));
                    send_string(replace_text);
                    self.activate_replace_freeze(&format!("{}{}", word_before, word_after));
                    return true;
                }
            }

            // Fallback: keyboard
            select_word_keyboard();
            send_string(replace_text);
            true
        }
    }

    /// Try to get a TextPattern2 from an element (for find/replace)
    fn try_get_text_pattern2(element: &IUIAutomationElement) -> Option<IUIAutomationTextPattern2> {
        unsafe {
            element.GetCurrentPatternAs::<IUIAutomationTextPattern2>(UIA_TextPattern2Id).ok()
        }
    }

    fn try_get_document_range(element: &IUIAutomationElement) -> Option<IUIAutomationTextRange> {
        unsafe {
            if let Ok(pattern2) =
                element.GetCurrentPatternAs::<IUIAutomationTextPattern2>(UIA_TextPattern2Id)
            {
                if let Ok(range) = pattern2.DocumentRange() {
                    return Some(range);
                }
            }

            if let Ok(pattern1) =
                element.GetCurrentPatternAs::<IUIAutomationTextPattern>(UIA_TextPatternId)
            {
                if let Ok(range) = pattern1.DocumentRange() {
                    return Some(range);
                }
            }

            None
        }
    }

    /// Get a TextPattern2 from any reachable element — tries focused, then HWND fallback
    fn get_text_pattern(&self) -> Option<IUIAutomationTextPattern2> {
        unsafe {
            let uia = crate::platform::windows::cached_uia()?;

            // Try focused element first
            if let Ok(focused) = uia.GetFocusedElement() {
                let name = focused.CurrentName().unwrap_or_default().to_string();
                eprintln!("  get_text_pattern: focused element name='{}'", name);
                if let Some(pat) = Self::try_get_text_pattern2(&focused) {
                    eprintln!("  get_text_pattern: got pattern from focused");
                    return Some(pat);
                }
            }

            // Fallback: saved target HWND
            let hwnd_val = self.target_hwnd.get();
            eprintln!("  get_text_pattern: target_hwnd={}", hwnd_val);
            if hwnd_val != 0 {
                let hwnd = HWND(hwnd_val as *mut _);
                if let Ok(element) = uia.ElementFromHandle(hwnd) {
                    let name = element.CurrentName().unwrap_or_default().to_string();
                    eprintln!("  get_text_pattern: hwnd element name='{}'", name);
                    if let Some(pat) = Self::try_get_text_pattern2(&element) {
                        eprintln!("  get_text_pattern: got pattern from hwnd element");
                        return Some(pat);
                    }
                    // Search descendants for element supporting TextPattern2
                    // Use TreeScope_Descendants (not just Children) — Notepad's
                    // text area is deeper in the tree
                    if let Ok(condition) = uia.CreateTrueCondition() {
                        if let Ok(descendants) = element.FindAll(TreeScope_Descendants, &condition) {
                            let count = descendants.Length().unwrap_or(0);
                            eprintln!("  get_text_pattern: {} descendants", count);
                            for i in 0..count.min(50) {
                                if let Ok(desc) = descendants.GetElement(i) {
                                    if let Some(pat) = Self::try_get_text_pattern2(&desc) {
                                        let dname = desc.CurrentName().unwrap_or_default().to_string();
                                        eprintln!("  got pattern from descendant {}: name='{}'", i, dname);
                                        return Some(pat);
                                    }
                                }
                            }
                        }
                    }
                } else {
                    eprintln!("  get_text_pattern: ElementFromHandle failed");
                }
            }

            None
        }
    }

    fn get_document_range(&self) -> Option<IUIAutomationTextRange> {
        unsafe {
            let saved = self.saved_element.borrow().clone();
            if let Some(ref element) = saved {
                if Self::is_visible_text_element(element) {
                    if let Some(range) = Self::try_get_document_range(element) {
                        bridge_log("get_document_range: saved element");
                        return Some(range);
                    }
                } else {
                    bridge_log("get_document_range: saved element hidden/offscreen - clearing");
                    drop(saved);
                    *self.saved_element.borrow_mut() = None;
                    self.saved_element_pid.set(0);
                }
            }

            let uia = crate::platform::windows::cached_uia()?;

            if let Ok(focused) = uia.GetFocusedElement() {
                if Self::is_visible_text_element(&focused) {
                    if let Some(range) = Self::try_get_document_range(&focused) {
                        bridge_log("get_document_range: focused element");
                        return Some(range);
                    }
                }
            }

            let hwnd_val = self.target_hwnd.get();
            if hwnd_val == 0 {
                return None;
            }

            let root = uia.ElementFromHandle(HWND(hwnd_val as *mut _)).ok()?;
            if Self::is_visible_text_element(&root) {
                if let Some(range) = Self::try_get_document_range(&root) {
                    bridge_log("get_document_range: hwnd root");
                    return Some(range);
                }
            }

            let condition = uia.CreateTrueCondition().ok()?;
            let descendants = root.FindAll(TreeScope_Descendants, &condition).ok()?;
            let count = descendants.Length().unwrap_or(0);
            bridge_log(&format!("get_document_range: {} descendants", count));
            for i in 0..count.min(100) {
                if let Ok(desc) = descendants.GetElement(i) {
                    if !Self::is_visible_text_element(&desc) {
                        continue;
                    }
                    if let Some(range) = Self::try_get_document_range(&desc) {
                        let name = desc.CurrentName().unwrap_or_default().to_string();
                        bridge_log(&format!("get_document_range: descendant {} name='{}'", i, name));
                        return Some(range);
                    }
                }
            }

            None
        }
    }

    /// Find text in document and replace via UIA
    fn find_replace_via_uia(&self, find: &str, replace: &str, context: &str) -> bool {
        bridge_log(&format!("=== find_replace_via_uia ==="));
        bridge_log(&format!("FIND: '{}'", find));
        bridge_log(&format!("REPLACE: '{}'", replace));
        // Same rationale as replace_word_impl: hand focus back to the
        // target app before searching/replacing — get_text_pattern() falls
        // back to GetFocusedElement, which would return our window after
        // the user clicked a grammar fix suggestion.
        self.restore_target_foreground();
        unsafe {
            let doc_range = match self.get_document_range() {
                Some(r) => r,
                None => { bridge_log("FAILED: no TextPattern document range"); return false; }
            };

            // Log full document text BEFORE change
            let doc_before = doc_range.GetText(-1).unwrap_or_default().to_string();
            bridge_log(&format!("DOC BEFORE ({} chars):\n{}", doc_before.len(), doc_before));

            let find_bstr = windows::core::BSTR::from(find);
            match doc_range.FindText(&find_bstr, false, false) {
                Ok(found_range) => {
                    let selected_text = found_range.GetText(-1).unwrap_or_default().to_string();
                    bridge_log(&format!("FindText matched: '{}' ({} chars)", selected_text, selected_text.len()));

                    if !context.is_empty() {
                        let check_range = found_range.Clone().unwrap();
                        let _ = check_range.MoveEndpointByUnit(
                            TextPatternRangeEndpoint_Start,
                            TextUnit_Character,
                            -50,
                        );
                        let _ = check_range.MoveEndpointByUnit(
                            TextPatternRangeEndpoint_End,
                            TextUnit_Character,
                            50,
                        );
                        let surrounding = check_range.GetText(-1).unwrap_or_default().to_string();
                        let ctx_words: Vec<&str> = context.split_whitespace().take(3).collect();
                        let matches = ctx_words.iter().any(|w| surrounding.contains(w));
                        bridge_log(&format!("Context check: matches={}", matches));
                        if !matches {
                            bridge_log("FAILED: context mismatch");
                            return false;
                        }
                    }

                    // Get or cache the edit control HWND
                    let hwnd_val = self.target_hwnd.get();
                    let edit_val = self.edit_hwnd.get();
                    let edit = if edit_val != 0 {
                        HWND(edit_val as *mut _)
                    } else if hwnd_val != 0 {
                        let found = find_edit_child(HWND(hwnd_val as *mut _));
                        self.edit_hwnd.set(found.0 as isize);
                        found
                    } else {
                        HWND(std::ptr::null_mut())
                    };
                    bridge_log(&format!("target_hwnd={} edit_hwnd={}", hwnd_val, edit.0 as isize));

                    // Select the text via UIA. Native Edit/RichEdit controls can
                    // then use EM_REPLACESEL. Electron/Slack exposes TextPattern
                    // but has no edit child, so SendInput types over the selected
                    // range in the foreground composer.
                    let sel_result = found_range.Select();
                    bridge_log(&format!("Select result: {:?}", sel_result));
                    if sel_result.is_err() {
                        return false;
                    }

                    std::thread::sleep(std::time::Duration::from_millis(35));
                    bridge_log(&format!("Sending replacement: '{}' ({} chars)", replace, replace.len()));

                    if edit.0 as isize != 0 {
                        // EM_REPLACESEL: atomic, no focus needed, targets edit control directly
                        replace_selection(edit, replace);
                    } else {
                        // Fallback: SendInput (needs focus)
                        if hwnd_val != 0 {
                            let _ = SetForegroundWindow(HWND(hwnd_val as *mut _));
                            std::thread::sleep(std::time::Duration::from_millis(100));
                        }
                        send_string(replace);
                    }

                    // Update cached doc so grammar checker sees the new text
                    {
                        let mut cached = self.cached_doc.borrow_mut();
                        *cached = cached.replacen(find, replace, 1);
                        bridge_log(&format!("DOC AFTER ({} chars):\n{}", cached.len(), &*cached));
                    }
                    self.activate_replace_freeze(find);

                    return true;
                }
                Err(e) => {
                    bridge_log(&format!("FAILED: FindText: {:?}", e));
                }
            }

            false
        }
    }
}

impl TextBridge for AccessibilityBridge {
    fn name(&self) -> &str {
        "Accessibility"
    }

    fn is_available(&self) -> bool {
        true
    }

    fn read_context(&self) -> Option<CursorContext> {
        let gui_caret = self.get_caret_pos().filter(|&pos| Self::is_plausible_gui_caret(pos));
        let raw = self.get_raw_text();
        match raw {
            Some((raw, estimated_caret)) => {
                let caret_pos = gui_caret.or(estimated_caret);
                let mut ctx = build_context(&raw, caret_pos);
                let cursor = raw.before.chars().count();
                ctx.cursor_doc_offset = Some(cursor);
                let cached = self.cached_doc.borrow();
                if let Some((para_id, _, _)) = Self::paragraph_window_from_cached(cached.as_str(), cursor) {
                    ctx.paragraph_id = para_id;
                }
                Some(ctx)
            }
            None => Some(CursorContext {
                caret_pos: gui_caret,
                ..Default::default()
            }),
        }
    }

    fn replace_word(&self, new_text: &str) -> bool {
        if let Some((prefix, word)) = new_text.split_once('|') {
            if word.to_lowercase().starts_with(&prefix.to_lowercase()) {
                let suffix: String = word.chars().skip(prefix.chars().count()).collect();
                if suffix.is_empty() {
                    return true;
                }
                self.restore_target_foreground();
                send_string(&suffix);
                self.activate_replace_freeze(prefix);
                return true;
            }
            return self.replace_word_impl(word);
        }
        self.replace_word_impl(new_text)
    }

    fn find_and_replace(&self, find: &str, replace: &str) -> bool {
        self.find_replace_via_uia(find, replace, "")
    }

    fn find_and_replace_in_context(&self, find: &str, replace: &str, context: &str) -> bool {
        self.find_replace_via_uia(find, replace, context)
    }

    fn read_document_context(&self) -> Option<String> {
        self.read_context_via_uia()
    }

    fn read_full_document(&self) -> Option<String> {
        // Use cached text from get_raw_text() — live UIA reads fail because
        // our always-on-top window steals focus between context read and grammar check
        let cached = self.cached_doc.borrow();
        if !cached.is_empty() {
            Some(cached.clone())
        } else {
            None
        }
    }

    fn read_paragraph_at(&self, _cursor_offset: usize) -> Option<(String, String, usize)> {
        let cached = self.cached_doc.borrow();
        Self::paragraph_window_from_cached(cached.as_str(), self.cached_cursor.get())
    }

    fn set_target_hwnd(&self, hwnd: isize) {
        if hwnd != self.target_hwnd.get() {
            self.edit_hwnd.set(0); // Reset cached edit control when app changes
        }
        self.target_hwnd.set(hwnd);
    }

    fn set_fg_hwnd(&self, hwnd: isize) {
        if hwnd != self.fg_hwnd.get() {
            bridge_log(&format!(
                "Foreground HWND changed: {} → {} — clearing cached UIA text",
                self.fg_hwnd.get(),
                hwnd
            ));
            self.edit_hwnd.set(0);
            self.cached_doc.borrow_mut().clear();
            self.cached_cursor.set(0);
            *self.saved_element.borrow_mut() = None;
            self.saved_element_pid.set(0);
        }
        self.fg_hwnd.set(hwnd);
    }
}

/// Select the word at cursor using keyboard shortcuts (Ctrl+Shift+Left)
fn select_word_keyboard() {
    unsafe {
        let inputs_end = [
            INPUT {
                r#type: INPUT_KEYBOARD,
                Anonymous: INPUT_0 {
                    ki: KEYBDINPUT {
                        wVk: VK_CONTROL,
                        wScan: 0,
                        dwFlags: KEYEVENTF_EXTENDEDKEY,
                        time: 0,
                        dwExtraInfo: 0,
                    },
                },
            },
            INPUT {
                r#type: INPUT_KEYBOARD,
                Anonymous: INPUT_0 {
                    ki: KEYBDINPUT {
                        wVk: VK_RIGHT,
                        wScan: 0,
                        dwFlags: KEYEVENTF_EXTENDEDKEY,
                        time: 0,
                        dwExtraInfo: 0,
                    },
                },
            },
            INPUT {
                r#type: INPUT_KEYBOARD,
                Anonymous: INPUT_0 {
                    ki: KEYBDINPUT {
                        wVk: VK_RIGHT,
                        wScan: 0,
                        dwFlags: KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP,
                        time: 0,
                        dwExtraInfo: 0,
                    },
                },
            },
            INPUT {
                r#type: INPUT_KEYBOARD,
                Anonymous: INPUT_0 {
                    ki: KEYBDINPUT {
                        wVk: VK_CONTROL,
                        wScan: 0,
                        dwFlags: KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP,
                        time: 0,
                        dwExtraInfo: 0,
                    },
                },
            },
        ];
        SendInput(&inputs_end, std::mem::size_of::<INPUT>() as i32);
        std::thread::sleep(std::time::Duration::from_millis(20));

        let inputs_sel = [
            INPUT {
                r#type: INPUT_KEYBOARD,
                Anonymous: INPUT_0 {
                    ki: KEYBDINPUT {
                        wVk: VK_CONTROL,
                        wScan: 0,
                        dwFlags: KEYEVENTF_EXTENDEDKEY,
                        time: 0,
                        dwExtraInfo: 0,
                    },
                },
            },
            INPUT {
                r#type: INPUT_KEYBOARD,
                Anonymous: INPUT_0 {
                    ki: KEYBDINPUT {
                        wVk: VK_SHIFT,
                        wScan: 0,
                        dwFlags: KEYEVENTF_EXTENDEDKEY,
                        time: 0,
                        dwExtraInfo: 0,
                    },
                },
            },
            INPUT {
                r#type: INPUT_KEYBOARD,
                Anonymous: INPUT_0 {
                    ki: KEYBDINPUT {
                        wVk: VK_LEFT,
                        wScan: 0,
                        dwFlags: KEYEVENTF_EXTENDEDKEY,
                        time: 0,
                        dwExtraInfo: 0,
                    },
                },
            },
            INPUT {
                r#type: INPUT_KEYBOARD,
                Anonymous: INPUT_0 {
                    ki: KEYBDINPUT {
                        wVk: VK_LEFT,
                        wScan: 0,
                        dwFlags: KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP,
                        time: 0,
                        dwExtraInfo: 0,
                    },
                },
            },
            INPUT {
                r#type: INPUT_KEYBOARD,
                Anonymous: INPUT_0 {
                    ki: KEYBDINPUT {
                        wVk: VK_SHIFT,
                        wScan: 0,
                        dwFlags: KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP,
                        time: 0,
                        dwExtraInfo: 0,
                    },
                },
            },
            INPUT {
                r#type: INPUT_KEYBOARD,
                Anonymous: INPUT_0 {
                    ki: KEYBDINPUT {
                        wVk: VK_CONTROL,
                        wScan: 0,
                        dwFlags: KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP,
                        time: 0,
                        dwExtraInfo: 0,
                    },
                },
            },
        ];
        SendInput(&inputs_sel, std::mem::size_of::<INPUT>() as i32);
        std::thread::sleep(std::time::Duration::from_millis(20));
    }
}

/// Recursively find the edit control inside a window (e.g., Notepad's RichEditD2DPT).
/// Windows 11 Notepad nests the edit control several levels deep.
fn find_edit_child(parent: HWND) -> HWND {
    let none = HWND(std::ptr::null_mut());

    fn search_recursive(parent: HWND, depth: u32) -> Option<HWND> {
        if depth > 10 { return None; }
        unsafe {
            use windows::Win32::UI::WindowsAndMessaging::FindWindowExW;
            let none = HWND(std::ptr::null_mut());
            let edit_classes = ["RichEditD2DPT", "Edit", "RichEdit20W", "RICHEDIT50W"];

            // Check direct children for edit classes
            for class in &edit_classes {
                let class_wide: Vec<u16> = class.encode_utf16().chain(std::iter::once(0)).collect();
                if let Ok(child) = FindWindowExW(Some(parent), Some(none),
                    windows::core::PCWSTR(class_wide.as_ptr()), windows::core::PCWSTR(std::ptr::null()))
                {
                    if child.0 as isize != 0 {
                        bridge_log(&format!("Found edit child: class='{}' hwnd={:?} depth={}", class, child, depth));
                        return Some(child);
                    }
                }
            }

            // Recurse into all children
            let mut prev = none;
            loop {
                match FindWindowExW(Some(parent), Some(prev),
                    windows::core::PCWSTR(std::ptr::null()), windows::core::PCWSTR(std::ptr::null()))
                {
                    Ok(child) if child.0 as isize != 0 => {
                        if let Some(found) = search_recursive(child, depth + 1) {
                            return Some(found);
                        }
                        prev = child;
                    }
                    _ => break,
                }
            }
            None
        }
    }

    if let Some(edit) = search_recursive(parent, 0) {
        return edit;
    }
    bridge_log("No edit child found");
    none
}

/// Replace the current selection in an edit control using EM_REPLACESEL.
/// This is a single atomic message — no focus issues, no character-by-character.
fn replace_selection(hwnd: HWND, text: &str) {
    unsafe {
        use windows::Win32::UI::WindowsAndMessaging::SendMessageW;
        const EM_REPLACESEL: u32 = 0x00C2;
        let wide: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
        SendMessageW(
            hwnd,
            EM_REPLACESEL,
            Some(windows::Win32::Foundation::WPARAM(1)), // fCanUndo = TRUE
            Some(windows::Win32::Foundation::LPARAM(wide.as_ptr() as isize)),
        );
        bridge_log(&format!("replace_selection: {} chars sent via EM_REPLACESEL", text.chars().count()));
    }
}

/// Type a string by sending WM_CHAR messages directly to a window handle.
/// Unlike SendInput (which goes to the foreground window), PostMessage targets
/// a specific window and works even if our GUI steals focus.
fn send_string_to_hwnd(text: &str, hwnd: HWND) {
    unsafe {
        use windows::Win32::UI::WindowsAndMessaging::{PostMessageW, WM_CHAR};
        std::thread::sleep(std::time::Duration::from_millis(50));

        for ch in text.encode_utf16() {
            let _ = PostMessageW(Some(hwnd), WM_CHAR, windows::Win32::Foundation::WPARAM(ch as usize), windows::Win32::Foundation::LPARAM(0));
        }
        bridge_log(&format!("send_string_to_hwnd: {} chars posted to hwnd {:?}", text.chars().count(), hwnd));
    }
}

/// Fallback: Type a string via SendInput (goes to foreground window)
fn send_string(text: &str) {
    unsafe {
        std::thread::sleep(std::time::Duration::from_millis(30));

        for ch in text.encode_utf16() {
            let inputs = [
                INPUT {
                    r#type: INPUT_KEYBOARD,
                    Anonymous: INPUT_0 {
                        ki: KEYBDINPUT {
                            wVk: VIRTUAL_KEY(0),
                            wScan: ch,
                            dwFlags: KEYEVENTF_UNICODE,
                            time: 0,
                            dwExtraInfo: 0,
                        },
                    },
                },
                INPUT {
                    r#type: INPUT_KEYBOARD,
                    Anonymous: INPUT_0 {
                        ki: KEYBDINPUT {
                            wVk: VIRTUAL_KEY(0),
                            wScan: ch,
                            dwFlags: KEYEVENTF_UNICODE | KEYEVENTF_KEYUP,
                            time: 0,
                            dwExtraInfo: 0,
                        },
                    },
                },
            ];
            SendInput(&inputs, std::mem::size_of::<INPUT>() as i32);
        }
        bridge_log(&format!("send_string: {} chars via SendInput", text.chars().count()));
    }
}
