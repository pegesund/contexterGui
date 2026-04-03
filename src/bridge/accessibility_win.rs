use super::{CursorContext, RawCursorText, TextBridge, build_context, extract_word_before_cursor, extract_word_after_cursor};
use std::io::Write;
use windows::Win32::Foundation::*;
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::System::Com::*;
use windows::Win32::UI::Accessibility::*;
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::Win32::UI::Input::KeyboardAndMouse::*;

fn bridge_log(msg: &str) {
    let path = std::env::temp_dir().join("acatts-bridge.log");
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
    /// Last known good UIA text element (e.g. the Edge textarea).
    /// Re-read from this when GetFocusedElement() returns something else.
    saved_element: std::cell::RefCell<Option<IUIAutomationElement>>,
    /// PID of the app that owns the saved element
    saved_element_pid: std::cell::Cell<u32>,
}

impl AccessibilityBridge {
    pub fn new() -> Self {
        AccessibilityBridge {
            target_hwnd: std::cell::Cell::new(0),
            fg_hwnd: std::cell::Cell::new(0),
            edit_hwnd: std::cell::Cell::new(0),
            cached_doc: std::cell::RefCell::new(String::new()),
            saved_element: std::cell::RefCell::new(None),
            saved_element_pid: std::cell::Cell::new(0),
        }
    }


    /// Try to read text from a UIA element using TextPattern2, TextPattern v1, or ValuePattern.
    fn try_read_raw(element: &IUIAutomationElement) -> Option<(RawCursorText, String)> {
        unsafe {
            // 1. TextPattern2 — best: gives caret position + before/after text (Notepad, Word)
            if let Ok(pattern2) =
                element.GetCurrentPatternAs::<IUIAutomationTextPattern2>(UIA_TextPattern2Id)
            {
                let mut is_active = windows::core::BOOL::default();
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
                    let text = doc_range.GetText(-1).unwrap_or_default().to_string();
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
                    let text = value.to_string();
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

    /// Get raw text from the user's text field. Strategy:
    /// 1. If fg_hwnd PID changed (user switched apps), clear saved element
    /// 2. Try GetFocusedElement() — if it's a text field, save it and read
    /// 3. If focused element is wrong (terminal, etc.), re-read from saved element
    fn get_raw_text(&self) -> Option<RawCursorText> {
        unsafe {
            let uia: IUIAutomation =
                CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER).ok()?;
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
                    if let Some((raw, doc)) = Self::try_read_raw(&focused) {
                        if !doc.is_empty() && Self::is_text_field(&doc) {
                            bridge_log(&format!("Focused text field: '{}' ({} chars)",
                                {let mut e=60.min(doc.len()); while e>0 && !doc.is_char_boundary(e){e-=1;} &doc[..e]}, doc.len()));
                            *self.cached_doc.borrow_mut() = doc;
                            *self.saved_element.borrow_mut() = Some(focused);
                            self.saved_element_pid.set(focused_pid);
                            return Some(raw);
                        }
                    }
                }
            }

            // Step 2: Focused element was wrong — re-read from saved element
            // This gives us LIVE text even when the terminal has focus
            let saved = self.saved_element.borrow().clone();
            if let Some(ref element) = saved {
                if let Some((raw, doc)) = Self::try_read_raw(element) {
                    if !doc.is_empty() {
                        bridge_log(&format!("Saved element re-read: '{}' ({} chars)",
                            {let mut e=60.min(doc.len()); while e>0 && !doc.is_char_boundary(e){e-=1;} &doc[..e]}, doc.len()));
                        *self.cached_doc.borrow_mut() = doc;
                        return Some(raw);
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

    fn read_context_via_uia(&self) -> Option<String> {
        unsafe {
            let uia: IUIAutomation =
                CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER).ok()?;
            let focused = uia.GetFocusedElement().ok()?;

            if let Ok(pattern2) =
                focused.GetCurrentPatternAs::<IUIAutomationTextPattern2>(UIA_TextPattern2Id)
            {
                let mut is_active = windows::core::BOOL::default();
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

    /// Replace word at cursor — try UIA TextPattern2, fall back to keyboard
    fn replace_word_impl(&self, replace_text: &str) -> bool {
        unsafe {
            let uia: IUIAutomation = match CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER) {
                Ok(u) => u,
                Err(_) => {
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
                let mut is_active = windows::core::BOOL::default();
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

    /// Get a TextPattern2 from any reachable element — tries focused, then HWND fallback
    fn get_text_pattern(&self) -> Option<IUIAutomationTextPattern2> {
        unsafe {
            let uia: IUIAutomation =
                CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER).ok()?;

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

    /// Find text in document and replace via UIA
    fn find_replace_via_uia(&self, find: &str, replace: &str, context: &str) -> bool {
        bridge_log(&format!("=== find_replace_via_uia ==="));
        bridge_log(&format!("FIND: '{}'", find));
        bridge_log(&format!("REPLACE: '{}'", replace));
        unsafe {
            let pattern = match self.get_text_pattern() {
                Some(p) => p,
                None => { bridge_log("FAILED: no TextPattern"); return false; }
            };

            let doc_range = match pattern.DocumentRange() {
                Ok(r) => r,
                Err(e) => { bridge_log(&format!("FAILED: DocumentRange: {:?}", e)); return false; }
            };

            // Log full document text BEFORE change
            let doc_before = doc_range.GetText(-1).unwrap_or_default().to_string();
            bridge_log(&format!("DOC BEFORE ({} chars):\n{}", doc_before.len(), doc_before));

            // Re-get doc range for FindText (GetText may have consumed it)
            let doc_range = pattern.DocumentRange().unwrap();
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

                    // Select the text via UIA (needed so EM_REPLACESEL knows what to replace)
                    let doc_range2 = pattern.DocumentRange().unwrap();
                    let found_range2 = doc_range2.FindText(&windows::core::BSTR::from(find), false, false);
                    match found_range2 {
                        Ok(fr) => {
                            let sel_result = fr.Select();
                            bridge_log(&format!("Select result: {:?}", sel_result));
                        }
                        Err(e) => {
                            bridge_log(&format!("Re-FindText failed: {:?}", e));
                            return false;
                        }
                    }

                    std::thread::sleep(std::time::Duration::from_millis(50));
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
        let caret_pos = self.get_caret_pos();
        let raw = self.get_raw_text();
        match raw {
            Some(raw) => Some(build_context(&raw, caret_pos)),
            None => Some(CursorContext {
                caret_pos,
                ..Default::default()
            }),
        }
    }

    fn replace_word(&self, new_text: &str) -> bool {
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

    fn set_target_hwnd(&self, hwnd: isize) {
        if hwnd != self.target_hwnd.get() {
            self.edit_hwnd.set(0); // Reset cached edit control when app changes
        }
        self.target_hwnd.set(hwnd);
    }

    fn set_fg_hwnd(&self, hwnd: isize) {
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
    unsafe {
        use windows::Win32::UI::WindowsAndMessaging::{FindWindowExW, GetClassNameW};
        let none = HWND(std::ptr::null_mut());

        fn search_recursive(parent: HWND, depth: u32) -> Option<HWND> {
            if depth > 10 { return None; }
            unsafe {
                use windows::Win32::UI::WindowsAndMessaging::{FindWindowExW, GetClassNameW};
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
        bridge_log("No edit child found, using parent window");
        parent
    }
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
