use super::{CursorContext, RawCursorText, TextBridge, build_context, extract_word_before_cursor, extract_word_after_cursor};
use windows::Win32::Foundation::*;
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::System::Com::*;
use windows::Win32::UI::Accessibility::*;
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::Win32::UI::Input::KeyboardAndMouse::*;

pub struct AccessibilityBridge {
    /// Saved HWND of the target app (set externally when good context is read)
    pub target_hwnd: std::cell::Cell<isize>,
    /// Cached full document text from last successful read
    cached_doc: std::cell::RefCell<String>,
}

impl AccessibilityBridge {
    pub fn new() -> Self {
        AccessibilityBridge {
            target_hwnd: std::cell::Cell::new(0),
            cached_doc: std::cell::RefCell::new(String::new()),
        }
    }

    /// Try to get TextPattern2 from a UIA element
    fn try_read_raw(element: &IUIAutomationElement) -> Option<(RawCursorText, String)> {
        unsafe {
            if let Ok(pattern2) =
                element.GetCurrentPatternAs::<IUIAutomationTextPattern2>(UIA_TextPattern2Id)
            {
                let mut is_active = windows::core::BOOL::default();
                if let Ok(caret_range) = pattern2.GetCaretRange(&mut is_active) {
                    let before_range = caret_range.Clone().ok()?;
                    let _ = before_range.MoveEndpointByUnit(
                        TextPatternRangeEndpoint_Start,
                        TextUnit_Character,
                        -2000,
                    );
                    let before = before_range.GetText(-1).ok()?.to_string();

                    let after_range = caret_range.Clone().ok()?;
                    let _ = after_range.MoveEndpointByUnit(
                        TextPatternRangeEndpoint_End,
                        TextUnit_Character,
                        2000,
                    );
                    let after = after_range.GetText(-1).ok()?.to_string();

                    let doc = format!("{}{}", before, after);
                    return Some((RawCursorText { before, after }, doc));
                }
            }
            None
        }
    }

    /// Get raw text before and after cursor via UIA TextPattern2
    fn get_raw_text(&self) -> Option<RawCursorText> {
        unsafe {
            let uia: IUIAutomation =
                CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER).ok()?;

            // Try focused element first (works when target app has focus)
            if let Ok(focused) = uia.GetFocusedElement() {
                if let Some((raw, doc)) = Self::try_read_raw(&focused) {
                    *self.cached_doc.borrow_mut() = doc;
                    return Some(raw);
                }
            }

            // Fallback: use saved target HWND (works when our window has focus)
            let hwnd_val = self.target_hwnd.get();
            if hwnd_val != 0 {
                let hwnd = HWND(hwnd_val as *mut _);
                if let Ok(element) = uia.ElementFromHandle(hwnd) {
                    // Try the element itself
                    if let Some((raw, doc)) = Self::try_read_raw(&element) {
                        *self.cached_doc.borrow_mut() = doc;
                        return Some(raw);
                    }
                    // Try direct children only (avoids slow deep traversal)
                    if let Ok(condition) = uia.CreateTrueCondition() {
                        if let Ok(children) = element.FindAll(TreeScope_Children, &condition) {
                            let count = children.Length().unwrap_or(0);
                            for i in 0..count.min(10) {
                                if let Ok(child) = children.GetElement(i) {
                                    if let Some((raw, doc)) = Self::try_read_raw(&child) {
                                        *self.cached_doc.borrow_mut() = doc;
                                        return Some(raw);
                                    }
                                }
                            }
                        }
                    }
                }
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
        eprintln!("find_replace_via_uia: find='{}' replace='{}' ctx='{}'", find, replace, context);
        unsafe {
            let pattern = match self.get_text_pattern() {
                Some(p) => { eprintln!("  got TextPattern"); p }
                None => { eprintln!("  FAILED: no TextPattern"); return false; }
            };

            let doc_range = match pattern.DocumentRange() {
                Ok(r) => { eprintln!("  got DocumentRange"); r }
                Err(e) => { eprintln!("  FAILED: DocumentRange: {:?}", e); return false; }
            };

            let find_bstr = windows::core::BSTR::from(find);
            match doc_range.FindText(&find_bstr, false, false) {
                Ok(found_range) => {
                    eprintln!("  FindText found match");
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
                        eprintln!("  context check: surrounding='{}' ctx_words={:?} matches={}", surrounding, ctx_words, matches);
                        if !matches {
                            eprintln!("  FAILED: context mismatch");
                            return false;
                        }
                    }

                    // Return focus to target app before typing
                    let hwnd_val = self.target_hwnd.get();
                    eprintln!("  target_hwnd={}", hwnd_val);
                    if hwnd_val != 0 {
                        let _ = SetForegroundWindow(HWND(hwnd_val as *mut _));
                        std::thread::sleep(std::time::Duration::from_millis(100));
                    }

                    let sel_result = found_range.Select();
                    eprintln!("  Select result: {:?}", sel_result);
                    std::thread::sleep(std::time::Duration::from_millis(50));
                    send_string(replace);
                    eprintln!("  sent replacement string");
                    return true;
                }
                Err(e) => {
                    eprintln!("  FAILED: FindText: {:?}", e);
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
        self.target_hwnd.set(hwnd);
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

/// Type a string by sending keyboard input events
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
    }
}
