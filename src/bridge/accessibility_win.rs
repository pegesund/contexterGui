use super::{CursorContext, TextBridge};
use windows::Win32::Foundation::*;
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::System::Com::*;
use windows::Win32::UI::Accessibility::*;
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::Win32::UI::Input::KeyboardAndMouse::*;

pub struct AccessibilityBridge;

impl AccessibilityBridge {
    pub fn new() -> Self {
        AccessibilityBridge
    }

    fn get_text_pattern(&self) -> Option<(IUIAutomationTextPattern2, bool)> {
        unsafe {
            let uia: IUIAutomation =
                CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER).ok()?;
            let focused = uia.GetFocusedElement().ok()?;

            // Try TextPattern2 first (has GetCaretRange)
            if let Ok(p2) =
                focused.GetCurrentPatternAs::<IUIAutomationTextPattern2>(UIA_TextPattern2Id)
            {
                return Some((p2, true));
            }

            // Fallback to TextPattern (cast to TextPattern2 won't work, but we handle it)
            None
        }
    }

    fn read_from_uia(&self) -> Option<(String, String, Option<String>)> {
        unsafe {
            let uia: IUIAutomation =
                CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER).ok()?;
            let focused = uia.GetFocusedElement().ok()?;

            // Try TextPattern2 first — has GetCaretRange for cursor-at position
            if let Ok(pattern2) =
                focused.GetCurrentPatternAs::<IUIAutomationTextPattern2>(UIA_TextPattern2Id)
            {
                let mut is_active = windows::core::BOOL::default();
                if let Ok(caret_range) = pattern2.GetCaretRange(&mut is_active) {
                    // Expand to word
                    let word_range = caret_range.Clone().ok()?;
                    let _ = word_range.ExpandToEnclosingUnit(TextUnit_Word);
                    let word_raw = word_range.GetText(-1).ok()?.to_string();
                    let word = word_raw.trim().to_string();

                    // Get surrounding context: ±2000 chars for sentence + BERT context
                    let context_range = caret_range.Clone().ok()?;
                    let _ = context_range.MoveEndpointByUnit(
                        TextPatternRangeEndpoint_Start,
                        TextUnit_Character,
                        -2000,
                    );
                    let _ = context_range.MoveEndpointByUnit(
                        TextPatternRangeEndpoint_End,
                        TextUnit_Character,
                        2000,
                    );
                    let context_text = context_range.GetText(-1).ok()?.to_string();

                    // Find sentence around cursor
                    let sentence = find_sentence_containing(&context_text, &word);

                    // Build masked_sentence for BERT completion
                    let masked = if !word.is_empty() && !sentence.is_empty() {
                        build_masked_sentence(&context_text, &word)
                    } else {
                        None
                    };

                    return Some((word, sentence, masked));
                }
            }

            // Fallback: TextPattern with selection
            if let Ok(pattern) =
                focused.GetCurrentPatternAs::<IUIAutomationTextPattern>(UIA_TextPatternId)
            {
                let selection = pattern.GetSelection().ok()?;
                let count = selection.Length().ok()?;
                if count > 0 {
                    let range: IUIAutomationTextRange = selection.GetElement(0).ok()?;
                    let text = range.GetText(-1).ok()?.to_string();
                    if !text.is_empty() {
                        let word = text.split_whitespace().next().unwrap_or("").to_string();
                        return Some((word, text, None));
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

    fn read_document_via_uia(&self) -> Option<String> {
        unsafe {
            let uia: IUIAutomation =
                CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER).ok()?;
            let focused = uia.GetFocusedElement().ok()?;

            if let Ok(pattern) =
                focused.GetCurrentPatternAs::<IUIAutomationTextPattern>(UIA_TextPatternId)
            {
                let doc_range = pattern.DocumentRange().ok()?;
                let text = doc_range.GetText(50000).ok()?.to_string();
                if !text.is_empty() {
                    return Some(text);
                }
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

    /// Replace text by selecting the word range and typing the replacement
    /// Replace word at cursor — try UIA TextPattern2, fall back to keyboard
    fn replace_word_impl(&self, replace_text: &str) -> bool {
        unsafe {
            let uia: IUIAutomation = match CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER) {
                Ok(u) => u,
                Err(e) => {
                    select_word_keyboard();
                    send_string(replace_text);
                    return true;
                }
            };
            let focused = match uia.GetFocusedElement() {
                Ok(f) => f,
                Err(e) => {
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
                    let word_range = caret_range.Clone().unwrap();
                    let _ = word_range.ExpandToEnclosingUnit(TextUnit_Word);
                    let current_word = word_range.GetText(-1).unwrap_or_default().to_string();

                    // Shrink range to exclude trailing whitespace
                    let trimmed = current_word.trim_end();
                    let trailing = current_word.len() - trimmed.len();
                    if trailing > 0 {
                        let _ = word_range.MoveEndpointByUnit(
                            TextPatternRangeEndpoint_End,
                            TextUnit_Character,
                            -(trailing as i32),
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

    /// Find text in document and replace via UIA
    fn find_replace_via_uia(&self, find: &str, replace: &str, context: &str) -> bool {
        unsafe {
            let uia: IUIAutomation = match CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER) {
                Ok(u) => u,
                Err(_) => return false,
            };
            let focused = match uia.GetFocusedElement() {
                Ok(f) => f,
                Err(_) => return false,
            };

            if let Ok(pattern) =
                focused.GetCurrentPatternAs::<IUIAutomationTextPattern>(UIA_TextPatternId)
            {
                let doc_range = match pattern.DocumentRange() {
                    Ok(r) => r,
                    Err(_) => return false,
                };

                // Search for the find text in the document
                let find_bstr = windows::core::BSTR::from(find);
                if let Ok(found_range) = doc_range.FindText(&find_bstr, false, false) {
                    // If we have context, verify it matches
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
                        // Check that context words appear near the found text
                        let ctx_words: Vec<&str> = context.split_whitespace().take(3).collect();
                        let matches = ctx_words.iter().any(|w| surrounding.contains(w));
                        if !matches {
                            return false;
                        }
                    }

                    // Select the found range and type replacement
                    let _ = found_range.Select();
                    send_string(replace);
                    return true;
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
        match self.read_from_uia() {
            Some((word, sentence, masked)) if !word.is_empty() => Some(CursorContext {
                word,
                sentence,
                masked_sentence: masked,
                caret_pos,
            }),
            _ => Some(CursorContext {
                word: String::new(),
                sentence: String::new(),
                masked_sentence: None,
                caret_pos,
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
        self.read_document_via_uia()
    }
}

/// Build a masked sentence for BERT fill-in-the-blank, replacing the word at cursor
fn build_masked_sentence(context: &str, word: &str) -> Option<String> {
    if word.is_empty() {
        return None;
    }
    // Find the sentence containing the word
    let sentence = find_sentence_containing(context, word);
    if sentence.is_empty() {
        return None;
    }

    // Replace last occurrence of word with <mask> (most likely to be the one at cursor)
    if let Some(pos) = sentence.rfind(word) {
        let mut masked = String::with_capacity(sentence.len() + 6);
        masked.push_str(&sentence[..pos]);
        masked.push_str("<mask>");
        masked.push_str(&sentence[pos + word.len()..]);
        Some(masked)
    } else {
        None
    }
}

fn find_sentence_containing(text: &str, word: &str) -> String {
    if word.is_empty() || text.is_empty() {
        return text.trim().to_string();
    }

    // Find where the word appears in the text
    let word_pos = text.find(word).unwrap_or(text.len() / 2);

    let bytes = text.as_bytes();

    // Scan backwards for sentence start
    let mut start = 0;
    for i in (0..word_pos).rev() {
        if i + 1 < bytes.len()
            && (bytes[i] == b'.' || bytes[i] == b'!' || bytes[i] == b'?')
            && (bytes[i + 1] == b' ' || bytes[i + 1] == b'\r' || bytes[i + 1] == b'\n')
        {
            start = i + 1;
            break;
        }
    }

    // Scan forwards for sentence end
    let mut end = bytes.len();
    for i in word_pos..bytes.len() {
        if bytes[i] == b'.' || bytes[i] == b'!' || bytes[i] == b'?' {
            end = i + 1;
            break;
        }
    }

    text[start..end].replace('\r', " ").replace('\n', " ").trim().to_string()
}

/// Select the word at cursor using keyboard shortcuts (Ctrl+Shift+Left)
fn select_word_keyboard() {
    unsafe {
        // First, move to end of word with Ctrl+Right
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

        // Then select word back with Ctrl+Shift+Left
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
        // Small delay to let the selection register
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
