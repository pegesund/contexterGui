use super::{CursorContext, TextBridge};
use windows::Win32::Foundation::*;
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::System::Com::*;
use windows::Win32::UI::Accessibility::*;
use windows::Win32::UI::WindowsAndMessaging::*;

pub struct AccessibilityBridge;

impl AccessibilityBridge {
    pub fn new() -> Self {
        AccessibilityBridge
    }

    fn read_from_uia(&self) -> Option<(String, String)> {
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
                    let word = word_range.GetText(-1).ok()?.to_string();

                    // Get surrounding text: move back/forward several sentences
                    let context_range = caret_range.Clone().ok()?;
                    let _ = context_range.MoveEndpointByUnit(
                        TextPatternRangeEndpoint_Start,
                        TextUnit_Character,
                        -500,
                    );
                    let _ = context_range.MoveEndpointByUnit(
                        TextPatternRangeEndpoint_End,
                        TextUnit_Character,
                        500,
                    );
                    let context_text = context_range.GetText(-1).ok()?.to_string();

                    // Find sentence around cursor in the context
                    let sentence = find_sentence_containing(&context_text, word.trim());
                    return Some((word.trim().to_string(), sentence));
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
                        return Some((word, text));
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
            Some((word, sentence)) if !word.is_empty() => Some(CursorContext {
                word,
                sentence,
                caret_pos,
            }),
            _ => Some(CursorContext {
                word: String::new(),
                sentence: String::new(),
                caret_pos,
            }),
        }
    }

    fn replace_word(&self, _new_text: &str) -> bool {
        false
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
