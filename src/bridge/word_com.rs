#![allow(non_snake_case)]

use super::{CursorContext, TextBridge};
use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use windows::core::*;
use windows::Win32::Foundation::*;
use windows::Win32::System::Com::*;
use windows::Win32::System::Variant::*;
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::UI::Accessibility::*;
use windows::Win32::UI::WindowsAndMessaging::*;

const WD_WORD: i32 = 2;
const OBJID_NATIVEOM: u32 = 0xFFFFFFF0;

// --- Raw VARIANT helpers (COM ABI layout) ---

unsafe fn var_vt(v: &VARIANT) -> u16 {
    unsafe { *(v as *const VARIANT as *const u16) }
}

unsafe fn var_val_ptr(v: &VARIANT) -> *mut std::ffi::c_void {
    unsafe { *((v as *const VARIANT as *const u8).add(8) as *const *mut std::ffi::c_void) }
}

fn make_i4(val: i32) -> VARIANT {
    unsafe {
        let mut v = VARIANT::default();
        let p = &mut v as *mut VARIANT as *mut u8;
        *(p as *mut u16) = VT_I4.0;
        *(p.add(8) as *mut i32) = val;
        v
    }
}

fn make_bstr(s: &str) -> VARIANT {
    unsafe {
        let bstr = BSTR::from(s);
        let mut v = VARIANT::default();
        let p = &mut v as *mut VARIANT as *mut u8;
        *(p as *mut u16) = VT_BSTR.0;
        let raw = bstr.into_raw() as *mut u16;
        *(p.add(8) as *mut *mut u16) = raw;
        v
    }
}

unsafe fn extract_dispatch(v: &VARIANT) -> Result<Dispatch> {
    unsafe {
        if var_vt(v) != VT_DISPATCH.0 {
            return Err(Error::from_hresult(E_FAIL));
        }
        let ptr = var_val_ptr(v);
        if ptr.is_null() {
            return Err(Error::from_hresult(E_FAIL));
        }
        let unk: IUnknown = IUnknown::from_raw_borrowed(&ptr).unwrap().clone();
        let disp: IDispatch = unk.cast()?;
        Ok(Dispatch(disp))
    }
}

unsafe fn extract_string(v: &VARIANT) -> Result<String> {
    unsafe {
        if var_vt(v) != VT_BSTR.0 {
            return Err(Error::from_hresult(E_FAIL));
        }
        let bstr_ptr = var_val_ptr(v) as *const u16;
        if bstr_ptr.is_null() {
            return Ok(String::new());
        }
        let len_bytes = *(bstr_ptr.cast::<u8>().sub(4) as *const u32);
        let len = (len_bytes / 2) as usize;
        let slice = std::slice::from_raw_parts(bstr_ptr, len);
        Ok(String::from_utf16_lossy(slice))
    }
}

unsafe fn extract_i32(v: &VARIANT) -> Result<i32> {
    unsafe {
        if var_vt(v) != VT_I4.0 {
            return Err(Error::from_hresult(E_FAIL));
        }
        Ok(*((v as *const VARIANT as *const u8).add(8) as *const i32))
    }
}

// --- IDispatch late-binding wrapper ---

struct Dispatch(IDispatch);

impl Dispatch {
    fn get_dispid(&self, name: &str) -> Result<i32> {
        let wname: Vec<u16> = OsStr::new(name)
            .encode_wide()
            .chain(std::iter::once(0))
            .collect();
        let names = [PCWSTR(wname.as_ptr())];
        let mut ids = [0i32];
        unsafe {
            self.0
                .GetIDsOfNames(&GUID::zeroed(), names.as_ptr(), 1, 0, ids.as_mut_ptr())?;
        }
        Ok(ids[0])
    }

    fn get(&self, name: &str) -> Result<VARIANT> {
        let id = self.get_dispid(name)?;
        let mut result = VARIANT::default();
        let params = DISPPARAMS::default();
        unsafe {
            self.0.Invoke(
                id, &GUID::zeroed(), 0, DISPATCH_PROPERTYGET,
                &params, Some(&mut result), None, None,
            )?;
        }
        Ok(result)
    }

    fn get_dispatch(&self, name: &str) -> Result<Dispatch> {
        let v = self.get(name)?;
        unsafe { extract_dispatch(&v) }
    }

    fn get_string(&self, name: &str) -> Result<String> {
        let v = self.get(name)?;
        unsafe { extract_string(&v) }
    }

    fn put(&self, name: &str, value: VARIANT) -> Result<()> {
        let id = self.get_dispid(name)?;
        let mut result = VARIANT::default();
        let mut named_arg: i32 = -3; // DISPID_PROPERTYPUT
        let mut args = [value];
        let params = DISPPARAMS {
            rgvarg: args.as_mut_ptr(),
            rgdispidNamedArgs: &mut named_arg,
            cArgs: 1,
            cNamedArgs: 1,
        };
        unsafe {
            self.0.Invoke(
                id, &GUID::zeroed(), 0, DISPATCH_PROPERTYPUT,
                &params, Some(&mut result), None, None,
            )?;
        }
        Ok(())
    }

    fn call(&self, name: &str, args: &[VARIANT]) -> Result<VARIANT> {
        let id = self.get_dispid(name)?;
        let mut result = VARIANT::default();
        let mut reversed: Vec<VARIANT> = args.iter().rev().cloned().collect();
        let params = DISPPARAMS {
            rgvarg: if reversed.is_empty() { std::ptr::null_mut() } else { reversed.as_mut_ptr() },
            cArgs: reversed.len() as u32,
            ..Default::default()
        };
        unsafe {
            self.0.Invoke(
                id, &GUID::zeroed(), 0, DISPATCH_METHOD,
                &params, Some(&mut result), None, None,
            )?;
        }
        Ok(result)
    }
}

// --- Find Word window ---

fn find_wwg_recursive(parent: HWND) -> HWND {
    unsafe {
        let direct = FindWindowExW(Some(parent), None, w!("_WwG"), None)
            .unwrap_or(HWND::default());
        if direct != HWND::default() {
            return direct;
        }
        let mut child = FindWindowExW(Some(parent), None, None, None)
            .unwrap_or(HWND::default());
        while child != HWND::default() {
            let found = find_wwg_recursive(child);
            if found != HWND::default() {
                return found;
            }
            child = FindWindowExW(Some(parent), Some(child), None, None)
                .unwrap_or(HWND::default());
        }
        HWND::default()
    }
}

// --- WordComBridge ---

pub struct WordComBridge {
    word_hwnd: HWND,
    wwg_hwnd: HWND,
}

impl WordComBridge {
    pub fn try_connect() -> Option<Self> {
        unsafe {
            let hwnd = FindWindowW(w!("OpusApp"), None).unwrap_or(HWND::default());
            if hwnd == HWND::default() {
                return None;
            }

            let wwg = find_wwg_recursive(hwnd);
            let target = if wwg != HWND::default() { wwg } else { hwnd };

            // Test that we can get the Application dispatch
            let mut result: *mut std::ffi::c_void = std::ptr::null_mut();
            AccessibleObjectFromWindow(target, OBJID_NATIVEOM, &IDispatch::IID, &mut result)
                .ok()?;
            let disp = IDispatch::from_raw(result);
            let window = Dispatch(disp);
            let _app = window.get_dispatch("Application").ok()?;

            Some(WordComBridge { word_hwnd: hwnd, wwg_hwnd: target })
        }
    }

    fn get_app(&self) -> Option<Dispatch> {
        unsafe {
            let mut result: *mut std::ffi::c_void = std::ptr::null_mut();
            AccessibleObjectFromWindow(self.wwg_hwnd, OBJID_NATIVEOM, &IDispatch::IID, &mut result)
                .ok()?;
            let disp = IDispatch::from_raw(result);
            let window = Dispatch(disp);
            window.get_dispatch("Application").ok()
        }
    }

    fn read_context_data(&self) -> Result<(String, String, Option<String>)> {
        let app = self.get_app().ok_or_else(|| Error::from_hresult(E_FAIL))?;
        let selection = app.get_dispatch("Selection")?;
        let sel_range = selection.get_dispatch("Range")?;
        let cursor_pos = unsafe { extract_i32(&sel_range.get("Start")?) }?;

        // Use Word's own word detection via a duplicated range
        let dup_v = sel_range.get("Duplicate")?;
        let word_range = unsafe { extract_dispatch(&dup_v) }?;
        word_range.call("Expand", &[make_i4(WD_WORD)])?;
        let full_word_text = word_range.get_string("Text").unwrap_or_default();
        let word_start = unsafe { extract_i32(&word_range.get("Start")?) }?;

        // Word's Expand(wdWord) may include trailing space — trim it
        let full_word = full_word_text.trim().to_string();
        let word_end = word_start + full_word.chars().count() as i32;
        let is_mid_word = cursor_pos > word_start && cursor_pos < word_end
            && full_word.chars().all(|c| c.is_alphanumeric());

        let doc = app.get_dispatch("ActiveDocument")?;

        // Check if cursor is past the word (e.g. after pressing space)
        // Word's Expand(wdWord) includes trailing space, so cursor_pos may be at word_end
        // but the actual cursor is in the space after the word
        let cursor_past_word = cursor_pos >= word_end && {
            // Check if char at cursor_pos is space/punct (not alphanumeric)
            let peek_v = doc.call("Range", &[make_i4(cursor_pos), make_i4(cursor_pos + 1)]);
            peek_v.ok().and_then(|v| {
                let r = unsafe { extract_dispatch(&v) }.ok()?;
                let t = r.get_string("Text").ok()?;
                t.chars().next().map(|c| !c.is_alphanumeric())
            }).unwrap_or(true) // at end of doc = past word
        };

        // If cursor is at word start, past word, or non-alphanumeric word → empty
        let word = if cursor_pos == word_start || cursor_past_word || !full_word.chars().all(|c| c.is_alphanumeric()) {
            // Expand(wdWord) failed to find a word — fallback: scan backwards from cursor
            let look_back = 50.min(cursor_pos);
            if look_back > 0 {
                let fb_v = doc.call("Range", &[make_i4(cursor_pos - look_back), make_i4(cursor_pos)])?;
                let fb_range = unsafe { extract_dispatch(&fb_v) }?;
                let fb_text = fb_range.get_string("Text").unwrap_or_default();
                let fallback_word: String = fb_text.chars().rev()
                    .take_while(|c| c.is_alphanumeric() || *c == '-' || *c == '\'')
                    .collect::<Vec<_>>().into_iter().rev().collect();
                if fallback_word.is_empty() { String::new() } else { fallback_word }
            } else {
                String::new()
            }
        } else if is_mid_word {
            // Cursor is mid-word (clicked on existing word) — use first letter only
            full_word.chars().next().map(|c| c.to_string()).unwrap_or_default()
        } else {
            // Cursor at end of word (typing) — use full word
            full_word.clone()
        };

        // Build masked context for fill-in-the-blank (BERT with both sides of context)
        // Used for: mid-word clicks AND typing in the middle of existing text
        let content_for_mask = doc.get_dispatch("Content")?;
        let doc_end_for_mask = unsafe { extract_i32(&content_for_mask.get("End")?) }?;
        let has_text_after = {
            let peek_end = (cursor_pos + 20).min(doc_end_for_mask);
            if peek_end > cursor_pos {
                let peek_v = doc.call("Range", &[make_i4(cursor_pos), make_i4(peek_end)]);
                peek_v.ok().and_then(|v| {
                    let r = unsafe { extract_dispatch(&v) }.ok()?;
                    let t = r.get_string("Text").ok()?;
                    let trimmed = t.trim();
                    // Has meaningful text after cursor (not just whitespace/empty)
                    Some(trimmed.len() > 2)
                }).unwrap_or(false)
            } else {
                false
            }
        };
        let has_context_before = cursor_pos > 1;
        let use_masked = is_mid_word || has_text_after || (word.is_empty() && has_context_before);
        let masked_sentence = if use_masked {
            let half_ctx = 2000;
            // For typing: back up past the partial word so <mask> replaces it
            let typed_len = word.chars().count() as i32;
            let mask_start = if is_mid_word { word_start } else { (cursor_pos - typed_len).max(0) };
            let mask_end = if is_mid_word { word_end } else { cursor_pos };
            let ctx_start = (mask_start - half_ctx).max(0);
            let ctx_end = (mask_end + half_ctx).min(doc_end_for_mask);
            let before_v = doc.call("Range", &[make_i4(ctx_start), make_i4(mask_start)])?;
            let before_range = unsafe { extract_dispatch(&before_v) }?;
            let before = before_range.get_string("Text").unwrap_or_default()
                .replace('\r', " ").replace('\n', " ");
            let after_v = doc.call("Range", &[make_i4(mask_end), make_i4(ctx_end)])?;
            let after_range = unsafe { extract_dispatch(&after_v) }?;
            let after = after_range.get_string("Text").unwrap_or_default()
                .replace('\r', " ").replace('\n', " ");
            let masked = format!("{}<mask> {}", before.trim(), after.trim());
            Some(masked)
        } else {
            None
        };

        // Sentence: read text around cursor (before + after for full sentence display)
        let content = doc.get_dispatch("Content")?;
        let doc_end_val = unsafe { extract_i32(&content.get("End")?) }?;
        let sent_start = (cursor_pos - 2000).max(0);
        let sent_end = (cursor_pos + 1000).min(doc_end_val);
        let ctx_v = doc.call("Range", &[make_i4(sent_start), make_i4(sent_end)])?;
        let ctx_range = unsafe { extract_dispatch(&ctx_v) }?;
        let around_text = ctx_range.get_string("Text")?.replace('\r', " ").replace('\n', " ");
        let char_offset = (cursor_pos - sent_start) as usize;
        let sentence = find_sentence_at_offset(&around_text, char_offset);

        Ok((word, sentence, masked_sentence))
    }

    fn caret_pos(&self) -> Option<(i32, i32)> {
        unsafe {
            let thread_id = GetWindowThreadProcessId(self.word_hwnd, None);
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

    fn is_foreground(&self) -> bool {
        unsafe {
            let fg = GetForegroundWindow();
            if fg == self.word_hwnd {
                return true;
            }
            let mut fg_pid = 0u32;
            let mut word_pid = 0u32;
            GetWindowThreadProcessId(fg, Some(&mut fg_pid));
            GetWindowThreadProcessId(self.word_hwnd, Some(&mut word_pid));
            fg_pid == word_pid
        }
    }
}

impl TextBridge for WordComBridge {
    fn name(&self) -> &str {
        "Word COM"
    }

    fn is_available(&self) -> bool {
        // COM calls work even when Word isn't foreground,
        // so always return true if we have a connection
        true
    }

    fn read_context(&self) -> Option<CursorContext> {
        let (word, sentence, masked_sentence) = match self.read_context_data() {
            Ok(data) => data,
            Err(_) => (String::new(), String::new(), None),
        };
        let caret_pos = self.caret_pos();
        Some(CursorContext {
            word: word.trim().to_string(),
            sentence,
            masked_sentence,
            caret_pos,
        })
    }

    fn read_document_context(&self) -> Option<String> {
        let app = self.get_app()?;
        let selection = app.get_dispatch("Selection").ok()?;
        let sel_range = selection.get_dispatch("Range").ok()?;
        let cursor_pos = unsafe { extract_i32(&sel_range.get("Start").ok()?).ok()? };

        let range_start = (cursor_pos - 5000).max(0);
        let doc = app.get_dispatch("ActiveDocument").ok()?;
        let context_v = doc.call("Range", &[make_i4(range_start), make_i4(cursor_pos)]).ok()?;
        let context_range = unsafe { extract_dispatch(&context_v).ok()? };
        context_range.get_string("Text").ok()
    }

    fn replace_word(&self, new_text: &str) -> bool {
        (|| -> Result<()> {
            let app = self.get_app().ok_or_else(|| Error::from_hresult(E_FAIL))?;
            let selection = app.get_dispatch("Selection")?;
            let sel_range = selection.get_dispatch("Range")?;
            let cursor_pos = unsafe { extract_i32(&sel_range.get("Start")?) }?;

            let doc = app.get_dispatch("ActiveDocument")?;
            let content = doc.get_dispatch("Content")?;
            let doc_end = unsafe { extract_i32(&content.get("End")?) }?;

            // Read text around cursor to find full word boundaries
            let look_back = 50.min(cursor_pos);
            let look_ahead = 50.min(doc_end - cursor_pos);
            let range_start = cursor_pos - look_back;
            let range_end = cursor_pos + look_ahead;
            let ctx_v = doc.call("Range", &[make_i4(range_start), make_i4(range_end)])?;
            let ctx_range = unsafe { extract_dispatch(&ctx_v) }?;
            let around_text = ctx_range.get_string("Text")?;

            let chars: Vec<char> = around_text.chars().collect();
            let cursor_offset = look_back as usize;

            // Scan backwards from cursor to find word start
            let mut word_start_off = cursor_offset;
            while word_start_off > 0 && chars[word_start_off - 1].is_alphanumeric() {
                word_start_off -= 1;
            }

            // Scan forwards from cursor to find word end
            let mut word_end_off = cursor_offset;
            while word_end_off < chars.len() && chars[word_end_off].is_alphanumeric() {
                word_end_off += 1;
            }

            if word_start_off == word_end_off {
                // No word at cursor — insert the word at cursor position
                let insert_range_v = doc.call("Range", &[make_i4(cursor_pos), make_i4(cursor_pos)])?;
                let insert_range = unsafe { extract_dispatch(&insert_range_v) }?;
                insert_range.put("Text", make_bstr(&format!("{} ", new_text)))?;
                let new_end = cursor_pos + new_text.chars().count() as i32 + 1;
                let cursor_v = doc.call("Range", &[make_i4(new_end), make_i4(new_end)])?;
                let cursor_range = unsafe { extract_dispatch(&cursor_v) }?;
                cursor_range.call("Select", &[])?;
                return Ok(());
            }

            // Create range covering the full word
            let word_start = range_start + word_start_off as i32;
            let word_end = range_start + word_end_off as i32;
            let word_range_v = doc.call("Range", &[make_i4(word_start), make_i4(word_end)])?;
            let word_range = unsafe { extract_dispatch(&word_range_v) }?;
            word_range.put("Text", make_bstr(new_text))?;
            // Move cursor to end of inserted word
            let new_end = word_start + new_text.chars().count() as i32;
            let cursor_v = doc.call("Range", &[make_i4(new_end), make_i4(new_end)])?;
            let cursor_range = unsafe { extract_dispatch(&cursor_v) }?;
            cursor_range.call("Select", &[])?;
            Ok(())
        })()
        .is_ok()
    }
}

// --- Sentence detection ---

fn find_word_at_offset(text: &str, char_offset: usize) -> String {
    let chars: Vec<char> = text.chars().collect();
    if chars.is_empty() {
        return String::new();
    }
    let pos = char_offset.min(chars.len());

    // Cursor is BETWEEN characters. Check the char just before the cursor.
    // If it's alphanumeric, we're at the end of a word (typing).
    // If it's space/punct, we're at a word boundary.
    if pos == 0 {
        return String::new();
    }

    let prev = chars[pos - 1];
    if !prev.is_alphanumeric() {
        // At a word boundary (just typed space or punct)
        return String::new();
    }

    // Scan backwards from pos-1 to find word start
    let mut start = pos - 1;
    while start > 0 && chars[start - 1].is_alphanumeric() {
        start -= 1;
    }

    // The word ends at pos (cursor position) — don't scan forward
    chars[start..pos].iter().collect::<String>().trim().to_string()
}

fn find_sentence_at_offset(text: &str, char_offset: usize) -> String {
    let byte_offset = text.char_indices()
        .nth(char_offset)
        .map(|(b, _)| b)
        .unwrap_or(text.len());

    let bytes = text.as_bytes();
    let offset = byte_offset.min(bytes.len().saturating_sub(1));

    let mut start = 0;
    for i in (0..offset).rev() {
        if i + 1 < bytes.len()
            && (bytes[i] == b'.' || bytes[i] == b'!' || bytes[i] == b'?')
            && (bytes[i + 1] == b' ' || bytes[i + 1] == b'\r' || bytes[i + 1] == b'\n')
        {
            start = i + 1;
            break;
        }
    }

    let mut end = bytes.len();
    for i in offset..bytes.len() {
        if bytes[i] == b'.' || bytes[i] == b'!' || bytes[i] == b'?' {
            end = i + 1;
            break;
        }
    }

    text[start..end].replace('\r', " ").replace('\n', " ").trim().to_string()
}
