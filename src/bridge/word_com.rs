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
const WD_UNDERLINE_WAVY: i32 = 11;
const WD_UNDERLINE_NONE: i32 = 0;
const WD_COLOR_RED: i32 = 0x0000FF;  // BGR format: red = 0x0000FF
const WD_COLOR_BLUE: i32 = 0xFF0000; // BGR format: blue = 0xFF0000
const WD_NORWEGIAN_BOKMAL: i32 = 1044; // wdNorwegianBokmal language ID

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

fn make_bool(val: bool) -> VARIANT {
    unsafe {
        let mut v = VARIANT::default();
        let p = &mut v as *mut VARIANT as *mut u8;
        *(p as *mut u16) = VT_BOOL.0;
        *(p.add(8) as *mut i16) = if val { -1 } else { 0 }; // VARIANT_TRUE=-1, VARIANT_FALSE=0
        v
    }
}

unsafe fn extract_bool(v: &VARIANT) -> Result<bool> {
    unsafe {
        let vt = var_vt(v);
        if vt == VT_BOOL.0 {
            let val = *(var_val_ptr(v) as *const i16);
            Ok(val != 0)
        } else if vt == VT_I4.0 {
            let val = *(var_val_ptr(v) as *const i32);
            Ok(val != 0)
        } else {
            Err(Error::from_hresult(E_FAIL))
        }
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

/// Extract a numeric value from a VARIANT, handling VT_I4, VT_R4, and VT_R8.
unsafe fn extract_numeric(v: &VARIANT) -> Result<f64> {
    unsafe {
        let vt = var_vt(v);
        let data = (v as *const VARIANT as *const u8).add(8);
        if vt == VT_I4.0 {
            Ok(*(data as *const i32) as f64)
        } else if vt == VT_R4.0 {
            Ok(*(data as *const f32) as f64)
        } else if vt == VT_R8.0 {
            Ok(*(data as *const f64))
        } else {
            Err(Error::from_hresult(E_FAIL))
        }
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
        // COM expects args in reverse order.
        // Use MaybeUninit<VARIANT> for proper alignment, then forget to prevent Drop
        // (avoids BSTR double-free — originals in caller's stack get dropped once).
        let n = args.len();
        let mut reversed: Vec<std::mem::MaybeUninit<VARIANT>> = Vec::with_capacity(n);
        for arg in args.iter().rev() {
            let mut slot = std::mem::MaybeUninit::<VARIANT>::uninit();
            unsafe {
                std::ptr::copy_nonoverlapping(
                    arg as *const VARIANT,
                    slot.as_mut_ptr(),
                    1,
                );
            }
            reversed.push(slot);
        }
        let params = DISPPARAMS {
            rgvarg: if n == 0 { std::ptr::null_mut() } else { reversed.as_mut_ptr() as *mut VARIANT },
            cArgs: n as u32,
            ..Default::default()
        };
        unsafe {
            self.0.Invoke(
                id, &GUID::zeroed(), 0, DISPATCH_METHOD,
                &params, Some(&mut result), None, None,
            )?;
        }
        // MaybeUninit doesn't run Drop — no double-free
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

    fn get_raw_text(&self) -> Result<(super::RawCursorText, usize)> {
        let app = self.get_app().ok_or_else(|| Error::from_hresult(E_FAIL))?;
        let selection = app.get_dispatch("Selection")?;
        let sel_range = selection.get_dispatch("Range")?;
        let cursor_pos = unsafe { extract_i32(&sel_range.get("Start")?) }?;

        let doc = app.get_dispatch("ActiveDocument")?;
        let content = doc.get_dispatch("Content")?;
        let doc_end = unsafe { extract_i32(&content.get("End")?) }?;

        // Read text before cursor (up to 2000 chars)
        let before_start = (cursor_pos - 2000).max(0);
        let before_v = doc.call("Range", &[make_i4(before_start), make_i4(cursor_pos)])?;
        let before_range = unsafe { extract_dispatch(&before_v) }?;
        let before = before_range.get_string("Text").unwrap_or_default()
            .replace('\r', " ").replace('\n', " ");

        // Read text after cursor (up to 2000 chars)
        let after_end = (cursor_pos + 2000).min(doc_end);
        let after_v = doc.call("Range", &[make_i4(cursor_pos), make_i4(after_end)])?;
        let after_range = unsafe { extract_dispatch(&after_v) }?;
        let after = after_range.get_string("Text").unwrap_or_default()
            .replace('\r', " ").replace('\n', " ");

        Ok((super::RawCursorText { before, after }, cursor_pos as usize))
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

    fn is_word_foreground(&self) -> bool {
        unsafe {
            let fg = GetForegroundWindow();
            let mut fg_pid = 0u32;
            GetWindowThreadProcessId(fg, Some(&mut fg_pid));
            let mut word_pid = 0u32;
            GetWindowThreadProcessId(self.word_hwnd, Some(&mut word_pid));
            fg_pid == word_pid
        }
    }

    /// Disable Word's built-in spell/grammar checking so our underlines don't conflict.
    pub fn disable_word_proofing(&self) -> bool {
        let result = (|| -> Result<bool> {
            let app = self.get_app().ok_or_else(|| Error::from_hresult(E_FAIL))?;
            // Application-level: turn off as-you-type checking
            let options = app.get_dispatch("Options")?;
            options.put("CheckSpellingAsYouType", make_bool(false))?;
            options.put("CheckGrammarAsYouType", make_bool(false))?;
            options.put("CheckGrammarWithSpelling", make_bool(false))?;
            eprintln!("Word: disabled Options-level proofing");
            // Document-level: hide existing squiggles
            let doc = app.get_dispatch("ActiveDocument")?;
            doc.put("ShowSpellingErrors", make_bool(false))?;
            doc.put("ShowGrammaticalErrors", make_bool(false))?;
            // Tell Word the document is already checked (clears existing marks)
            doc.put("SpellingChecked", make_bool(true))?;
            doc.put("GrammarChecked", make_bool(true))?;
            // Set document content language to "No Proofing" (wdNoProofing=1024)
            // This is the most reliable way to suppress Word 365 Editor squiggles
            match doc.get_dispatch("Content") {
                Ok(content) => {
                    match content.put("LanguageID", make_i4(1024)) {
                        Ok(_) => eprintln!("Word: Content.LanguageID = wdNoProofing OK"),
                        Err(e) => eprintln!("Word: Content.LanguageID FAILED: {:?}", e),
                    }
                    match content.put("NoProofing", make_bool(true)) {
                        Ok(_) => eprintln!("Word: Content.NoProofing = True OK"),
                        Err(e) => eprintln!("Word: Content.NoProofing FAILED: {:?}", e),
                    }
                }
                Err(e) => eprintln!("Word: get Content FAILED: {:?}", e),
            }
            eprintln!("Word: disabled document-level squiggles");
            Ok(true)
        })();
        match &result {
            Ok(_) => eprintln!("Word: disabled built-in spell/grammar checking"),
            Err(e) => eprintln!("Word: FAILED to disable proofing: {:?}", e),
        }
        result.unwrap_or(false)
    }

    /// Re-enable Word's built-in spell/grammar checking.
    pub fn enable_word_proofing(&self) -> bool {
        (|| -> Result<bool> {
            let app = self.get_app().ok_or_else(|| Error::from_hresult(E_FAIL))?;
            let options = app.get_dispatch("Options")?;
            options.put("CheckSpellingAsYouType", make_bool(true))?;
            options.put("CheckGrammarAsYouType", make_bool(true))?;
            options.put("CheckGrammarWithSpelling", make_bool(true))?;
            let doc = app.get_dispatch("ActiveDocument")?;
            doc.put("ShowSpellingErrors", make_bool(true))?;
            doc.put("ShowGrammaticalErrors", make_bool(true))?;
            // Restore Norwegian Bokmål language
            if let Ok(content) = doc.get_dispatch("Content") {
                let _ = content.put("LanguageID", make_i4(WD_NORWEGIAN_BOKMAL));
                let _ = content.put("NoProofing", make_i4(0)); // VARIANT_FALSE
            }
            eprintln!("Word: re-enabled built-in spell/grammar checking");
            Ok(true)
        })().unwrap_or(false)
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
        // Word COM works via COM calls regardless of which window has focus
        true
    }

    fn read_context(&self) -> Option<CursorContext> {
        // Read text from Word via COM. Caret position is optional — when our
        // always-on-top window has focus, Word's caret disappears but the
        // text and cursor position are still available via COM.
        let caret_pos = self.caret_pos(); // None when our window has focus — OK
        match self.get_raw_text() {
            Ok((raw, cursor_offset)) => {
                let mut ctx = super::build_context(&raw, caret_pos);
                ctx.cursor_doc_offset = Some(cursor_offset);
                Some(ctx)
            }
            Err(e) => {
                log!("Word COM get_raw_text failed: {:?}", e);
                None
            }
        }
    }

    fn read_document_context(&self) -> Option<String> {
        // No blinking caret = no cursor = don't read
        if self.caret_pos().is_none() { return None; }
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

    fn read_full_document(&self) -> Option<String> {
        if self.caret_pos().is_none() { return None; }
        let app = self.get_app()?;
        let doc = app.get_dispatch("ActiveDocument").ok()?;
        let content = doc.get_dispatch("Content").ok()?;
        let text = content.get_string("Text").ok()?;
        Some(text.replace('\r', " "))
    }


    fn select_range(&self, char_start: usize, char_end: usize) -> Option<(i32, i32)> {
        // Step 1: Select the range in Word (clamp to doc length)
        let select_ok = (|| -> Result<()> {
            let app = self.get_app().ok_or_else(|| Error::from_hresult(E_FAIL))?;
            let doc = app.get_dispatch("ActiveDocument")?;
            let content = doc.get_dispatch("Content")?;
            let doc_end = unsafe { extract_i32(&content.get("End")?) }.unwrap_or(99999) as usize;
            let start = char_start.min(doc_end.saturating_sub(1));
            let end = char_end.min(doc_end);
            let range_v = doc.call("Range", &[make_i4(start as i32), make_i4(end as i32)])?;
            let range = unsafe { extract_dispatch(&range_v)? };
            range.call("Select", &[])?;
            app.call("Activate", &[])?;
            Ok(())
        })();
        if let Err(e) = select_ok {
            log!("select_range: Select failed: {:?}", e);
            return None;
        }

        // Step 2: Get screen position of the selection via GetGUIThreadInfo
        // After Select+Activate, Word's thread has the caret at the selection
        unsafe {
            use windows::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId;
            let thread_id = GetWindowThreadProcessId(self.word_hwnd, None);
            if thread_id == 0 { return None; }
            let mut gui = GUITHREADINFO::default();
            gui.cbSize = std::mem::size_of::<GUITHREADINFO>() as u32;
            if GetGUIThreadInfo(thread_id, &mut gui).is_ok() && gui.rcCaret.bottom > 0 {
                // rcCaret is in client coordinates of the focused window
                let mut pt = POINT { x: gui.rcCaret.left, y: gui.rcCaret.bottom };
                let focus_hwnd = if gui.hwndFocus != HWND::default() { gui.hwndFocus } else { self.word_hwnd };
                let _ = ClientToScreen(focus_hwnd, &mut pt);
                log!("select_range: caret screen pos ({}, {})", pt.x, pt.y);
                return Some((pt.x, pt.y + 5));
            }
            // Fallback: center of Word window
            let mut rect = RECT::default();
            let _ = GetWindowRect(self.word_hwnd, &mut rect);
            let cx = (rect.left + rect.right) / 2;
            let cy = (rect.top + rect.bottom) / 2;
            log!("select_range: using Word window center ({}, {})", cx, cy);
            Some((cx, cy))
        }
    }

    fn replace_word(&self, new_text: &str) -> bool {
        // The Mac Word Add-in sends "prefix|replacement" format. Extract replacement only.
        let replacement = if let Some((_prefix, word)) = new_text.split_once('|') {
            word
        } else {
            new_text
        };
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
                insert_range.put("Text", make_bstr(&format!("{} ", replacement)))?;
                let new_end = cursor_pos + replacement.chars().count() as i32 + 1;
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
            word_range.put("Text", make_bstr(replacement))?;
            // Move cursor to end of inserted word
            let new_end = word_start + replacement.chars().count() as i32;
            let cursor_v = doc.call("Range", &[make_i4(new_end), make_i4(new_end)])?;
            let cursor_range = unsafe { extract_dispatch(&cursor_v) }?;
            cursor_range.call("Select", &[])?;
            Ok(())
        })()
        .is_ok()
    }

    fn find_and_replace(&self, find_text: &str, replace_text: &str) -> bool {
        // Use Range-based search instead of Find.Execute to avoid VARIANT/BSTR issues.
        // Scan document text for the word and replace via Range.Text property.
        (|| -> Result<bool> {
            let app = self.get_app().ok_or_else(|| Error::from_hresult(E_FAIL))?;
            let doc = app.get_dispatch("ActiveDocument")?;
            let content = doc.get_dispatch("Content")?;
            let doc_text = content.get_string("Text")?;

            // Find the word (case-insensitive, whole word)
            let find_lower = find_text.to_lowercase();
            let chars: Vec<char> = doc_text.chars().collect();
            let find_chars: Vec<char> = find_lower.chars().collect();
            let find_len = find_chars.len();

            for i in 0..chars.len().saturating_sub(find_len - 1) {
                // Check word boundary before
                if i > 0 && chars[i - 1].is_alphanumeric() {
                    continue;
                }
                // Check word boundary after
                let end = i + find_len;
                if end < chars.len() && chars[end].is_alphanumeric() {
                    continue;
                }
                // Check match (case-insensitive)
                let candidate: String = chars[i..end].iter().collect();
                if candidate.to_lowercase() != find_lower {
                    continue;
                }
                // Found! Replace via Range.Text
                let range_v = doc.call("Range", &[make_i4(i as i32), make_i4(end as i32)])?;
                let range = unsafe { extract_dispatch(&range_v) }?;
                range.put("Text", make_bstr(replace_text))?;
                eprintln!("Find&Replace: '{}' → '{}' at position {}", find_text, replace_text, i);
                return Ok(true);
            }
            eprintln!("Find&Replace: '{}' not found in document", find_text);
            Ok(false)
        })()
        .unwrap_or(false)
    }

    /// Replace a word only within the context of a specific sentence.
    /// First finds the sentence in the document, then finds the word within that sentence range.
    fn find_and_replace_in_context(&self, find_text: &str, replace_text: &str, sentence_context: &str) -> bool {
        (|| -> Result<bool> {
            let app = self.get_app().ok_or_else(|| Error::from_hresult(E_FAIL))?;
            let doc = app.get_dispatch("ActiveDocument")?;
            let content = doc.get_dispatch("Content")?;
            let doc_text = content.get_string("Text")?;

            let doc_lower = doc_text.to_lowercase();
            let ctx_lower = sentence_context.to_lowercase();

            // Find the sentence in the document
            let ctx_start = match doc_lower.find(&ctx_lower) {
                Some(pos) => pos,
                None => {
                    eprintln!("Find&Replace: sentence context not found in document");
                    return Ok(false);
                }
            };
            let ctx_end = ctx_start + sentence_context.len();

            // Find the word within the sentence range
            let find_lower = find_text.to_lowercase();
            let chars: Vec<char> = doc_text.chars().collect();
            let find_chars: Vec<char> = find_lower.chars().collect();
            let find_len = find_chars.len();

            // Convert byte offsets to char offsets for the sentence range
            let ctx_char_start = doc_text[..ctx_start].chars().count();
            let ctx_char_end = ctx_char_start + doc_text[ctx_start..ctx_end].chars().count();

            for i in ctx_char_start..ctx_char_end.saturating_sub(find_len - 1) {
                if i > 0 && chars[i - 1].is_alphanumeric() { continue; }
                let end = i + find_len;
                if end < chars.len() && chars[end].is_alphanumeric() { continue; }
                let candidate: String = chars[i..end].iter().collect();
                if candidate.to_lowercase() != find_lower { continue; }
                let range_v = doc.call("Range", &[make_i4(i as i32), make_i4(end as i32)])?;
                let range = unsafe { extract_dispatch(&range_v) }?;
                range.put("Text", make_bstr(replace_text))?;
                eprintln!("Find&Replace: '{}' → '{}' at position {} (in context)", find_text, replace_text, i);
                return Ok(true);
            }
            eprintln!("Find&Replace: '{}' not found in sentence context", find_text);
            Ok(false)
        })()
        .unwrap_or(false)
    }

    fn find_and_replace_in_context_at(&self, find_text: &str, replace_text: &str, sentence_context: &str, char_offset: usize) -> bool {
        (|| -> Result<bool> {
            let app = self.get_app().ok_or_else(|| Error::from_hresult(E_FAIL))?;
            let doc = app.get_dispatch("ActiveDocument")?;
            let content = doc.get_dispatch("Content")?;
            let doc_text = content.get_string("Text")?;

            // Normalize \r to space for searching (sentence_context uses spaces)
            let doc_lower = doc_text.replace('\r', " ").to_lowercase();
            let ctx_lower = sentence_context.to_lowercase();

            // Find the sentence near the known char offset
            // Convert char_offset to byte offset for string searching
            let byte_offset: usize = doc_text.chars().take(char_offset).map(|c| c.len_utf8()).sum();
            // Search from a bit before the expected position (allow some drift)
            let search_start = byte_offset.saturating_sub(50);
            let ctx_start = match doc_lower[search_start..].find(&ctx_lower) {
                Some(pos) => search_start + pos,
                None => {
                    // Fallback: search from beginning
                    match doc_lower.find(&ctx_lower) {
                        Some(pos) => pos,
                        None => {
                            eprintln!("Find&Replace@offset: sentence context not found in document");
                            return Ok(false);
                        }
                    }
                }
            };
            let ctx_end = ctx_start + sentence_context.len();

            // Find the word within the sentence range
            let find_lower = find_text.to_lowercase();
            let chars: Vec<char> = doc_text.chars().collect();
            let find_chars: Vec<char> = find_lower.chars().collect();
            let find_len = find_chars.len();

            let ctx_char_start = doc_text[..ctx_start].chars().count();
            let ctx_char_end = ctx_char_start + doc_text[ctx_start..ctx_end].chars().count();

            for i in ctx_char_start..ctx_char_end.saturating_sub(find_len - 1) {
                if i > 0 && chars[i - 1].is_alphanumeric() { continue; }
                let end = i + find_len;
                if end < chars.len() && chars[end].is_alphanumeric() { continue; }
                let candidate: String = chars[i..end].iter().collect();
                if candidate.to_lowercase() != find_lower { continue; }
                let range_v = doc.call("Range", &[make_i4(i as i32), make_i4(end as i32)])?;
                let range = unsafe { extract_dispatch(&range_v) }?;
                range.put("Text", make_bstr(replace_text))?;
                eprintln!("Find&Replace@offset: '{}' → '{}' at position {} (char_offset hint={})", find_text, replace_text, i, char_offset);
                return Ok(true);
            }
            eprintln!("Find&Replace@offset: '{}' not found in sentence context", find_text);
            Ok(false)
        })()
        .unwrap_or(false)
    }

    fn mark_error_underline(&self, char_start: usize, char_end: usize, color: super::ErrorUnderlineColor) -> bool {
        (|| -> Result<bool> {
            let app = self.get_app().ok_or_else(|| Error::from_hresult(E_FAIL))?;
            let doc = app.get_dispatch("ActiveDocument")?;
            let range_v = doc.call("Range", &[make_i4(char_start as i32), make_i4(char_end as i32)])?;
            let range = unsafe { extract_dispatch(&range_v)? };
            let font = range.get_dispatch("Font")?;
            font.put("Underline", make_i4(WD_UNDERLINE_WAVY))?;
            let bgr_color = match color {
                super::ErrorUnderlineColor::Red => WD_COLOR_RED,
                super::ErrorUnderlineColor::Blue => WD_COLOR_BLUE,
            };
            font.put("UnderlineColor", make_i4(bgr_color))?;
            Ok(true)
        })().unwrap_or(false)
    }

    fn clear_error_underline(&self, char_start: usize, char_end: usize) -> bool {
        (|| -> Result<bool> {
            let app = self.get_app().ok_or_else(|| Error::from_hresult(E_FAIL))?;
            let doc = app.get_dispatch("ActiveDocument")?;
            let range_v = doc.call("Range", &[make_i4(char_start as i32), make_i4(char_end as i32)])?;
            let range = unsafe { extract_dispatch(&range_v)? };
            let font = range.get_dispatch("Font")?;
            font.put("Underline", make_i4(WD_UNDERLINE_NONE))?;
            Ok(true)
        })().unwrap_or(false)
    }

    fn clear_all_error_underlines(&self) -> bool {
        // Set entire document to no underline.
        // Called on app exit to clean up our markings.
        (|| -> Result<bool> {
            let app = self.get_app().ok_or_else(|| Error::from_hresult(E_FAIL))?;
            let doc = app.get_dispatch("ActiveDocument")?;
            let content = doc.get_dispatch("Content")?;
            let font = content.get_dispatch("Font")?;
            font.put("Underline", make_i4(WD_UNDERLINE_NONE))?;
            Ok(true)
        })().unwrap_or(false)
    }
}

