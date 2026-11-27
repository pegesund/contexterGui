// Classic Notepad exposes a standard Edit control, even on systems where its
// accessibility provider offers ValuePattern but no TextPattern.
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::*;

fn class_is(window: HWND, expected: &str) -> bool {
    let mut class = [0; 64];
    let length = unsafe { GetClassNameW(window, &mut class) };
    class[..length as usize]
        .iter()
        .copied()
        .eq(expected.encode_utf16())
}

pub(crate) fn editor(root: HWND) -> Option<HWND> {
    unsafe {
        if !class_is(root, "Notepad") || GetForegroundWindow() != root {
            return None;
        }
        let mut info = GUITHREADINFO {
            cbSize: std::mem::size_of::<GUITHREADINFO>() as u32,
            ..Default::default()
        };
        GetGUIThreadInfo(GetWindowThreadProcessId(root, None), &mut info).ok()?;
        let edit = info.hwndFocus;
        (class_is(edit, "Edit") && GetParent(edit).ok() == Some(root)).then_some(edit)
    }
}

/// A toolbar click takes focus away from the editor. Classic Notepad has one
/// direct Edit child; use only that child of the last observed document window.
pub(crate) fn correction_editor(root: HWND) -> Option<HWND> {
    use windows::core::w;
    unsafe {
        if !class_is(root, "Notepad") {
            return None;
        }
        let fg = GetForegroundWindow();
        let mut pid = 0;
        GetWindowThreadProcessId(fg, Some(&mut pid));
        if fg != root && pid != std::process::id() {
            return None;
        }
        let edit = FindWindowExW(Some(root), None, w!("Edit"), None).ok()?;
        if !IsWindowVisible(edit).as_bool()
            || FindWindowExW(Some(root), Some(edit), w!("Edit"), None).is_ok()
        {
            return None;
        }
        Some(edit)
    }
}

fn message(edit: HWND, id: u32, w: usize, l: isize) -> Option<usize> {
    let mut result = 0;
    unsafe {
        SendMessageTimeoutW(
            edit,
            id,
            WPARAM(w),
            LPARAM(l),
            SMTO_ABORTIFHUNG,
            250,
            Some(&mut result),
        )
        .0
        .ne(&0)
        .then_some(result)
    }
}

fn text(edit: HWND) -> Option<Vec<u16>> {
    let len = message(edit, WM_GETTEXTLENGTH, 0, 0)?;
    if len > 4_000_000 {
        return None;
    }
    let mut buf = vec![0; len + 1];
    let copied = message(edit, WM_GETTEXT, buf.len(), buf.as_mut_ptr() as isize)?;
    buf.truncate(copied.min(len));
    Some(buf)
}

fn selection(edit: HWND) -> Option<(usize, usize)> {
    let (mut start, mut end) = (0u32, 0u32);
    message(
        edit,
        0x00B0,
        (&mut start as *mut u32) as usize,
        (&mut end as *mut u32) as isize,
    )?;
    Some((start as usize, end as usize))
}

pub(crate) fn selected_text(edit: HWND) -> Option<String> {
    let (start, end) = selection(edit)?;
    if start == end {
        return None;
    }
    String::from_utf16(text(edit)?.get(start..end)?).ok()
}

pub(crate) fn correct(
    edit: HWND,
    find: &str,
    replacement: &str,
    context: &str,
    offset: usize,
) -> Option<bool> {
    let original = text(edit)?;
    let doc = String::from_utf16(&original).ok()?;
    eprintln!("Classic edit correction: document units={} context_match={} first_word_match={}", original.len(),doc.starts_with(context),doc.starts_with(find));
    let sentence = doc.char_indices().nth(offset)?.0;
    let byte = if !context.is_empty() && doc[sentence..].starts_with(context) {
        let mut matches = context.match_indices(find);
        let relative = matches.next()?.0;
        if matches.next().is_some() {
            return Some(false);
        }
        sentence + relative
    } else {
        sentence
    };
    if find.is_empty() || !doc[byte..].starts_with(find) || replacement.contains('\0') {
        return Some(false);
    }
    let start = doc[..byte].encode_utf16().count();
    let end = start + find.encode_utf16().count();
    message(edit, 0x00B1, start, end as isize)?; // EM_SETSEL
    let actual=selection(edit)?;
    eprintln!("Classic edit expected selection {}..{}, got {:?}",start,end,actual);
    if actual != (start, end) {
        return Some(false);
    }
    let wide: Vec<_> = replacement.encode_utf16().chain(Some(0)).collect();
    message(edit, 0x00C2, 1, wide.as_ptr() as isize)?; // EM_REPLACESEL, undoable
    let mut expected = original;
    expected.splice(start..end, wide[..wide.len() - 1].iter().copied());
    Some(text(edit)? == expected)
}
