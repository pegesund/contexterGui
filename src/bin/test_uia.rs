//! Small console tool to test UIA text reading from Edge textarea.
//! Usage: leave cursor in the textarea, then run this.

use windows::Win32::Foundation::*;
use windows::Win32::System::Com::*;
use windows::Win32::UI::Accessibility::*;
use windows::Win32::UI::WindowsAndMessaging::*;

fn main() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);

        let uia: IUIAutomation =
            CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER).unwrap();

        // 1. Foreground window info
        let fg = GetForegroundWindow();
        let mut buf = [0u16; 256];
        let len = GetWindowTextW(fg, &mut buf);
        let title = String::from_utf16_lossy(&buf[..len as usize]);
        let mut fg_pid = 0u32;
        GetWindowThreadProcessId(fg, Some(&mut fg_pid));
        println!("Foreground: '{}' HWND={:?} PID={}", title, fg, fg_pid);

        // 2. Focused element
        println!("\n--- GetFocusedElement ---");
        match uia.GetFocusedElement() {
            Ok(focused) => {
                dump_element(&focused, "Focused");
                try_all_patterns(&focused);
            }
            Err(e) => println!("GetFocusedElement failed: {:?}", e),
        }

        // 3. ElementFromHandle on foreground window
        println!("\n--- ElementFromHandle(fg) ---");
        match uia.ElementFromHandle(fg) {
            Ok(fg_el) => {
                dump_element(&fg_el, "FG root");

                // Search descendants
                println!("\n--- Searching descendants for Edit/Document elements ---");
                if let Ok(cond) = uia.CreateTrueCondition() {
                    if let Ok(all) = fg_el.FindAll(TreeScope_Descendants, &cond) {
                        let count = all.Length().unwrap_or(0);
                        println!("Total descendants: {}", count);
                        for i in 0..count.min(500) {
                            if let Ok(el) = all.GetElement(i) {
                                let ct = el.CurrentControlType().unwrap_or_default();
                                let name = el.CurrentName().unwrap_or_default().to_string();
                                let cls = el.CurrentClassName().unwrap_or_default().to_string();

                                // Show Edit, Document, and any element with ValuePattern
                                let is_edit = ct == UIA_EditControlTypeId;
                                let is_doc = ct == UIA_DocumentControlTypeId;
                                let has_value = el.GetCurrentPatternAs::<IUIAutomationValuePattern>(
                                    UIA_ValuePatternId).is_ok();
                                let has_tp2 = el.GetCurrentPatternAs::<IUIAutomationTextPattern2>(
                                    UIA_TextPattern2Id).is_ok();
                                let has_tp1 = el.GetCurrentPatternAs::<IUIAutomationTextPattern>(
                                    UIA_TextPatternId).is_ok();

                                if is_edit || is_doc || has_value || has_tp2 {
                                    println!("\n  [{}] name='{}' class='{}' type={} edit={} doc={} val={} tp1={} tp2={}",
                                        i, &name[..name.len().min(50)], &cls[..cls.len().min(30)],
                                        ct.0, is_edit, is_doc, has_value, has_tp1, has_tp2);
                                    try_all_patterns(&el);
                                }
                            }
                        }
                    }
                }
            }
            Err(e) => println!("ElementFromHandle failed: {:?}", e),
        }
    }
}

unsafe fn dump_element(el: &IUIAutomationElement, label: &str) {
    let name = el.CurrentName().unwrap_or_default().to_string();
    let cls = el.CurrentClassName().unwrap_or_default().to_string();
    let ct = el.CurrentControlType().unwrap_or_default();
    let pid = el.CurrentProcessId().unwrap_or(0);
    let hwnd = el.CurrentNativeWindowHandle().unwrap_or_default();
    println!("  {} name='{}' class='{}' type={} pid={} hwnd={:?}",
        label, &name[..name.len().min(60)], &cls[..cls.len().min(30)], ct.0, pid, hwnd);
}

unsafe fn try_all_patterns(el: &IUIAutomationElement) {
    // ValuePattern
    if let Ok(vp) = el.GetCurrentPatternAs::<IUIAutomationValuePattern>(UIA_ValuePatternId) {
        match vp.CurrentValue() {
            Ok(v) => {
                let s = v.to_string();
                println!("    ValuePattern: '{}' ({} chars)", &s[..s.len().min(80)], s.len());
            }
            Err(e) => println!("    ValuePattern.CurrentValue failed: {:?}", e),
        }
    }

    // TextPattern2
    if let Ok(tp2) = el.GetCurrentPatternAs::<IUIAutomationTextPattern2>(UIA_TextPattern2Id) {
        match tp2.DocumentRange() {
            Ok(dr) => {
                let text = dr.GetText(-1).unwrap_or_default().to_string();
                println!("    TextPattern2.DocumentRange: '{}' ({} chars)", &text[..text.len().min(80)], text.len());
            }
            Err(e) => println!("    TextPattern2.DocumentRange failed: {:?}", e),
        }
        let mut is_active = windows::core::BOOL::default();
        match tp2.GetCaretRange(&mut is_active) {
            Ok(cr) => {
                let before = cr.Clone().and_then(|r| {
                    let _ = r.MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, -200);
                    r.GetText(-1)
                }).unwrap_or_default().to_string();
                println!("    TextPattern2.CaretRange before: '{}' ({} chars) active={}",
                    &before[..before.len().min(80)], before.len(), is_active.as_bool());
            }
            Err(e) => println!("    TextPattern2.GetCaretRange failed: {:?}", e),
        }
    }

    // TextPattern (v1)
    if let Ok(tp1) = el.GetCurrentPatternAs::<IUIAutomationTextPattern>(UIA_TextPatternId) {
        match tp1.DocumentRange() {
            Ok(dr) => {
                let text = dr.GetText(-1).unwrap_or_default().to_string();
                println!("    TextPattern.DocumentRange: '{}' ({} chars)", &text[..text.len().min(80)], text.len());
            }
            Err(e) => println!("    TextPattern.DocumentRange failed: {:?}", e),
        }
    }
}
