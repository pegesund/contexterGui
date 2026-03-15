//! Clipboard-based OCR.
//! Windows: uses Windows.Media.Ocr. Other platforms: no-op stub.

use std::sync::mpsc;

// ── Windows implementation ──

#[cfg(target_os = "windows")]
mod windows_impl {
    use windows::core::HSTRING;
    use windows::Globalization::Language;
    use windows::Graphics::Imaging::{BitmapDecoder, SoftwareBitmap, BitmapPixelFormat};
    use windows::Media::Ocr::OcrEngine;
    use windows::Storage::Streams::{InMemoryRandomAccessStream, DataWriter};
    use windows::Win32::Foundation::HGLOBAL;
    use windows::Win32::System::DataExchange::*;
    use windows::Win32::System::Memory::{GlobalLock, GlobalSize, GlobalUnlock};
    use windows::Win32::System::Ole::CF_DIB;
    use std::sync::mpsc;

    pub struct OcrClipboard {
        last_seq: u32,
        pending_image: bool,
        dismissed_seq: u32,
    }

    impl OcrClipboard {
        pub fn new() -> Result<Self, String> {
            let lang = Language::CreateLanguage(&HSTRING::from("nb"))
                .map_err(|e| format!("Failed to create Language('nb'): {}", e))?;
            if !OcrEngine::IsLanguageSupported(&lang)
                .map_err(|e| format!("IsLanguageSupported: {}", e))? {
                return Err("Norwegian OCR language pack not installed.".into());
            }
            let _engine = OcrEngine::TryCreateFromLanguage(&lang)
                .map_err(|e| format!("TryCreateFromLanguage: {}", e))?;
            let seq = unsafe { GetClipboardSequenceNumber() };
            Ok(Self { last_seq: seq, pending_image: false, dismissed_seq: 0 })
        }
        pub fn poll(&mut self) {
            let seq = unsafe { GetClipboardSequenceNumber() };
            if seq != self.last_seq {
                self.last_seq = seq;
                if seq == self.dismissed_seq { return; }
                if clipboard_has_image() { self.pending_image = true; }
            }
        }
        pub fn has_pending_image(&self) -> bool { self.pending_image }
        pub fn dismiss(&mut self) { self.pending_image = false; self.dismissed_seq = self.last_seq; }
        pub fn start_ocr(&mut self) -> Option<mpsc::Receiver<Result<String, String>>> {
            self.pending_image = false;
            let dib_data = read_clipboard_dib()?;
            let (tx, rx) = mpsc::channel();
            std::thread::spawn(move || {
                unsafe {
                    let _ = windows::Win32::System::Com::CoInitializeEx(
                        None, windows::Win32::System::Com::COINIT_MULTITHREADED);
                }
                let _ = tx.send(run_ocr_on_dib(&dib_data));
            });
            Some(rx)
        }
    }

    fn clipboard_has_image() -> bool {
        unsafe {
            if OpenClipboard(None).is_ok() {
                let has = IsClipboardFormatAvailable(CF_DIB.0.into()).is_ok();
                let _ = CloseClipboard();
                has
            } else { false }
        }
    }

    fn read_clipboard_dib() -> Option<Vec<u8>> {
        unsafe {
            if OpenClipboard(None).is_err() { return None; }
            let handle = GetClipboardData(CF_DIB.0.into());
            if handle.is_err() { let _ = CloseClipboard(); return None; }
            let hglobal = HGLOBAL(handle.unwrap().0 as *mut _);
            let size = GlobalSize(hglobal);
            if size == 0 { let _ = CloseClipboard(); return None; }
            let ptr = GlobalLock(hglobal);
            if ptr.is_null() { let _ = CloseClipboard(); return None; }
            let data = std::slice::from_raw_parts(ptr as *const u8, size).to_vec();
            let _ = GlobalUnlock(hglobal);
            let _ = CloseClipboard();
            Some(data)
        }
    }

    fn run_ocr_on_dib(dib_data: &[u8]) -> Result<String, String> {
        let file_header_size: u32 = 14;
        let file_size = file_header_size + dib_data.len() as u32;
        let header_size = u32::from_le_bytes([dib_data[0], dib_data[1], dib_data[2], dib_data[3]]);
        let bit_count = u16::from_le_bytes([dib_data[14], dib_data[15]]);
        let color_table_size = if bit_count <= 8 {
            let cu = u32::from_le_bytes([dib_data[32], dib_data[33], dib_data[34], dib_data[35]]);
            (if cu == 0 { 1u32 << bit_count } else { cu }) * 4
        } else { 0 };
        let pixel_offset = file_header_size + header_size + color_table_size;
        let mut bmp = Vec::with_capacity(file_size as usize);
        bmp.extend_from_slice(b"BM");
        bmp.extend_from_slice(&file_size.to_le_bytes());
        bmp.extend_from_slice(&0u16.to_le_bytes());
        bmp.extend_from_slice(&0u16.to_le_bytes());
        bmp.extend_from_slice(&pixel_offset.to_le_bytes());
        bmp.extend_from_slice(dib_data);

        let stream = InMemoryRandomAccessStream::new().map_err(|e| format!("{}", e))?;
        let writer = DataWriter::CreateDataWriter(&stream).map_err(|e| format!("{}", e))?;
        writer.WriteBytes(&bmp).map_err(|e| format!("{}", e))?;
        writer.StoreAsync().map_err(|e| format!("{}", e))?.get().map_err(|e| format!("{}", e))?;
        writer.FlushAsync().map_err(|e| format!("{}", e))?.get().map_err(|e| format!("{}", e))?;
        stream.Seek(0).map_err(|e| format!("{}", e))?;

        let decoder = BitmapDecoder::CreateAsync(&stream).map_err(|e| format!("{}", e))?
            .get().map_err(|e| format!("{}", e))?;
        let bitmap = decoder.GetSoftwareBitmapAsync().map_err(|e| format!("{}", e))?
            .get().map_err(|e| format!("{}", e))?;
        let pf = bitmap.BitmapPixelFormat().map_err(|e| format!("{}", e))?;
        let ocr_bitmap = if pf != BitmapPixelFormat::Gray8 && pf != BitmapPixelFormat::Bgra8 {
            SoftwareBitmap::Convert(&bitmap, BitmapPixelFormat::Bgra8).map_err(|e| format!("{}", e))?
        } else { bitmap };

        let lang = Language::CreateLanguage(&HSTRING::from("nb")).map_err(|e| format!("{}", e))?;
        let engine = OcrEngine::TryCreateFromLanguage(&lang).map_err(|e| format!("{}", e))?;
        let result = engine.RecognizeAsync(&ocr_bitmap).map_err(|e| format!("{}", e))?
            .get().map_err(|e| format!("{}", e))?;
        let lines = result.Lines().map_err(|e| format!("{}", e))?;
        let mut text = String::new();
        for line in &lines {
            let lt = line.Text().map_err(|e| format!("{}", e))?;
            if !text.is_empty() { text.push('\n'); }
            text.push_str(&lt.to_string());
        }
        Ok(text)
    }
}

// ── Cross-platform public API ──

#[cfg(target_os = "windows")]
pub use windows_impl::OcrClipboard;

#[cfg(not(target_os = "windows"))]
pub struct OcrClipboard;

#[cfg(not(target_os = "windows"))]
impl OcrClipboard {
    pub fn new() -> Result<Self, String> { Err("OCR not available on this platform".into()) }
    pub fn poll(&mut self) {}
    pub fn has_pending_image(&self) -> bool { false }
    pub fn dismiss(&mut self) {}
    pub fn start_ocr(&mut self) -> Option<mpsc::Receiver<Result<String, String>>> { None }
}
