//! Clipboard-based OCR using Windows.Media.Ocr.
//!
//! Monitors the clipboard for new images (from Win+Shift+S, PrtScn, etc.)
//! and runs Norwegian OCR via the built-in Windows OCR engine.

use windows::core::HSTRING;
use windows::Globalization::Language;
use windows::Graphics::Imaging::{BitmapDecoder, SoftwareBitmap, BitmapPixelFormat};
use windows::Media::Ocr::OcrEngine;
use windows::Storage::Streams::{InMemoryRandomAccessStream, DataWriter};
use windows::Win32::Foundation::HGLOBAL;
use windows::Win32::System::DataExchange::{
    CloseClipboard, GetClipboardData, GetClipboardSequenceNumber, OpenClipboard,
};
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
            .map_err(|e| format!("IsLanguageSupported: {}", e))?
        {
            return Err("Norwegian OCR language pack not installed. Install via Settings > Time & Language > Language > Norsk Bokmål > Options.".into());
        }

        // Verify the engine can be created (validates language pack)
        let _engine = OcrEngine::TryCreateFromLanguage(&lang)
            .map_err(|e| format!("TryCreateFromLanguage: {}", e))?;

        let seq = unsafe { GetClipboardSequenceNumber() };

        Ok(Self {
            last_seq: seq,
            pending_image: false,
            dismissed_seq: 0,
        })
    }

    /// Check if the clipboard has changed and contains an image.
    pub fn poll(&mut self) {
        let seq = unsafe { GetClipboardSequenceNumber() };
        if seq != self.last_seq {
            self.last_seq = seq;
            // Don't re-prompt for a dismissed image
            if seq == self.dismissed_seq {
                return;
            }
            // Check if clipboard has a bitmap
            if clipboard_has_image() {
                self.pending_image = true;
            }
        }
    }

    pub fn has_pending_image(&self) -> bool {
        self.pending_image
    }

    pub fn dismiss(&mut self) {
        self.pending_image = false;
        self.dismissed_seq = self.last_seq;
    }

    /// Start OCR on a background thread. Returns a receiver for the result.
    pub fn start_ocr(&mut self) -> Option<mpsc::Receiver<Result<String, String>>> {
        self.pending_image = false;

        // Read the DIB from clipboard on the main thread (clipboard is STA-bound)
        let dib_data = match read_clipboard_dib() {
            Some(data) => data,
            None => {
                eprintln!("OCR: failed to read clipboard image");
                return None;
            }
        };

        let (tx, rx) = mpsc::channel();

        std::thread::spawn(move || {
            // Initialize COM for this thread (MTA)
            unsafe {
                let _ = windows::Win32::System::Com::CoInitializeEx(
                    None,
                    windows::Win32::System::Com::COINIT_MULTITHREADED,
                );
            }

            let result = run_ocr_on_dib(&dib_data);
            let _ = tx.send(result);
        });

        Some(rx)
    }
}

fn clipboard_has_image() -> bool {
    unsafe {
        if OpenClipboard(None).is_ok() {
            let has = windows::Win32::System::DataExchange::IsClipboardFormatAvailable(
                CF_DIB.0.into(),
            )
            .is_ok();
            let _ = CloseClipboard();
            has
        } else {
            false
        }
    }
}

fn read_clipboard_dib() -> Option<Vec<u8>> {
    unsafe {
        if OpenClipboard(None).is_err() {
            return None;
        }

        let handle = GetClipboardData(CF_DIB.0.into());
        if handle.is_err() {
            let _ = CloseClipboard();
            return None;
        }
        let hglobal = HGLOBAL(handle.unwrap().0 as *mut _);

        let size = GlobalSize(hglobal);
        if size == 0 {
            let _ = CloseClipboard();
            return None;
        }

        let ptr = GlobalLock(hglobal);
        if ptr.is_null() {
            let _ = CloseClipboard();
            return None;
        }

        let data = std::slice::from_raw_parts(ptr as *const u8, size).to_vec();
        let _ = GlobalUnlock(hglobal);
        let _ = CloseClipboard();

        Some(data)
    }
}

fn run_ocr_on_dib(dib_data: &[u8]) -> Result<String, String> {
    // Build a BMP file from the DIB data (add 14-byte BMP file header)
    let file_header_size: u32 = 14;
    let file_size = file_header_size + dib_data.len() as u32;

    // Read the DIB header to get the offset to pixel data
    let header_size = u32::from_le_bytes([dib_data[0], dib_data[1], dib_data[2], dib_data[3]]);
    let bit_count = u16::from_le_bytes([dib_data[14], dib_data[15]]);

    // Color table size (for ≤8bpp images)
    let color_table_size = if bit_count <= 8 {
        let colors_used =
            u32::from_le_bytes([dib_data[32], dib_data[33], dib_data[34], dib_data[35]]);
        let num_colors = if colors_used == 0 {
            1u32 << bit_count
        } else {
            colors_used
        };
        num_colors * 4
    } else {
        0
    };

    let pixel_offset = file_header_size + header_size + color_table_size;

    let mut bmp = Vec::with_capacity(file_size as usize);
    // BMP file header (14 bytes)
    bmp.extend_from_slice(b"BM");
    bmp.extend_from_slice(&file_size.to_le_bytes());
    bmp.extend_from_slice(&0u16.to_le_bytes()); // reserved1
    bmp.extend_from_slice(&0u16.to_le_bytes()); // reserved2
    bmp.extend_from_slice(&pixel_offset.to_le_bytes());
    // DIB data
    bmp.extend_from_slice(dib_data);

    // Write BMP to an in-memory stream
    let stream = InMemoryRandomAccessStream::new()
        .map_err(|e| format!("InMemoryRandomAccessStream: {}", e))?;
    let writer = DataWriter::CreateDataWriter(&stream)
        .map_err(|e| format!("DataWriter: {}", e))?;
    writer
        .WriteBytes(&bmp)
        .map_err(|e| format!("WriteBytes: {}", e))?;
    writer
        .StoreAsync()
        .map_err(|e| format!("StoreAsync: {}", e))?
        .get()
        .map_err(|e| format!("StoreAsync.get: {}", e))?;
    writer
        .FlushAsync()
        .map_err(|e| format!("FlushAsync: {}", e))?
        .get()
        .map_err(|e| format!("FlushAsync.get: {}", e))?;

    // Seek stream to beginning
    stream
        .Seek(0)
        .map_err(|e| format!("Seek: {}", e))?;

    // Decode BMP to SoftwareBitmap
    let decoder = BitmapDecoder::CreateAsync(&stream)
        .map_err(|e| format!("BitmapDecoder::CreateAsync: {}", e))?
        .get()
        .map_err(|e| format!("BitmapDecoder.get: {}", e))?;

    let bitmap = decoder
        .GetSoftwareBitmapAsync()
        .map_err(|e| format!("GetSoftwareBitmapAsync: {}", e))?
        .get()
        .map_err(|e| format!("GetSoftwareBitmapAsync.get: {}", e))?;

    // OcrEngine requires Gray8 or Bgra8 — convert if needed
    let pixel_format = bitmap
        .BitmapPixelFormat()
        .map_err(|e| format!("BitmapPixelFormat: {}", e))?;

    let ocr_bitmap = if pixel_format != BitmapPixelFormat::Gray8
        && pixel_format != BitmapPixelFormat::Bgra8
    {
        SoftwareBitmap::Convert(&bitmap, BitmapPixelFormat::Bgra8)
            .map_err(|e| format!("SoftwareBitmap::Convert: {}", e))?
    } else {
        bitmap
    };

    // Create OCR engine (on this MTA thread)
    let lang = Language::CreateLanguage(&HSTRING::from("nb"))
        .map_err(|e| format!("Language: {}", e))?;
    let engine = OcrEngine::TryCreateFromLanguage(&lang)
        .map_err(|e| format!("OcrEngine: {}", e))?;

    // Run OCR
    let result = engine
        .RecognizeAsync(&ocr_bitmap)
        .map_err(|e| format!("RecognizeAsync: {}", e))?
        .get()
        .map_err(|e| format!("RecognizeAsync.get: {}", e))?;

    // Extract text from lines
    let lines = result
        .Lines()
        .map_err(|e| format!("Lines: {}", e))?;

    let mut text = String::new();
    for line in &lines {
        let line_text = line
            .Text()
            .map_err(|e| format!("Line.Text: {}", e))?;
        if !text.is_empty() {
            text.push('\n');
        }
        text.push_str(&line_text.to_string());
    }

    eprintln!("OCR result ({} lines, {} chars): {}", lines.Size().unwrap_or(0), text.len(),
        &text[..text.len().min(100)]);

    Ok(text)
}
