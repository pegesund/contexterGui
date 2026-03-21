// OCR helper library — exposes C-callable function for text recognition.
// Compiled as dylib, loaded from Rust via libloading.
import Vision
import AppKit
import Foundation

/// Recognize text in an image file. Returns a C string (caller must free with ocr_free).
/// Returns NULL on error.
@_cdecl("ocr_recognize_file")
public func ocr_recognize_file(path: UnsafePointer<CChar>) -> UnsafeMutablePointer<CChar>? {
    let pathStr = String(cString: path)

    guard let data = FileManager.default.contents(atPath: pathStr),
          !data.isEmpty,
          let image = NSImage(data: data),
          let cgImage = image.cgImage(forProposedRect: nil, context: nil, hints: nil) else {
        return nil
    }

    let request = VNRecognizeTextRequest()
    request.recognitionLevel = .accurate
    request.recognitionLanguages = ["nb-NO", "nn-NO", "en-US"]
    request.usesLanguageCorrection = true

    let handler = VNImageRequestHandler(cgImage: cgImage, options: [:])
    do {
        try handler.perform([request])
    } catch {
        return nil
    }

    guard let observations = request.results else { return nil }

    var lines: [String] = []
    for obs in observations {
        if let text = obs.topCandidates(1).first?.string {
            lines.append(text)
        }
    }

    let result = lines.joined(separator: "\n")
    if result.isEmpty { return nil }

    // Return a C string that the caller must free with ocr_free
    return strdup(result)
}

/// Free a string returned by ocr_recognize_file.
@_cdecl("ocr_free")
public func ocr_free(ptr: UnsafeMutablePointer<CChar>?) {
    if let ptr = ptr {
        free(ptr)
    }
}
