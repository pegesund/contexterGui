// OCR helper — reads image from stdin (PNG), outputs recognized text to stdout.
// Compiled once, called from Rust as a subprocess.
import Vision
import AppKit
import Foundation

// Read PNG data from stdin or from file path argument
let imageData: Data
if CommandLine.arguments.count > 1 {
    let path = CommandLine.arguments[1]
    guard let data = FileManager.default.contents(atPath: path) else {
        fputs("Cannot read file: \(path)\n", stderr)
        exit(1)
    }
    imageData = data
} else {
    imageData = FileHandle.standardInput.readDataToEndOfFile()
}

guard !imageData.isEmpty,
      let image = NSImage(data: imageData),
      let cgImage = image.cgImage(forProposedRect: nil, context: nil, hints: nil) else {
    fputs("Invalid image data\n", stderr)
    exit(1)
}

let request = VNRecognizeTextRequest()
request.recognitionLevel = .accurate
request.recognitionLanguages = ["nb-NO", "nn-NO", "en-US"]
request.usesLanguageCorrection = true

let handler = VNImageRequestHandler(cgImage: cgImage, options: [:])
do {
    try handler.perform([request])
} catch {
    fputs("OCR failed: \(error)\n", stderr)
    exit(1)
}

guard let observations = request.results else { exit(0) }

for obs in observations {
    if let text = obs.topCandidates(1).first?.string {
        print(text)
    }
}
