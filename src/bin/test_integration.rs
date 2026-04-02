//! Integration test for Windows Word COM.
//! Requires: Word open with a document, acatts-rust running.
//!
//! Tests typing text, waiting for error detection, and verifying results
//! via the HTTP /errors endpoint on port 52580.
//!
//! Usage: cargo run --release --bin test_integration

use std::thread;
use std::time::{Duration, Instant};

fn main() {
    println!("=== NorskTale Windows Integration Test ===\n");

    // Wait for the app's HTTP endpoint to be ready
    println!("Waiting for app HTTP endpoint...");
    let start = Instant::now();
    loop {
        if let Ok(json) = fetch_errors() {
            println!("  Connected! Current errors: {}", json.len());
            break;
        }
        if start.elapsed() > Duration::from_secs(10) {
            eprintln!("FATAL: Could not connect to http://127.0.0.1:52580/errors");
            eprintln!("       Make sure acatts-rust is running.");
            std::process::exit(1);
        }
        thread::sleep(Duration::from_millis(500));
    }

    // Connect to Word
    println!("Connecting to Word COM...");
    let word = match WordCom::connect() {
        Some(w) => w,
        None => {
            eprintln!("FATAL: Could not connect to Word. Is Word open?");
            std::process::exit(1);
        }
    };

    // Save original document text for restoration
    let orig_text = word.read_full_text();
    println!("  Document: {} chars\n", orig_text.len());

    let mut pass = 0u32;
    let mut fail = 0u32;

    // --- Test 1: Spelling error detected ---
    {
        println!("Test 1: Spelling error detection");
        word.go_to_end();
        word.type_text("\nJeg liker fiskk.");
        wait_for_errors(1, Duration::from_secs(8));
        let errors = fetch_errors().unwrap_or_default();
        if errors.iter().any(|e| e.category == "spelling" && e.word.contains("fiskk")) {
            println!("  PASS: 'fiskk' detected as spelling error");
            pass += 1;
        } else {
            println!("  FAIL: 'fiskk' not detected. Errors: {:?}", errors);
            fail += 1;
        }
        word.undo(18); // undo typed text
        thread::sleep(Duration::from_millis(500));
    }

    // --- Test 2: Grammar error detected ---
    {
        println!("Test 2: Grammar error detection");
        word.go_to_end();
        word.type_text("\nJeg har gått på kino.");
        // This should be clean — no error
        thread::sleep(Duration::from_secs(5));
        let errors = fetch_errors().unwrap_or_default();
        let has_gatt_error = errors.iter().any(|e| e.sentence.contains("gått"));
        if !has_gatt_error {
            println!("  PASS: 'Jeg har gått på kino' — no false positive");
            pass += 1;
        } else {
            println!("  FAIL: false positive on 'Jeg har gått på kino'. Errors: {:?}", errors);
            fail += 1;
        }
        word.undo(23);
        thread::sleep(Duration::from_millis(500));
    }

    // --- Test 3: Grammar error — har + presens ---
    {
        println!("Test 3: Grammar error 'har + presens'");
        word.go_to_end();
        word.type_text("\nJeg har liker fotball.");
        wait_for_errors(1, Duration::from_secs(8));
        let errors = fetch_errors().unwrap_or_default();
        if errors.iter().any(|e| e.category == "grammar" && e.sentence.contains("liker")) {
            println!("  PASS: 'har liker' grammar error detected");
            pass += 1;
        } else {
            println!("  FAIL: 'har liker' not detected. Errors: {:?}", errors);
            fail += 1;
        }
        word.undo(23);
        thread::sleep(Duration::from_millis(500));
    }

    // --- Test 4: No error on correct text ---
    {
        println!("Test 4: No error on correct text");
        word.go_to_end();
        word.type_text("\nDette er en korrekt setning.");
        thread::sleep(Duration::from_secs(5));
        let errors = fetch_errors().unwrap_or_default();
        let has_our_error = errors.iter().any(|e| e.sentence.contains("korrekt setning"));
        if !has_our_error {
            println!("  PASS: correct sentence — no errors");
            pass += 1;
        } else {
            println!("  FAIL: false positive on correct text. Errors: {:?}", errors);
            fail += 1;
        }
        word.undo(30);
        thread::sleep(Duration::from_millis(500));
    }

    // --- Test 5: Error clears after fix ---
    {
        println!("Test 5: Error clears after fix");
        word.go_to_end();
        word.type_text("\nJeg liker fiskk.");
        wait_for_errors(1, Duration::from_secs(8));
        let errors_before = fetch_errors().unwrap_or_default();
        let had_error = errors_before.iter().any(|e| e.word.contains("fiskk"));

        // Fix: select "fiskk" and replace with "fisk"
        // Move back to "fiskk", select it, type correct word
        for _ in 0..2 { word.key_left(); } // before period and space
        word.select_word_left(); // select "fiskk"
        word.type_text("fisk");
        thread::sleep(Duration::from_secs(5));
        let errors_after = fetch_errors().unwrap_or_default();
        let still_has_error = errors_after.iter().any(|e| e.word.contains("fiskk"));

        if had_error && !still_has_error {
            println!("  PASS: error cleared after fix");
            pass += 1;
        } else {
            println!("  FAIL: had_error={}, still_has={}", had_error, still_has_error);
            fail += 1;
        }
        word.undo(25);
        thread::sleep(Duration::from_millis(500));
    }

    // --- Test 6: Paragraph add/delete ---
    {
        println!("Test 6: Paragraph with error, delete paragraph, error clears");
        word.go_to_end();
        word.type_text("\nDette er feilx ord her.");
        wait_for_errors(1, Duration::from_secs(8));
        let errors_before = fetch_errors().unwrap_or_default();
        let had_error = errors_before.iter().any(|e| e.word.contains("feilx"));

        // Select the whole line and delete
        word.key_home();
        word.select_to_end();
        word.type_text(""); // delete selection
        word.key_backspace(); // delete the newline
        thread::sleep(Duration::from_secs(5));
        let errors_after = fetch_errors().unwrap_or_default();
        let still_has_error = errors_after.iter().any(|e| e.word.contains("feilx"));

        if had_error && !still_has_error {
            println!("  PASS: paragraph deleted — error cleared");
            pass += 1;
        } else {
            println!("  FAIL: had={}, still={}", had_error, still_has_error);
            fail += 1;
        }
        word.undo(30);
        thread::sleep(Duration::from_millis(500));
    }

    println!("\n=== Results: {} passed, {} failed ===", pass, fail);
    if fail > 0 {
        std::process::exit(1);
    }
}

// --- Error types ---

#[derive(Debug)]
struct ErrorInfo {
    category: String,
    word: String,
    suggestion: String,
    rule: String,
    sentence: String,
}

fn fetch_errors() -> Result<Vec<ErrorInfo>, String> {
    let resp = std::net::TcpStream::connect_timeout(
        &"127.0.0.1:52580".parse().unwrap(),
        Duration::from_secs(2),
    ).map_err(|e| format!("connect: {}", e))?;

    use std::io::{Read, Write};
    let mut stream = resp;
    stream.write_all(b"GET /errors HTTP/1.0\r\nHost: localhost\r\n\r\n")
        .map_err(|e| format!("write: {}", e))?;
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf).map_err(|e| format!("read: {}", e))?;
    let text = String::from_utf8_lossy(&buf);

    // Skip HTTP headers
    let body = text.split("\r\n\r\n").nth(1).unwrap_or("[]");
    parse_errors_json(body)
}

fn parse_errors_json(json: &str) -> Result<Vec<ErrorInfo>, String> {
    // Simple JSON array parser for our known format
    let mut errors = Vec::new();
    if json.trim() == "[]" { return Ok(errors); }

    // Split by },{ to get individual error objects
    let inner = json.trim().trim_start_matches('[').trim_end_matches(']');
    for obj in inner.split("},{") {
        let obj = obj.trim().trim_start_matches('{').trim_end_matches('}');
        let mut category = String::new();
        let mut word = String::new();
        let mut suggestion = String::new();
        let mut rule = String::new();
        let mut sentence = String::new();

        for field in obj.split("\",\"") {
            let field = field.trim_matches(|c| c == '{' || c == '}' || c == '"');
            if let Some((key, val)) = field.split_once("\":\"") {
                let key = key.trim_matches('"');
                let val = val.trim_matches('"');
                match key {
                    "category" => category = val.to_string(),
                    "word" => word = val.to_string(),
                    "suggestion" => suggestion = val.to_string(),
                    "rule" => rule = val.to_string(),
                    "sentence" => sentence = val.to_string(),
                    _ => {}
                }
            }
        }
        if !category.is_empty() {
            errors.push(ErrorInfo { category, word, suggestion, rule, sentence });
        }
    }
    Ok(errors)
}

fn wait_for_errors(min_count: usize, timeout: Duration) {
    let start = Instant::now();
    while start.elapsed() < timeout {
        if let Ok(errors) = fetch_errors() {
            if errors.len() >= min_count {
                return;
            }
        }
        thread::sleep(Duration::from_millis(300));
    }
}

// --- Word COM wrapper ---

struct WordCom {
    // We use Python for COM calls since Rust COM late binding is complex
    // and we already have it working in Python
}

impl WordCom {
    fn connect() -> Option<Self> {
        // Verify Word is running
        let output = std::process::Command::new("py")
            .args(["-c", "import win32com.client; w = win32com.client.Dispatch('Word.Application'); print(w.Documents.Count)"])
            .output().ok()?;
        let count = String::from_utf8_lossy(&output.stdout).trim().to_string();
        if count.parse::<i32>().unwrap_or(0) > 0 {
            Some(WordCom {})
        } else {
            None
        }
    }

    fn py(&self, script: &str) {
        let full = format!(
            "import win32com.client\nw = win32com.client.Dispatch('Word.Application')\nsel = w.Selection\ndoc = w.ActiveDocument\n{}",
            script
        );
        let _ = std::process::Command::new("py").args(["-c", &full]).output();
    }

    fn read_full_text(&self) -> String {
        let output = std::process::Command::new("py")
            .args(["-c", "import win32com.client; w = win32com.client.Dispatch('Word.Application'); print(w.ActiveDocument.Content.Text)"])
            .output().ok();
        output.map(|o| String::from_utf8_lossy(&o.stdout).trim().to_string()).unwrap_or_default()
    }

    fn type_text(&self, text: &str) {
        // Type character by character with small delay for realism
        for ch in text.chars() {
            if ch == '\n' {
                self.py("sel.TypeParagraph()");
            } else {
                let escaped = ch.to_string().replace('\\', "\\\\").replace('\'', "\\'");
                self.py(&format!("sel.TypeText('{}')", escaped));
            }
            thread::sleep(Duration::from_millis(50));
        }
    }

    fn go_to_end(&self) {
        self.py("sel.EndKey(Unit=6)"); // wdStory = 6
    }

    fn key_left(&self) {
        self.py("sel.MoveLeft(Unit=1, Count=1)");
    }

    fn key_home(&self) {
        self.py("sel.HomeKey(Unit=5)"); // wdLine = 5
    }

    fn key_backspace(&self) {
        self.py("sel.TypeBackspace()");
    }

    fn select_word_left(&self) {
        self.py("sel.MoveLeft(Unit=2, Count=1, Extend=1)"); // wdWord = 2, wdExtend = 1
    }

    fn select_to_end(&self) {
        self.py("sel.EndKey(Unit=5, Extend=1)"); // wdLine, wdExtend
    }

    fn undo(&self, count: u32) {
        self.py(&format!("doc.Undo({})", count));
    }
}
