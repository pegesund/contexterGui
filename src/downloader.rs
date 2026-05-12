//! S3 model/data downloader with progress reporting.
//!
//! Downloads language data and models from Contabo S3 (eu2.contabostorage.com).
//! Uses presigned URLs so credentials never leave the app binary.

use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};

// ── S3 config (read-only credentials for the spell bucket) ──

const S3_ENDPOINT: &str = "https://eu2.contabostorage.com";
const S3_BUCKET: &str = "spell";
const S3_ACCESS_KEY: &str = "cd59e2c4bbbd7bd29951f126d87a096a";
const S3_SECRET_KEY: &str = "3f28f3941d0d20aaa829ef17c50fe4e7";
const S3_REGION: &str = "eu2";

// ── Public types ──

/// Progress of a single file download.
#[derive(Clone, Debug)]
pub struct DownloadProgress {
    /// Human-readable label (e.g. "Ordbok", not the filename)
    pub label: String,
    /// Bytes downloaded so far
    pub downloaded: u64,
    /// Total size in bytes (0 if unknown)
    pub total: u64,
    /// True when this file is done
    pub done: bool,
    /// Error message if download failed
    pub error: Option<String>,
}

/// Shared progress state polled by the UI.
pub type SharedProgress = Arc<Mutex<Vec<DownloadProgress>>>;

/// A file to download.
pub struct DownloadItem {
    /// S3 object key (e.g. "lang/nb/fullform_bm.mfst")
    pub s3_key: String,
    /// Local destination path
    pub local_path: PathBuf,
    /// Human-readable label for the progress bar
    pub label: String,
}

// ── S3 presigned URL generation (AWS Signature V4) ──

/// AWS Sig v4 URI-path encoding: encode every byte except A-Za-z0-9-_.~ and
/// `/`. Without this, S3 keys containing `!`, `(`, `)`, spaces, etc. produce
/// presigned URLs that S3 rejects with 403 because S3 percent-encodes the
/// path before recomputing the signature on its side. Discovered after the
/// espeak-ng manifest tried to download files like
/// `bin/mac/espeak-ng-data/voices/!v/adam` and failed at the first `!`.
fn aws_uri_encode_path(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for b in s.bytes() {
        if b.is_ascii_alphanumeric()
            || b == b'-' || b == b'_' || b == b'.' || b == b'~' || b == b'/'
        {
            out.push(b as char);
        } else {
            out.push_str(&format!("%{:02X}", b));
        }
    }
    out
}

fn presign_url(key: &str, expires_secs: u64) -> String {
    use hmac::{Hmac, Mac};
    use sha2::Sha256;

    let now = chrono::Utc::now();
    let date_stamp = now.format("%Y%m%d").to_string();
    let amz_date = now.format("%Y%m%dT%H%M%SZ").to_string();

    let host = S3_ENDPOINT.trim_start_matches("https://");
    let canonical_uri = aws_uri_encode_path(&format!("/{}/{}", S3_BUCKET, key));
    let scope = format!("{}/{}/s3/aws4_request", date_stamp, S3_REGION);

    let credential = format!("{}/{}", S3_ACCESS_KEY, scope);
    let credential_encoded = credential.replace('/', "%2F");

    let query = format!(
        "X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Credential={}&X-Amz-Date={}&X-Amz-Expires={}&X-Amz-SignedHeaders=host",
        credential_encoded, amz_date, expires_secs
    );

    let canonical_request = format!(
        "GET\n{}\n{}\nhost:{}\n\nhost\nUNSIGNED-PAYLOAD",
        canonical_uri, query, host
    );

    let hash = {
        use sha2::Digest;
        let mut h = Sha256::new();
        h.update(canonical_request.as_bytes());
        hex::encode(h.finalize())
    };

    let string_to_sign = format!(
        "AWS4-HMAC-SHA256\n{}\n{}\n{}",
        amz_date, scope, hash
    );

    // Derive signing key
    fn hmac_sha256(key: &[u8], data: &[u8]) -> Vec<u8> {
        let mut mac = Hmac::<Sha256>::new_from_slice(key).unwrap();
        mac.update(data);
        mac.finalize().into_bytes().to_vec()
    }

    let k_date = hmac_sha256(format!("AWS4{}", S3_SECRET_KEY).as_bytes(), date_stamp.as_bytes());
    let k_region = hmac_sha256(&k_date, S3_REGION.as_bytes());
    let k_service = hmac_sha256(&k_region, b"s3");
    let k_signing = hmac_sha256(&k_service, b"aws4_request");

    let signature = hex::encode(hmac_sha256(&k_signing, string_to_sign.as_bytes()));

    format!(
        "{}{}?{}&X-Amz-Signature={}",
        S3_ENDPOINT, canonical_uri, query, signature
    )
}

// ── Download logic ──

/// Append a download-failure line to `<data_dir>/download_errors.log`. We need
/// our own logger because the lib crate can't reach the binary's `log!` macro,
/// and stderr from a packaged .app goes to nowhere the user can find.
fn log_download_error(idx: usize, total: usize, key: &str, err: &str) {
    let log_path = data_dir().join("download_errors.log");
    if let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(&log_path)
    {
        let ts = chrono::Utc::now().format("%Y-%m-%dT%H:%M:%SZ");
        let _ = writeln!(f, "{} [{}/{}] {} — {}", ts, idx, total, key, err);
    }
}

/// Classify a ureq::Error into a short tag and a long description.
/// Helps users (and us) tell apart "network is down" from "the server
/// returned 403" from "TLS handshake failed" — without these tags the
/// download_errors.log was just an opaque string.
///
/// Takes the error by value so we can consume the wrapped Response on
/// the Status variant to surface S3's error XML body. The body is the
/// difference between "we know it was 403 but no idea why" and
/// "SignatureDoesNotMatch — clock skew / wrong key / etc."
fn classify_ureq_err(e: ureq::Error) -> (&'static str, String) {
    match e {
        ureq::Error::Status(code, resp) => {
            let url = resp.get_url().to_string();
            // S3 error responses are XML payloads <1KB. Read up to 8KB
            // to be safe and stop. into_string() consumes the response.
            let body = resp.into_string().unwrap_or_default();
            let tag = match code {
                400 => "HTTP_400_BAD_REQUEST",
                401 => "HTTP_401_UNAUTHORIZED",
                403 => "HTTP_403_FORBIDDEN_OR_SIG_INVALID",
                404 => "HTTP_404_KEY_NOT_FOUND",
                408 => "HTTP_408_TIMEOUT",
                429 => "HTTP_429_RATE_LIMITED",
                500..=599 => "HTTP_5XX_SERVER_ERROR",
                _ => "HTTP_OTHER",
            };
            (tag, format!("status={} url={} body_first_500={}",
                code, url, body.chars().take(500).collect::<String>()))
        }
        ureq::Error::Transport(t) => {
            // ureq's Transport variants tell us a lot: DNS, Connect, Tls, Io...
            let kind = format!("{:?}", t.kind());
            let tag = match t.kind() {
                ureq::ErrorKind::Dns => "NET_DNS_FAILED",
                ureq::ErrorKind::ConnectionFailed => "NET_CONNECT_REFUSED",
                ureq::ErrorKind::TooManyRedirects => "NET_TOO_MANY_REDIRECTS",
                ureq::ErrorKind::BadStatus => "NET_BAD_STATUS",
                ureq::ErrorKind::BadHeader => "NET_BAD_HEADER",
                ureq::ErrorKind::Io => "NET_IO_ERROR",
                ureq::ErrorKind::InvalidUrl => "URL_INVALID",
                ureq::ErrorKind::UnknownScheme => "URL_BAD_SCHEME",
                ureq::ErrorKind::ProxyConnect => "NET_PROXY_CONNECT_FAILED",
                ureq::ErrorKind::ProxyUnauthorized => "NET_PROXY_AUTH_REQUIRED",
                _ => "NET_OTHER",
            };
            (tag, format!("transport_kind={} msg={}", kind, t))
        }
    }
}

/// Download a single file from S3 to a local path, reporting progress.
fn download_one(item: &DownloadItem, progress: &SharedProgress, index: usize) -> Result<(), String> {
    use std::io::Read;
    let start = std::time::Instant::now();

    // Create parent directories
    if let Some(parent) = item.local_path.parent() {
        std::fs::create_dir_all(parent)
            .map_err(|e| format!("LOCAL_MKDIR_FAILED parent={} msg={}",
                parent.display(), e))?;
    }

    let url = presign_url(&item.s3_key, 3600);
    // Strip the signature query string before logging URLs — the key
    // is what we want to see in support logs, not the signature.
    let safe_url = url.splitn(2, '?').next().unwrap_or(&url).to_string();

    // Use a builder with explicit timeouts so a stalled connection
    // fails fast with a clear error instead of hanging the entire
    // language-pack flow.
    let agent = ureq::AgentBuilder::new()
        .timeout_connect(std::time::Duration::from_secs(15))
        .timeout_read(std::time::Duration::from_secs(120))
        .build();
    let resp = match agent.get(&url).call() {
        Ok(r) => r,
        Err(e) => {
            let (tag, detail) = classify_ureq_err(e);
            return Err(format!("{} url={} elapsed_ms={} {}",
                tag, safe_url, start.elapsed().as_millis(), detail));
        }
    };

    let status = resp.status();
    let total: u64 = resp.header("Content-Length")
        .and_then(|s| s.parse().ok())
        .unwrap_or(0);
    let server = resp.header("Server").unwrap_or("").to_string();
    let date_hdr = resp.header("Date").unwrap_or("").to_string();

    // Update total in progress
    if let Ok(mut p) = progress.lock() {
        p[index].total = total;
    }

    let mut reader = resp.into_reader();
    let tmp_path = item.local_path.with_extension("download");
    let mut file = std::fs::File::create(&tmp_path)
        .map_err(|e| format!("LOCAL_CREATE_FAILED path={} msg={}",
            tmp_path.display(), e))?;

    let mut buf = [0u8; 65536];
    let mut downloaded: u64 = 0;

    loop {
        let n = match reader.read(&mut buf) {
            Ok(n) => n,
            Err(e) => {
                return Err(format!(
                    "NET_READ_FAILED url={} status={} got={}/{} server={} date={} elapsed_ms={} msg={}",
                    safe_url, status, downloaded, total, server, date_hdr,
                    start.elapsed().as_millis(), e));
            }
        };
        if n == 0 { break; }
        if let Err(e) = file.write_all(&buf[..n]) {
            return Err(format!("LOCAL_WRITE_FAILED path={} got={} msg={}",
                tmp_path.display(), downloaded, e));
        }
        downloaded += n as u64;

        if let Ok(mut p) = progress.lock() {
            p[index].downloaded = downloaded;
        }
    }

    // Atomic rename
    std::fs::rename(&tmp_path, &item.local_path)
        .map_err(|e| format!("LOCAL_RENAME_FAILED from={} to={} msg={}",
            tmp_path.display(), item.local_path.display(), e))?;

    if let Ok(mut p) = progress.lock() {
        p[index].done = true;
    }

    Ok(())
}

/// Check if a local file is up to date by comparing size with S3.
/// Returns true if the file exists and has the expected size.
fn is_cached(item: &DownloadItem) -> bool {
    if !item.local_path.exists() {
        return false;
    }
    // If file exists and is non-empty, consider it cached.
    // A more sophisticated check could HEAD the S3 object and compare
    // Last-Modified, but for now size > 0 is sufficient.
    std::fs::metadata(&item.local_path)
        .map(|m| m.len() > 0)
        .unwrap_or(false)
}

/// Local data directory: ~/Library/Application Support/Spell/data/
pub fn data_dir() -> PathBuf {
    let dir = if cfg!(target_os = "macos") {
        dirs::home_dir()
            .map(|h| h.join("Library/Application Support/Spell/data"))
            .unwrap_or_else(|| PathBuf::from("/tmp/spell/data"))
    } else {
        dirs::config_dir()
            .map(|c| c.join("Spell/data"))
            .unwrap_or_else(|| PathBuf::from("spell/data"))
    };
    let _ = std::fs::create_dir_all(&dir);
    dir
}

// ── Language definitions ──

/// Files needed for a language.
pub fn language_files(lang_code: &str) -> Vec<DownloadItem> {
    let base = data_dir();

    let mut items = Vec::new();

    // BERT model — Norwegian shares NorBERT4, English uses ModernBERT
    let bert_dir = base.join("models/bert");
    match lang_code {
        "en" => {
            items.push(DownloadItem {
                s3_key: "models/bert/modernbert_base_int8.onnx".into(),
                local_path: bert_dir.join("modernbert_base_int8.onnx"),
                label: "Language model".into(),
            });
            items.push(DownloadItem {
                s3_key: "models/bert/tokenizer_en.json".into(),
                local_path: bert_dir.join("tokenizer_en.json"),
                label: "Tokenizer".into(),
            });
        }
        _ => {
            // Norwegian (nb/nn) shares NorBERT4
            items.push(DownloadItem {
                s3_key: "models/bert/norbert4_base_int8.onnx".into(),
                local_path: bert_dir.join("norbert4_base_int8.onnx"),
                label: "Språkmodell".into(),
            });
            items.push(DownloadItem {
                s3_key: "models/bert/tokenizer.json".into(),
                local_path: bert_dir.join("tokenizer.json"),
                label: "Tokenizer".into(),
            });
        }
    }

    // Per-language files
    match lang_code {
        "nb" => {
            let dir = base.join("lang/nb");
            items.push(DownloadItem { s3_key: "lang/nb/fullform_bm.mfst".into(), local_path: dir.join("fullform_bm.mfst"), label: "Ordbok".into() });
            items.push(DownloadItem { s3_key: "lang/nb/wordfreq_bm.tsv".into(), local_path: dir.join("wordfreq_bm.tsv"), label: "Ordfrekvenser".into() });
            items.push(DownloadItem { s3_key: "lang/nb/grammar_rules.pl".into(), local_path: dir.join("grammar_rules.pl"), label: "Grammatikk".into() });
            items.push(DownloadItem { s3_key: "lang/nb/compound_data.pl".into(), local_path: dir.join("compound_data.pl"), label: "Sammensatte ord".into() });
            items.push(DownloadItem { s3_key: "lang/nb/sentence_split.pl".into(), local_path: dir.join("sentence_split.pl"), label: "Setningsdeling".into() });
        }
        "nn" => {
            let dir = base.join("lang/nn");
            items.push(DownloadItem { s3_key: "lang/nn/fullform_nn.mfst".into(), local_path: dir.join("fullform_nn.mfst"), label: "Ordbok".into() });
            items.push(DownloadItem { s3_key: "lang/nn/wordfreq_nn.tsv".into(), local_path: dir.join("wordfreq_nn.tsv"), label: "Ordfrekvens".into() });
            items.push(DownloadItem { s3_key: "lang/nn/grammar_rules.pl".into(), local_path: dir.join("grammar_rules.pl"), label: "Grammatikk".into() });
            items.push(DownloadItem { s3_key: "lang/nn/compound_data.pl".into(), local_path: dir.join("compound_data.pl"), label: "Samansette ord".into() });
            items.push(DownloadItem { s3_key: "lang/nn/sentence_split.pl".into(), local_path: dir.join("sentence_split.pl"), label: "Setningsdeling".into() });
        }
        "en" => {
            let dir = base.join("lang/en");
            items.push(DownloadItem { s3_key: "lang/en/fullform_en.mfst".into(), local_path: dir.join("fullform_en.mfst"), label: "Dictionary".into() });
            items.push(DownloadItem { s3_key: "lang/en/wordfreq_en.tsv".into(), local_path: dir.join("wordfreq_en.tsv"), label: "Word frequencies".into() });
            items.push(DownloadItem { s3_key: "lang/en/grammar_rules.pl".into(), local_path: dir.join("grammar_rules.pl"), label: "Grammar".into() });
            items.push(DownloadItem { s3_key: "lang/en/compound_data.pl".into(), local_path: dir.join("compound_data.pl"), label: "Compound words".into() });
            items.push(DownloadItem { s3_key: "lang/en/sentence_split.pl".into(), local_path: dir.join("sentence_split.pl"), label: "Sentence splitting".into() });
        }
        _ => {}
    }

    items
}

/// Piper TTS files for a language. Returns the model files plus, for English,
/// the espeak-ng binary + data files listed in the per-platform manifest.
///
/// Performs a synchronous HTTP fetch of the espeak manifest when `lang_code`
/// is `"en"` — call from a context that can block briefly (~100 ms).
pub fn piper_files(lang_code: &str) -> Vec<DownloadItem> {
    let base = data_dir().join("piper");
    let mut items = Vec::new();

    match lang_code {
        "nb" | "nn" => {
            let dir = base.join("nb-NO");
            for (key, fname, label) in &[
                ("epoch_649_v5.onnx", "epoch_649_v5.onnx", "Norsk Piper-modell"),
                ("epoch_649_v5.onnx.json", "epoch_649_v5.onnx.json", "Modellkonfig"),
                ("lexicon.fst", "lexicon.fst", "Ordbok-FST"),
                ("lexicon_values.bin", "lexicon_values.bin", "Ordbok-verdier"),
                ("lexicon_phonemes.txt", "lexicon_phonemes.txt", "Fonemtabell"),
                ("pronunciation_overrides.tsv", "pronunciation_overrides.tsv", "Uttaleregler"),
            ] {
                items.push(DownloadItem {
                    s3_key: format!("models/piper/nb-NO/{}", key),
                    local_path: dir.join(fname),
                    label: (*label).into(),
                });
            }
        }
        "en" => {
            for voice in &[
                "en_US-lessac-medium",
                "en_US-amy-medium",
                "en_GB-alba-medium",
                "en_GB-northern_english_male-medium",
            ] {
                let dir = base.join(voice);
                items.push(DownloadItem {
                    s3_key: format!("models/piper/{}/{}.onnx", voice, voice),
                    local_path: dir.join(format!("{}.onnx", voice)),
                    label: format!("English: {}", voice),
                });
                items.push(DownloadItem {
                    s3_key: format!("models/piper/{}/{}.onnx.json", voice, voice),
                    local_path: dir.join(format!("{}.onnx.json", voice)),
                    label: format!("Config: {}", voice),
                });
            }
            items.extend(piper_espeak_items());
        }
        _ => {}
    }

    items
}

/// Fetch the per-platform espeak-ng manifest and turn each listed file into
/// a `DownloadItem` rooted at `<piper>/bin/`. Synchronous — keep callers off
/// the UI thread.
fn piper_espeak_items() -> Vec<DownloadItem> {
    let bin_root = data_dir().join("piper").join("bin");
    let platform = if cfg!(target_os = "windows") {
        "win"
    } else if cfg!(target_os = "macos") {
        "mac"
    } else {
        return Vec::new();
    };

    let manifest_url = presign_url(&format!("bin/{}/manifest.txt", platform), 3600);
    let manifest = match ureq::get(&manifest_url).call() {
        Ok(r) => r.into_string().unwrap_or_default(),
        Err(e) => {
            eprintln!("Could not fetch espeak manifest for {}: {}", platform, e);
            return Vec::new();
        }
    };

    let mut items = Vec::new();
    for line in manifest.lines() {
        let rel = line.trim();
        if rel.is_empty() {
            continue;
        }
        let label = if rel == "espeak-ng" || rel == "espeak-ng.exe" {
            "espeak-ng".to_string()
        } else if rel.ends_with(".dll") || rel.ends_with(".dylib") {
            "espeak-ng library".to_string()
        } else {
            "espeak-ng data".to_string()
        };
        items.push(DownloadItem {
            s3_key: format!("bin/{}/{}", platform, rel),
            local_path: bin_root.join(rel),
            label,
        });
    }
    items
}

/// True if every Piper file for the language exists locally. For English this
/// also checks that the espeak-ng binary is present (but does not re-fetch the
/// manifest).
pub fn piper_cached(lang_code: &str) -> bool {
    let base = data_dir().join("piper");
    match lang_code {
        "nb" | "nn" => {
            let dir = base.join("nb-NO");
            ["epoch_649_v5.onnx", "lexicon.fst", "lexicon_values.bin"]
                .iter()
                .all(|f| dir.join(f).exists())
        }
        "en" => {
            let lessac = base.join("en_US-lessac-medium").join("en_US-lessac-medium.onnx");
            let espeak = base.join("bin").join(if cfg!(target_os = "windows") {
                "espeak-ng.exe"
            } else {
                "espeak-ng"
            });
            lessac.exists() && espeak.exists()
        }
        _ => true,
    }
}

/// Whisper STT model files for a language.
pub fn whisper_files(lang_code: &str, mode: u8) -> Vec<DownloadItem> {
    let base = data_dir();
    let mut items = Vec::new();

    match lang_code {
        "nb" | "nn" => {
            let dir = base.join("models/whisper/nb");
            if mode == 0 {
                // Rask: tiny only
                items.push(DownloadItem {
                    s3_key: "models/whisper/nb/ggml-nb-whisper-tiny.bin".into(),
                    local_path: dir.join("ggml-nb-whisper-tiny.bin"),
                    label: "Talemodell (rask)".into(),
                });
            } else {
                // Beste: base + medium-q5
                items.push(DownloadItem {
                    s3_key: "models/whisper/nb/ggml-nb-whisper-base.bin".into(),
                    local_path: dir.join("ggml-nb-whisper-base.bin"),
                    label: "Talemodell (strøyming)".into(),
                });
                items.push(DownloadItem {
                    s3_key: "models/whisper/nb/ggml-nb-whisper-medium-q5.bin".into(),
                    local_path: dir.join("ggml-nb-whisper-medium-q5.bin"),
                    label: "Talemodell (beste)".into(),
                });
            }
        }
        _ => {}
    }

    items
}

/// Download all items that aren't already cached.
/// Returns a SharedProgress that the UI can poll.
/// Spawns a background thread; returns immediately.
/// Capture one-time environment context to the head of download_errors.log:
/// OS, system clock, S3 reachability via a HEAD probe. Lets us correlate a
/// later batch of "[1..13] Spelling/Tokenizer/... LOCAL/NET error" lines
/// against the machine state at the moment downloads started. Particularly
/// useful for diagnosing the personal-machine-only failures users report
/// (clock skew, AV TLS interception, DNS blocks) — the failing pattern is
/// often consistent across all 13 items but the root cause is in the
/// environment, not the per-item logic.
fn log_download_session_start(total_items: usize) {
    let log_path = data_dir().join("download_errors.log");
    let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(&log_path)
    else { return };
    let now = chrono::Utc::now();
    let _ = writeln!(f, "");
    let _ = writeln!(f, "═══ session {} ═══", now.format("%Y-%m-%dT%H:%M:%SZ"));
    let _ = writeln!(f, "  items_queued: {}", total_items);
    let _ = writeln!(f, "  os: {} {}", std::env::consts::OS, std::env::consts::ARCH);
    let _ = writeln!(f, "  exe: {}", std::env::current_exe()
        .map(|p| p.display().to_string())
        .unwrap_or_else(|_| "<unknown>".into()));
    // Local clock check vs S3's Date header. If they're > 10 min apart,
    // AWS Sig V4 will reject every presigned URL with 403. Many "all
    // 13 items failed instantly" cases are just a wrong system clock.
    let agent = ureq::AgentBuilder::new()
        .timeout_connect(std::time::Duration::from_secs(10))
        .timeout_read(std::time::Duration::from_secs(10))
        .build();
    let probe_url = format!("{}/", S3_ENDPOINT);
    let head_start = std::time::Instant::now();
    match agent.request("HEAD", &probe_url).call() {
        Ok(r) => {
            let server_date = r.header("Date").unwrap_or("<missing>").to_string();
            let server_hdr = r.header("Server").unwrap_or("").to_string();
            let _ = writeln!(f, "  probe_head: status={} elapsed_ms={} server='{}'",
                r.status(), head_start.elapsed().as_millis(), server_hdr);
            let _ = writeln!(f, "  local_utc:  {}", now.format("%a, %d %b %Y %H:%M:%S GMT"));
            let _ = writeln!(f, "  server_utc: {}", server_date);
            if let Ok(server_t) = chrono::DateTime::parse_from_rfc2822(&server_date) {
                let skew = (now - server_t.with_timezone(&chrono::Utc)).num_seconds();
                let _ = writeln!(f, "  clock_skew_seconds: {}{}", skew,
                    if skew.abs() > 600 { "  ⚠ > 10 MIN — Sig V4 WILL reject" } else { "" });
            }
        }
        Err(e) => {
            let (tag, detail) = classify_ureq_err(e);
            let _ = writeln!(f, "  probe_head: FAILED {} elapsed_ms={} {}",
                tag, head_start.elapsed().as_millis(), detail);
        }
    }
    let _ = writeln!(f, "─── per-item results ───");
}

pub fn download_missing(items: Vec<DownloadItem>) -> SharedProgress {
    // Filter to only items not yet cached
    let needed: Vec<DownloadItem> = items.into_iter()
        .filter(|item| !is_cached(item))
        .collect();

    let progress: SharedProgress = Arc::new(Mutex::new(
        needed.iter().map(|item| DownloadProgress {
            label: item.label.clone(),
            downloaded: 0,
            total: 0,
            done: false,
            error: None,
        }).collect()
    ));

    if needed.is_empty() {
        return progress;
    }

    // Capture environment snapshot once per session — useful for
    // correlating a wave of failures against clock/DNS/TLS state.
    log_download_session_start(needed.len());

    let prog = Arc::clone(&progress);
    std::thread::Builder::new()
        .name("s3-download".into())
        .spawn(move || {
            for (i, item) in needed.iter().enumerate() {
                if let Err(e) = download_one(item, &prog, i) {
                    // Persist to a side log so we can post-mortem failed
                    // downloads in packaged builds (eprintln stderr from a
                    // .app bundle goes nowhere visible). The log file lives
                    // alongside the data dir so it's easy to grep.
                    log_download_error(i + 1, needed.len(), &item.s3_key, &e);
                    eprintln!("Download failed: {} — {}", item.s3_key, e);
                    if let Ok(mut p) = prog.lock() {
                        p[i].error = Some(e);
                        p[i].done = true;
                    }
                }
            }
        })
        .expect("Failed to spawn download thread");

    progress
}

/// Check if all downloads are complete (or errored).
pub fn all_done(progress: &SharedProgress) -> bool {
    if let Ok(p) = progress.lock() {
        p.is_empty() || p.iter().all(|d| d.done)
    } else {
        false
    }
}

/// Check if any download had an error.
pub fn any_error(progress: &SharedProgress) -> Option<String> {
    if let Ok(p) = progress.lock() {
        p.iter().find_map(|d| d.error.clone())
    } else {
        None
    }
}

/// Returns true if all files for a language are already cached locally.
pub fn language_cached(lang_code: &str) -> bool {
    language_files(lang_code).iter().all(|item| is_cached(item))
}
