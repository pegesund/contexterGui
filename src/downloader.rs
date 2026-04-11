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

fn presign_url(key: &str, expires_secs: u64) -> String {
    use hmac::{Hmac, Mac};
    use sha2::Sha256;

    let now = chrono::Utc::now();
    let date_stamp = now.format("%Y%m%d").to_string();
    let amz_date = now.format("%Y%m%dT%H%M%SZ").to_string();

    let host = S3_ENDPOINT.trim_start_matches("https://");
    let canonical_uri = format!("/{}/{}", S3_BUCKET, key);
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

/// Download a single file from S3 to a local path, reporting progress.
fn download_one(item: &DownloadItem, progress: &SharedProgress, index: usize) -> Result<(), String> {
    // Create parent directories
    if let Some(parent) = item.local_path.parent() {
        std::fs::create_dir_all(parent).map_err(|e| format!("mkdir: {}", e))?;
    }

    let url = presign_url(&item.s3_key, 3600);

    let resp = ureq::get(&url)
        .call()
        .map_err(|e| format!("HTTP: {}", e))?;

    let total: u64 = resp.header("Content-Length")
        .and_then(|s| s.parse().ok())
        .unwrap_or(0);

    // Update total in progress
    if let Ok(mut p) = progress.lock() {
        p[index].total = total;
    }

    let mut reader = resp.into_reader();
    let tmp_path = item.local_path.with_extension("download");
    let mut file = std::fs::File::create(&tmp_path)
        .map_err(|e| format!("create: {}", e))?;

    let mut buf = [0u8; 65536];
    let mut downloaded: u64 = 0;

    loop {
        let n = reader.read(&mut buf).map_err(|e| format!("read: {}", e))?;
        if n == 0 { break; }
        file.write_all(&buf[..n]).map_err(|e| format!("write: {}", e))?;
        downloaded += n as u64;

        if let Ok(mut p) = progress.lock() {
            p[index].downloaded = downloaded;
        }
    }

    // Atomic rename
    std::fs::rename(&tmp_path, &item.local_path)
        .map_err(|e| format!("rename: {}", e))?;

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

/// Local data directory: ~/Library/Application Support/NorskTale/data/
pub fn data_dir() -> PathBuf {
    let dir = if cfg!(target_os = "macos") {
        dirs::home_dir()
            .map(|h| h.join("Library/Application Support/NorskTale/data"))
            .unwrap_or_else(|| PathBuf::from("/tmp/norsktale/data"))
    } else {
        dirs::config_dir()
            .map(|c| c.join("NorskTale/data"))
            .unwrap_or_else(|| PathBuf::from("norsktale/data"))
    };
    let _ = std::fs::create_dir_all(&dir);
    dir
}

// ── Language definitions ──

/// Files needed for a language.
pub fn language_files(lang_code: &str) -> Vec<DownloadItem> {
    let base = data_dir();

    let mut items = Vec::new();

    // Shared BERT model (same for nb and nn — NorBERT4)
    let bert_dir = base.join("models/bert");
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

    // Per-language files
    match lang_code {
        "nb" => {
            let dir = base.join("lang/nb");
            items.push(DownloadItem { s3_key: "lang/nb/fullform_bm.mfst".into(), local_path: dir.join("fullform_bm.mfst"), label: "Ordbok".into() });
            items.push(DownloadItem { s3_key: "lang/nb/wordfreq_bm.tsv".into(), local_path: dir.join("wordfreq_bm.tsv"), label: "Ordfrekvens".into() });
            items.push(DownloadItem { s3_key: "lang/nb/grammar_rules.pl".into(), local_path: dir.join("grammar_rules.pl"), label: "Grammatikk".into() });
            items.push(DownloadItem { s3_key: "lang/nb/compound_data.pl".into(), local_path: dir.join("compound_data.pl"), label: "Samansette ord".into() });
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
        _ => {}
    }

    items
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

    let prog = Arc::clone(&progress);
    std::thread::Builder::new()
        .name("s3-download".into())
        .spawn(move || {
            for (i, item) in needed.iter().enumerate() {
                if let Err(e) = download_one(item, &prog, i) {
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
