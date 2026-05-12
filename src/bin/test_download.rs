use acatts_rust::downloader;
use std::env;
use std::fs::{self, File, OpenOptions};
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant};

const DEFAULT_TEST_KEY: &str = "lang/nb/sentence_split.pl";

struct Reporter {
    path: PathBuf,
    file: File,
}

impl Reporter {
    fn new() -> Result<Self, String> {
        let cwd_path = env::current_dir()
            .unwrap_or_else(|_| downloader::data_dir())
            .join("spell-download-diagnostic.log");

        let path = match OpenOptions::new()
            .create(true)
            .write(true)
            .truncate(true)
            .open(&cwd_path)
        {
            Ok(file) => {
                return Ok(Self {
                    path: cwd_path,
                    file,
                });
            }
            Err(_) => downloader::data_dir().join("spell-download-diagnostic.log"),
        };

        let file = OpenOptions::new()
            .create(true)
            .write(true)
            .truncate(true)
            .open(&path)
            .map_err(|e| format!("Could not create report file {}: {}", path.display(), e))?;
        Ok(Self { path, file })
    }

    fn line(&mut self, msg: impl AsRef<str>) {
        let msg = msg.as_ref();
        println!("{}", msg);
        let _ = writeln!(self.file, "{}", msg);
    }

    fn blank(&mut self) {
        self.line("");
    }
}

#[derive(Default)]
struct Args {
    key: String,
    keep_file: bool,
    no_pause: bool,
}

fn main() {
    let args = parse_args();
    let mut report = match Reporter::new() {
        Ok(report) => report,
        Err(e) => {
            eprintln!("{}", e);
            wait_before_exit(&args);
            std::process::exit(2);
        }
    };

    let data_dir = downloader::data_dir();
    let app_download_log = data_dir.join("download_errors.log");
    let download_log_offset = file_len(&app_download_log);
    let target = data_dir
        .join("diagnostics")
        .join(format!("probe_{}", safe_file_name(&args.key)));
    let temp_target = target.with_extension("download");

    report.line("Spell S3 download diagnostic");
    report.line(format!("Started UTC: {}", chrono::Utc::now().to_rfc3339()));
    report.line(format!(
        "OS/arch: {}/{}",
        env::consts::OS,
        env::consts::ARCH
    ));
    report.line(format!("Exe: {}", current_exe()));
    report.line(format!("Working dir: {}", current_dir()));
    report.line(format!("Report file: {}", report.path.display()));
    report.line(format!("Spell data dir: {}", data_dir.display()));
    report.line(format!("App download log: {}", app_download_log.display()));
    report.line(format!("Test S3 key: {}", args.key));
    report.line(format!("Local test path: {}", target.display()));
    report.blank();

    log_proxy_env(&mut report);
    log_basic_network_probe(&mut report);
    report.blank();

    let _ = fs::remove_file(&target);
    let _ = fs::remove_file(&temp_target);

    let items = vec![downloader::DownloadItem {
        s3_key: args.key.clone(),
        local_path: target.clone(),
        label: "S3 diagnostic file".into(),
    }];

    report.line("Starting production downloader path...");
    let start = Instant::now();
    let progress = downloader::download_missing(items);
    let mut last_downloaded = None;

    loop {
        if let Ok(p) = progress.lock() {
            if p.is_empty() {
                report.line("No download was queued. The test file may already be cached.");
                break;
            }

            let d = &p[0];
            if last_downloaded != Some(d.downloaded) || d.done {
                let pct = if d.total > 0 {
                    d.downloaded * 100 / d.total
                } else {
                    0
                };
                report.line(format!(
                    "Progress: {} / {} bytes ({}%), done={}, error={}",
                    d.downloaded,
                    d.total,
                    pct,
                    d.done,
                    d.error.as_deref().unwrap_or("")
                ));
                last_downloaded = Some(d.downloaded);
            }
        }

        if downloader::all_done(&progress) {
            break;
        }
        std::thread::sleep(Duration::from_millis(200));
    }

    report.blank();
    if let Some(err) = downloader::any_error(&progress) {
        report.line("RESULT: FAIL");
        report.line(format!("Downloader error: {}", err));
    } else {
        report.line("RESULT: OK");
        report.line(format!("Elapsed: {:.3}s", start.elapsed().as_secs_f64()));
        match fs::metadata(&target) {
            Ok(meta) => report.line(format!("Downloaded size: {} bytes", meta.len())),
            Err(e) => report.line(format!("Downloaded file metadata failed: {}", e)),
        }
    }

    append_app_download_log(&mut report, &app_download_log, download_log_offset);

    if !args.keep_file {
        let _ = fs::remove_file(&target);
        let _ = fs::remove_file(&temp_target);
        report.line("Cleaned up diagnostic download file.");
    } else {
        report.line(format!("Kept diagnostic file: {}", target.display()));
    }

    report.blank();
    report.line(format!(
        "Send this report file back to support: {}",
        report.path.display()
    ));
    wait_before_exit(&args);
}

fn parse_args() -> Args {
    let mut args = Args {
        key: DEFAULT_TEST_KEY.to_string(),
        keep_file: false,
        no_pause: false,
    };

    for arg in env::args().skip(1) {
        match arg.as_str() {
            "--keep" => args.keep_file = true,
            "--no-pause" => args.no_pause = true,
            "--help" | "-h" => {
                println!("Usage: test_download.exe [S3_KEY] [--keep] [--no-pause]");
                println!("Default S3_KEY: {}", DEFAULT_TEST_KEY);
                std::process::exit(0);
            }
            _ if arg.starts_with("--") => {
                eprintln!("Unknown option: {}", arg);
                std::process::exit(2);
            }
            _ => args.key = arg,
        }
    }

    args
}

fn safe_file_name(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        if ch.is_ascii_alphanumeric() || matches!(ch, '.' | '-' | '_') {
            out.push(ch);
        } else {
            out.push('_');
        }
    }
    out
}

fn current_exe() -> String {
    env::current_exe()
        .map(|p| p.display().to_string())
        .unwrap_or_else(|e| format!("<unknown: {}>", e))
}

fn current_dir() -> String {
    env::current_dir()
        .map(|p| p.display().to_string())
        .unwrap_or_else(|e| format!("<unknown: {}>", e))
}

fn file_len(path: &Path) -> u64 {
    fs::metadata(path).map(|m| m.len()).unwrap_or(0)
}

fn log_proxy_env(report: &mut Reporter) {
    report.line("Proxy environment:");
    for key in ["HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "NO_PROXY"] {
        let value = env::var(key).unwrap_or_default();
        if value.is_empty() {
            report.line(format!("  {}=<unset>", key));
        } else {
            report.line(format!("  {}={}", key, value));
        }
    }
}

fn log_basic_network_probe(report: &mut Reporter) {
    use std::net::{TcpStream, ToSocketAddrs};

    report.line("Basic network probe:");
    let start = Instant::now();
    match ("eu2.contabostorage.com", 443).to_socket_addrs() {
        Ok(addrs) => {
            let addrs: Vec<_> = addrs.collect();
            report.line(format!(
                "  DNS OK in {} ms: {:?}",
                start.elapsed().as_millis(),
                addrs
            ));
            if let Some(addr) = addrs.first() {
                let tcp_start = Instant::now();
                match TcpStream::connect_timeout(addr, Duration::from_secs(10)) {
                    Ok(_) => report.line(format!(
                        "  TCP 443 OK to {} in {} ms",
                        addr,
                        tcp_start.elapsed().as_millis()
                    )),
                    Err(e) => report.line(format!(
                        "  TCP 443 FAILED to {} after {} ms: {}",
                        addr,
                        tcp_start.elapsed().as_millis(),
                        e
                    )),
                }
            }
        }
        Err(e) => report.line(format!(
            "  DNS FAILED after {} ms: {}",
            start.elapsed().as_millis(),
            e
        )),
    }
}

fn append_app_download_log(report: &mut Reporter, path: &Path, offset: u64) {
    report.blank();
    report.line("Production download_errors.log entries from this run:");

    let mut file = match File::open(path) {
        Ok(file) => file,
        Err(e) => {
            report.line(format!("  Could not open {}: {}", path.display(), e));
            return;
        }
    };

    if let Err(e) = file.seek(SeekFrom::Start(offset)) {
        report.line(format!("  Could not seek {}: {}", path.display(), e));
        return;
    }

    let mut buf = String::new();
    match file.read_to_string(&mut buf) {
        Ok(_) if buf.trim().is_empty() => {
            report.line("  <no new production log lines>");
        }
        Ok(_) => {
            for line in buf.lines() {
                report.line(format!("  {}", line));
            }
        }
        Err(e) => report.line(format!("  Could not read {}: {}", path.display(), e)),
    }
}

fn wait_before_exit(args: &Args) {
    if !cfg!(windows) || args.no_pause {
        return;
    }
    println!();
    println!("Press Enter to close this window.");
    let mut s = String::new();
    let _ = std::io::stdin().read_line(&mut s);
}
