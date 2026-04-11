use acatts_rust::downloader;
use std::time::{Duration, Instant};

fn main() {
    println!("Testing S3 download...");
    println!("Data dir: {}", downloader::data_dir().display());

    // Download just the small wordfreq file to test
    let items = vec![
        downloader::DownloadItem {
            s3_key: "lang/nb/wordfreq_bm.tsv".into(),
            local_path: downloader::data_dir().join("test_wordfreq.tsv"),
            label: "Ordfrekvens (test)".into(),
        },
    ];

    // Remove cached file to force download
    let _ = std::fs::remove_file(downloader::data_dir().join("test_wordfreq.tsv"));

    let progress = downloader::download_missing(items);
    let start = Instant::now();

    loop {
        if downloader::all_done(&progress) { break; }

        if let Ok(p) = progress.lock() {
            for d in p.iter() {
                let pct = if d.total > 0 { d.downloaded * 100 / d.total } else { 0 };
                println!("  {} — {}/{} bytes ({}%)", d.label, d.downloaded, d.total, pct);
            }
        }

        std::thread::sleep(Duration::from_millis(200));
    }

    if let Some(err) = downloader::any_error(&progress) {
        println!("ERROR: {}", err);
    } else {
        let elapsed = start.elapsed();
        println!("Done in {:.1}s", elapsed.as_secs_f64());

        // Verify file
        let path = downloader::data_dir().join("test_wordfreq.tsv");
        let size = std::fs::metadata(&path).map(|m| m.len()).unwrap_or(0);
        println!("Downloaded file: {} ({} bytes)", path.display(), size);

        // Cleanup test file
        let _ = std::fs::remove_file(&path);
    }
}
