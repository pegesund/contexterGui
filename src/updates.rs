//! Velopack-backed auto-update service.
//!
//! Mirrors `ConcentrateDotNet/Services/Updates/UpdateService.cs`:
//!   - GithubSource against the binaries repo (`pegesund/spell_binaries`)
//!   - Polls on startup + every 6h
//!   - Returns Available(version) so the UI can render a banner
//!   - Download + apply + restart on user click
//!
//! Velopack only operates when the app was installed via a Velopack package
//! (Spell.app produced by `vpk pack`). For dev runs, raw cargo builds, or
//! the legacy `create-dmg`-only DMG, `UpdateManager::new` errors out at the
//! locator step and we keep `Status::NotInstalled` permanently. UI hides the
//! updater bits in that case.
//!
//! API note: the velopack 0.0.1589 Rust crate uses PascalCase struct fields
//! (`TargetFullRelease`, `Version`) because they're serde-deserialized from
//! the .NET-side JSON manifest. Hence the `#[allow(non_snake_case)]` shape
//! when we read those — not a style choice on our side.

use std::sync::{Arc, Mutex};
use std::time::Duration;

use velopack::sources::GithubSource;
use velopack::{UpdateCheck, UpdateInfo, UpdateManager};

const RELEASES_REPO_URL: &str = "https://github.com/pegesund/spell_binaries";
/// Re-check at this cadence after the initial startup poll. 6 hours matches
/// Concentrate's PreLoginUpdateGate cadence and is gentle on the GitHub API
/// rate limit (60 unauthenticated req/hr per IP — we use ~4/day).
const POLL_INTERVAL: Duration = Duration::from_secs(6 * 60 * 60);

/// Snapshot of what the UI needs to render.
#[derive(Debug, Clone)]
pub enum Status {
    /// App wasn't installed via Velopack (dev build, old DMG, etc). Hide the
    /// updater UI entirely.
    NotInstalled,
    /// Velopack is wired up but the last poll hasn't fired yet.
    Idle,
    /// A poll is currently in flight.
    Checking,
    /// Latest poll found no update.
    UpToDate,
    /// A newer release is available. `version` is the human-readable tag
    /// (e.g. "0.1.1") shown in the banner.
    Available { version: String },
    /// Download in progress.
    Downloading,
    /// Download finished; clicking the banner restarts into the new version.
    Ready,
    /// Last poll or download errored. Stored for diagnostics; the banner
    /// shows a retry button rather than the raw error.
    Error { message: String },
}

/// Thread-safe update service. UI thread reads `status()`; a background thread
/// owns polling + downloads.
pub struct UpdateService {
    inner: Arc<Inner>,
}

struct Inner {
    status: Mutex<Status>,
    /// `None` outside of an Available/Ready state. Held so the UI's
    /// "Last ned" / "Start på nytt" button can hand it back to apply.
    pending: Mutex<Option<UpdateInfo>>,
    /// `None` when the app isn't Velopack-installed (dev runs, plain DMG,
    /// etc). All other methods become no-ops in that case.
    manager: Option<UpdateManager>,
    current_version: Option<String>,
}

impl UpdateService {
    /// Build a service. If the app isn't Velopack-installed (dev runs,
    /// manual DMG), the constructor still succeeds but every operation is
    /// a no-op and `status()` reports `NotInstalled` — caller can construct
    /// unconditionally and the UI hides itself based on `status()`.
    pub fn new() -> Self {
        let source = GithubSource::new(RELEASES_REPO_URL, None, false);
        // None options → defaults; None locator → auto-locate which fails
        // (returns Err) when this exe isn't sitting inside a Velopack-managed
        // app bundle. We treat that as the "NotInstalled" sentinel.
        let manager = UpdateManager::new(source, None, None).ok();
        let current_version = manager.as_ref().map(|m| m.get_current_version_as_string());

        let initial = if manager.is_some() {
            Status::Idle
        } else {
            Status::NotInstalled
        };

        Self {
            inner: Arc::new(Inner {
                status: Mutex::new(initial),
                pending: Mutex::new(None),
                manager,
                current_version,
            }),
        }
    }

    pub fn status(&self) -> Status {
        self.inner.status.lock().map(|s| s.clone()).unwrap_or(Status::Idle)
    }

    /// Velopack-managed version string (e.g. "0.1.0"). `None` outside a
    /// Velopack install — the UI falls back to `env!("CARGO_PKG_VERSION")`.
    pub fn current_version(&self) -> Option<String> {
        self.inner.current_version.clone()
    }

    /// Spawns the polling thread. Idempotent — safe to call once at startup.
    /// No-op when not Velopack-installed.
    pub fn start_polling(&self) {
        if self.inner.manager.is_none() {
            return;
        }
        let inner = Arc::clone(&self.inner);
        std::thread::Builder::new()
            .name("update-poller".into())
            .spawn(move || loop {
                Self::run_check(&inner);
                std::thread::sleep(POLL_INTERVAL);
            })
            .ok();
    }

    /// Manual re-check trigger — wired to the "Sjekk på nytt" button when an
    /// earlier poll errored.
    pub fn check_now(&self) {
        let inner = Arc::clone(&self.inner);
        std::thread::spawn(move || Self::run_check(&inner));
    }

    /// Download + apply + restart. Called when the user clicks the banner.
    /// Spawned in a background thread; the UI just sets Downloading and the
    /// process is terminated by the bootstrapper inside
    /// `apply_updates_and_restart`.
    pub fn download_and_restart(&self) {
        let inner = Arc::clone(&self.inner);
        std::thread::spawn(move || Self::run_download(&inner));
    }

    fn run_check(inner: &Arc<Inner>) {
        let Some(manager) = inner.manager.as_ref() else { return };
        Self::set_status(inner, Status::Checking);
        match manager.check_for_updates() {
            Ok(UpdateCheck::UpdateAvailable(info)) => {
                let version = info.TargetFullRelease.Version.clone();
                *inner.pending.lock().unwrap() = Some(info);
                Self::set_status(inner, Status::Available { version });
            }
            Ok(UpdateCheck::NoUpdateAvailable) | Ok(UpdateCheck::RemoteIsEmpty) => {
                *inner.pending.lock().unwrap() = None;
                Self::set_status(inner, Status::UpToDate);
            }
            Err(e) => {
                Self::set_status(
                    inner,
                    Status::Error { message: format!("{}", e) },
                );
            }
        }
    }

    fn run_download(inner: &Arc<Inner>) {
        let Some(manager) = inner.manager.as_ref() else { return };
        let Some(info) = inner.pending.lock().unwrap().clone() else { return };
        Self::set_status(inner, Status::Downloading);
        match manager.download_updates(&info, None) {
            Ok(()) => {
                Self::set_status(inner, Status::Ready);
                // apply_updates_and_restart never returns on success — it
                // execs Velopack's update binary which terminates us.
                if let Err(e) = manager.apply_updates_and_restart(&info) {
                    Self::set_status(
                        inner,
                        Status::Error { message: format!("apply failed: {}", e) },
                    );
                }
            }
            Err(e) => {
                Self::set_status(
                    inner,
                    Status::Error { message: format!("download failed: {}", e) },
                );
            }
        }
    }

    fn set_status(inner: &Arc<Inner>, status: Status) {
        if let Ok(mut s) = inner.status.lock() {
            *s = status;
        }
    }
}

impl Default for UpdateService {
    fn default() -> Self {
        Self::new()
    }
}
