//! Mac Word add-in setup wizard.
//!
//! What this does:
//!   1. Detects whether Microsoft Word is installed
//!   2. Generates a per-user localhost CA + leaf cert (rcgen; pure Rust)
//!   3. Adds the CA to /Library/Keychains/System.keychain via the `security`
//!      command (one graphical sudo prompt via osascript)
//!   4. Copies the bundled manifest.xml into Word's wef folder so Word loads
//!      Spell as an add-in
//!
//! Re-run-safe: each step is idempotent. If the IT admin has already done
//! cert + manifest deployment via MDM/M365, the status check returns Ready
//! and the wizard skips everything.

use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

use anyhow::{anyhow, Context, Result};
use rcgen::{
    BasicConstraints, CertificateParams, DnType, ExtendedKeyUsagePurpose, IsCa, KeyPair,
};

const CA_COMMON_NAME: &str = "Spell Word Add-in Local CA";
const MANIFEST_FILENAME: &str = "Spell-manifest.xml";

// ── Paths ────────────────────────────────────────────────────────────────────

/// Per-user cert storage at `~/Library/Application Support/Spell/word-addin-certs/`.
/// Lives outside the .app so it survives app upgrades.
fn cert_dir() -> Result<PathBuf> {
    let dir = dirs::data_dir()
        .or_else(dirs::home_dir)
        .ok_or_else(|| anyhow!("no home dir"))?
        .join("Spell")
        .join("word-addin-certs");
    fs::create_dir_all(&dir).context("create cert dir")?;
    Ok(dir)
}

pub fn ca_cert_path() -> Result<PathBuf> {
    Ok(cert_dir()?.join("rootCA.pem"))
}
pub fn ca_key_path() -> Result<PathBuf> {
    Ok(cert_dir()?.join("rootCA-key.pem"))
}
pub fn leaf_cert_path() -> Result<PathBuf> {
    Ok(cert_dir()?.join("fullchain.pem"))
}
pub fn leaf_key_path() -> Result<PathBuf> {
    Ok(cert_dir()?.join("key.pem"))
}

/// Word's add-in manifest folder on macOS.
pub fn word_wef_dir() -> Option<PathBuf> {
    let home = dirs::home_dir()?;
    Some(home.join("Library/Containers/com.microsoft.Word/Data/Documents/wef"))
}

/// True iff `/Applications/Microsoft Word.app` exists.
pub fn is_word_installed() -> bool {
    Path::new("/Applications/Microsoft Word.app").exists()
}

/// Path inside the running .app where the manifest is bundled (built by
/// scripts/build-mac.sh into `Contents/Resources/word-addin/manifest.xml`).
fn bundled_manifest_path() -> Result<PathBuf> {
    let exe = std::env::current_exe().context("current_exe")?;
    let macos = exe.parent().ok_or_else(|| anyhow!("no exe parent"))?;
    let contents = macos.parent().ok_or_else(|| anyhow!("no Contents parent"))?;
    Ok(contents.join("Resources/word-addin/manifest.xml"))
}

// ── Status ───────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SetupStatus {
    /// Microsoft Word isn't installed — wizard has nothing to do.
    NoWord,
    /// Cert installed + trusted, manifest in wef folder. Nothing needed.
    Ready,
    /// One or more pieces are missing. Run `run_full_setup()`.
    NeedsSetup,
}

pub fn check_status() -> SetupStatus {
    if !is_word_installed() {
        return SetupStatus::NoWord;
    }
    let cert_present = ca_cert_path()
        .ok()
        .and_then(|p| if p.exists() { Some(()) } else { None })
        .is_some();
    let trusted = is_ca_trusted();
    let manifest_in_place = is_manifest_installed();
    if cert_present && trusted && manifest_in_place {
        SetupStatus::Ready
    } else {
        SetupStatus::NeedsSetup
    }
}

/// Path to the user's login keychain on modern macOS.
fn login_keychain_path() -> Option<PathBuf> {
    Some(
        dirs::home_dir()?
            .join("Library/Keychains/login.keychain-db"),
    )
}

pub fn is_ca_trusted() -> bool {
    // Check user's login keychain first (where this version installs the
    // cert). Fall back to System.keychain so Macs that were set up with the
    // older sudo-based wizard still report Ready.
    if let Some(user_kc) = login_keychain_path() {
        let in_user = Command::new("security")
            .args(["find-certificate", "-c", CA_COMMON_NAME])
            .arg(&user_kc)
            .output()
            .map(|o| o.status.success())
            .unwrap_or(false);
        if in_user {
            return true;
        }
    }
    Command::new("security")
        .args([
            "find-certificate",
            "-c",
            CA_COMMON_NAME,
            "/Library/Keychains/System.keychain",
        ])
        .output()
        .map(|o| o.status.success())
        .unwrap_or(false)
}

pub fn is_manifest_installed() -> bool {
    word_wef_dir()
        .map(|d| d.join(MANIFEST_FILENAME).exists())
        .unwrap_or(false)
}

// ── Cert generation ──────────────────────────────────────────────────────────

/// Generate the per-user CA + leaf cert if any of the four files is missing.
/// Idempotent: re-running with all four present is a no-op.
pub fn generate_certs_if_missing() -> Result<()> {
    let ca_pem = ca_cert_path()?;
    let ca_key = ca_key_path()?;
    let leaf_pem = leaf_cert_path()?;
    let leaf_key = leaf_key_path()?;
    if ca_pem.exists() && ca_key.exists() && leaf_pem.exists() && leaf_key.exists() {
        return Ok(());
    }

    // CA — self-signed root, 10 years.
    let ca_keypair = KeyPair::generate().context("generate CA key")?;
    let mut ca_params = CertificateParams::default();
    ca_params.is_ca = IsCa::Ca(BasicConstraints::Unconstrained);
    ca_params
        .distinguished_name
        .push(DnType::CommonName, CA_COMMON_NAME);
    ca_params
        .distinguished_name
        .push(DnType::OrganizationName, "Cognio AS");
    ca_params.not_before = rcgen::date_time_ymd(2024, 1, 1);
    ca_params.not_after = rcgen::date_time_ymd(2034, 1, 1);
    let ca_cert = ca_params
        .self_signed(&ca_keypair)
        .context("self-sign CA")?;

    // Leaf — signed by our CA, valid for localhost + 127.0.0.1, 5 years.
    let leaf_keypair = KeyPair::generate().context("generate leaf key")?;
    let mut leaf_params =
        CertificateParams::new(vec!["localhost".to_string(), "127.0.0.1".to_string()])
            .context("leaf params")?;
    leaf_params
        .distinguished_name
        .push(DnType::CommonName, "localhost");
    leaf_params
        .distinguished_name
        .push(DnType::OrganizationName, "Cognio AS");
    leaf_params
        .extended_key_usages
        .push(ExtendedKeyUsagePurpose::ServerAuth);
    leaf_params.not_before = rcgen::date_time_ymd(2024, 1, 1);
    leaf_params.not_after = rcgen::date_time_ymd(2029, 1, 1);
    let leaf = leaf_params
        .signed_by(&leaf_keypair, &ca_cert, &ca_keypair)
        .context("sign leaf with CA")?;

    fs::write(&ca_pem, ca_cert.pem())?;
    fs::write(&ca_key, ca_keypair.serialize_pem())?;
    fs::write(&leaf_pem, leaf.pem())?;
    fs::write(&leaf_key, leaf_keypair.serialize_pem())?;

    // Lock down private keys (POSIX chmod 600). System.keychain doesn't need
    // access to them — the rust HTTPS server reads them in-process.
    use std::os::unix::fs::PermissionsExt;
    fs::set_permissions(&ca_key, fs::Permissions::from_mode(0o600))?;
    fs::set_permissions(&leaf_key, fs::Permissions::from_mode(0o600))?;

    Ok(())
}

// ── System trust install ─────────────────────────────────────────────────────

/// Add the per-user CA to the user's login keychain so Word, Safari, and other
/// native apps trust localhost HTTPS from Spell.
///
/// **No admin password required** — the login keychain is unlocked while the
/// user is logged in, and writing to it doesn't escalate privileges. macOS's
/// TLS validation consults both System.keychain and login.keychain-db by
/// default, so a user-domain trust is sufficient for Word.
///
/// This deliberately avoids `osascript ... with administrator privileges`:
/// (a) the dialog title says "osascript" which spooks users and flags
/// MDM "suspicious script execution" rules in school deployments, (b) school-
/// managed Macs typically deny pupils the admin password, blocking the
/// wizard entirely.
pub fn install_ca_trust() -> Result<()> {
    let ca_pem = ca_cert_path()?;
    if !ca_pem.exists() {
        return Err(anyhow!(
            "CA cert not found at {} — call generate_certs_if_missing() first",
            ca_pem.display()
        ));
    }
    let user_kc = login_keychain_path()
        .ok_or_else(|| anyhow!("can't locate user login keychain"))?;

    // -r trustRoot   → treat as a trust anchor (validates SSL from anything signed by it)
    // -k <keychain>  → install in user's keychain (no sudo)
    // (no -d flag    → user domain, NOT admin/system domain)
    let output = Command::new("security")
        .args(["add-trusted-cert", "-r", "trustRoot", "-k"])
        .arg(&user_kc)
        .arg(&ca_pem)
        .output()
        .context("invoke security add-trusted-cert")?;

    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(anyhow!(
            "security add-trusted-cert failed (exit {}): {}",
            output.status.code().unwrap_or(-1),
            stderr.trim()
        ));
    }
    Ok(())
}

// ── Manifest install ─────────────────────────────────────────────────────────

/// Copy the bundled manifest into Word's wef folder under a filename
/// unique to Spell so it doesn't collide with other add-ins.
pub fn install_manifest() -> Result<()> {
    let wef = word_wef_dir().ok_or_else(|| anyhow!("can't locate Word wef folder"))?;
    fs::create_dir_all(&wef).context("create wef folder")?;

    let bundled = bundled_manifest_path()?;
    if !bundled.exists() {
        return Err(anyhow!(
            "bundled manifest not found at {} — Spell.app may be missing Resources/word-addin/manifest.xml",
            bundled.display()
        ));
    }

    let target = wef.join(MANIFEST_FILENAME);
    fs::copy(&bundled, &target).context("copy manifest")?;
    Ok(())
}

pub fn uninstall_manifest() -> Result<()> {
    if let Some(wef) = word_wef_dir() {
        let target = wef.join(MANIFEST_FILENAME);
        if target.exists() {
            fs::remove_file(&target).context("remove manifest")?;
        }
    }
    Ok(())
}

// ── Orchestration ────────────────────────────────────────────────────────────

/// Run the full first-launch setup. Each step is idempotent so this is safe to
/// re-run if a previous attempt was interrupted partway through.
pub fn run_full_setup() -> Result<()> {
    if !is_word_installed() {
        return Err(anyhow!("Microsoft Word is not installed"));
    }
    generate_certs_if_missing()?;
    if !is_ca_trusted() {
        install_ca_trust()?;
    }
    install_manifest()?;
    Ok(())
}
