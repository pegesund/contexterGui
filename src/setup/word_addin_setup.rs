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
    // Authoritative location: System.keychain (admin domain). Office add-in
    // WKWebView only honors trust anchors here.
    let in_system = Command::new("security")
        .args([
            "find-certificate",
            "-c",
            CA_COMMON_NAME,
            "/Library/Keychains/System.keychain",
        ])
        .output()
        .map(|o| o.status.success())
        .unwrap_or(false);
    if in_system {
        return true;
    }
    // Backwards-compat: 0.1.0-test7 briefly used login keychain. If a user has
    // that cert leftover but no System one, treat as not-yet-trusted so the
    // wizard re-prompts and installs to the right place.
    false
}

pub fn is_manifest_installed() -> bool {
    word_wef_dir()
        .map(|d| d.join(MANIFEST_FILENAME).exists())
        .unwrap_or(false)
}

// ── Cert generation ──────────────────────────────────────────────────────────

/// Generate the per-user CA + leaf cert. Always regenerates the LEAF on every
/// call (cheap, ~50ms) so we never ship a near-expired cert. The CA is only
/// generated once — its trust is what's installed in System.keychain, so
/// regenerating it would break Word until the user re-runs the wizard.
///
/// **Critical**: Apple's TLS policy requires SSL leaf certs to have a validity
/// period of 398 days or fewer (issued ≥ 2020-09-01). WKWebView (used by
/// Office add-ins) silently rejects longer-validity certs as
/// "isn't signed by a valid security certificate" even when the chain is
/// trusted. curl + openssl don't enforce this so the cert appears fine
/// in shell tests but fails inside Word. We use 397 days to stay under the
/// limit. The leaf is regenerated on every wizard run / app start to avoid
/// the cert silently expiring after a year of use.
pub fn generate_certs_if_missing() -> Result<()> {
    let ca_pem = ca_cert_path()?;
    let ca_key = ca_key_path()?;
    let leaf_pem = leaf_cert_path()?;
    let leaf_key = leaf_key_path()?;

    // CA — generate ONCE per machine. Re-using the same key preserves the
    // System.keychain trust install. We rebuild the CA's Certificate struct
    // in-memory each call (using the persisted key) so we can re-sign the
    // leaf — the CA *file* on disk is left alone (don't want to invalidate
    // the trust anchor that's already in System.keychain).
    fn build_ca_params() -> CertificateParams {
        let mut p = CertificateParams::default();
        p.is_ca = IsCa::Ca(BasicConstraints::Unconstrained);
        p.distinguished_name
            .push(DnType::CommonName, CA_COMMON_NAME);
        p.distinguished_name
            .push(DnType::OrganizationName, "Cognio AS");
        // CA validity isn't restricted by Apple's leaf-cert policy. 10 years
        // is fine and matches mkcert's default.
        p.not_before = rcgen::date_time_ymd(2024, 1, 1);
        p.not_after = rcgen::date_time_ymd(2034, 1, 1);
        p
    }

    let (ca_cert, ca_keypair) = if ca_pem.exists() && ca_key.exists() {
        let ca_key_pem = fs::read_to_string(&ca_key).context("read CA key")?;
        let ca_kp = KeyPair::from_pem(&ca_key_pem).context("parse CA key")?;
        // Re-sign with the same params + key. Subject DN + public key match
        // the CA cert in System.keychain → leaf signed with this validates
        // against that trust anchor.
        let ca = build_ca_params()
            .self_signed(&ca_kp)
            .context("re-sign CA from existing key")?;
        (ca, ca_kp)
    } else {
        let ca_kp = KeyPair::generate().context("generate CA key")?;
        let ca = build_ca_params()
            .self_signed(&ca_kp)
            .context("self-sign CA")?;
        fs::write(&ca_pem, ca.pem())?;
        fs::write(&ca_key, ca_kp.serialize_pem())?;
        use std::os::unix::fs::PermissionsExt;
        fs::set_permissions(&ca_key, fs::Permissions::from_mode(0o600))?;
        (ca, ca_kp)
    };

    // LEAF — always regenerate. Apple's max validity is 398 days; we use 397
    // to stay safely under it. Regenerated on every call so the cert never
    // gets close to expiring during normal use.
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
    // Use chrono (already in deps) to compute current + 397 days, then convert
    // to rcgen's date format. Avoids adding a separate `time` dep.
    {
        use chrono::{Datelike, Duration as ChronoDuration, Utc};
        let now = Utc::now();
        let later = now + ChronoDuration::days(397);
        leaf_params.not_before = rcgen::date_time_ymd(
            now.year(), now.month() as u8, now.day() as u8,
        );
        leaf_params.not_after = rcgen::date_time_ymd(
            later.year(), later.month() as u8, later.day() as u8,
        );
    }
    let leaf = leaf_params
        .signed_by(&leaf_keypair, &ca_cert, &ca_keypair)
        .context("sign leaf with CA")?;

    // fullchain.pem = leaf cert FOLLOWED BY CA cert (standard PEM chain order).
    // rustls reads the file as a chain and serves all certs in the TLS handshake
    // so strict TLS clients (like Office's WKWebView) can validate without
    // needing to look up the CA separately. Many TLS clients require the chain
    // in the handshake even when the CA is in their trust store.
    let chain_pem = format!("{}{}", leaf.pem(), ca_cert.pem());
    fs::write(&leaf_pem, chain_pem)?;
    fs::write(&leaf_key, leaf_keypair.serialize_pem())?;
    use std::os::unix::fs::PermissionsExt;
    fs::set_permissions(&leaf_key, fs::Permissions::from_mode(0o600))?;

    Ok(())
}

// ── System trust install ─────────────────────────────────────────────────────

/// Add the per-user CA to `/Library/Keychains/System.keychain` so Word, Safari,
/// and other native apps trust localhost HTTPS from Spell.
///
/// **Admin password required.** Office add-ins on macOS use WKWebView, which
/// only honors trust anchors in the admin (system) domain. School deployments
/// where pupils don't have admin can pre-install the cert via MDM (see
/// docs/SCHOOL_DEPLOYMENT.md).
///
/// Uses macOS's AuthorizationServices framework directly. The native password
/// dialog shows "Spell" as the requesting app (NOT "osascript") with our
/// custom prompt text. Replaces the previous osascript+sudo flow which had
/// a "first attempt fails, retry succeeds" quirk and looked unprofessional.
pub fn install_ca_trust() -> Result<()> {
    let ca_pem = ca_cert_path()?;
    if !ca_pem.exists() {
        return Err(anyhow!(
            "CA cert not found at {} — call generate_certs_if_missing() first",
            ca_pem.display()
        ));
    }

    let ca_pem_str = ca_pem
        .to_str()
        .ok_or_else(|| anyhow!("CA cert path contains invalid UTF-8"))?;

    // /usr/bin/security add-trusted-cert -d -r trustRoot \
    //   -k /Library/Keychains/System.keychain <ca_pem>
    // -d            → admin domain (System.keychain) — REQUIRED for Office add-ins
    // -r trustRoot  → trust anchor for all policies
    auth::run_with_admin(
        "/usr/bin/security",
        &[
            "add-trusted-cert",
            "-d",
            "-r",
            "trustRoot",
            "-k",
            "/Library/Keychains/System.keychain",
            ca_pem_str,
        ],
        "Spell needs to install a security certificate so Microsoft Word can trust the local connection.",
    )
}

/// Tiny FFI wrapper around macOS's AuthorizationServices framework.
///
/// `osascript ... with administrator privileges` shows a dialog titled
/// "osascript" and triggers MDM "suspicious script execution" alerts in school
/// environments. The proper API for a GUI app to elevate is
/// AuthorizationCreate + AuthorizationCopyRights (shows "Spell" with our app
/// icon and a custom prompt) + AuthorizationExecuteWithPrivileges (runs the
/// command as root).
///
/// AuthorizationExecuteWithPrivileges is technically deprecated since 10.7
/// but Apple's official replacement (SMJobBless / SMAppService) requires
/// shipping a separate privileged helper tool. For a one-off cert install at
/// first launch, the deprecated API is the pragmatic choice — it still works
/// in macOS 26 and is what mkcert, Tunnelblick, Hyper-V Mac, and many other
/// tools use today.
mod auth {
    use anyhow::{anyhow, Result};
    use security_framework_sys::authorization as sf;
    use std::ffi::CString;
    use std::io::Read;
    use std::os::raw::c_char;

    pub fn run_with_admin(cmd: &str, args: &[&str], prompt: &str) -> Result<()> {
        unsafe {
            // 1. AuthorizationCreate — get an empty authorization ref.
            let mut auth: sf::AuthorizationRef = std::ptr::null_mut();
            let st = sf::AuthorizationCreate(
                std::ptr::null(),
                std::ptr::null(),
                sf::kAuthorizationFlagDefaults,
                &mut auth,
            );
            if st != 0 {
                return Err(anyhow!("AuthorizationCreate failed: {}", st));
            }
            // RAII free
            struct AuthGuard(sf::AuthorizationRef);
            impl Drop for AuthGuard {
                fn drop(&mut self) {
                    unsafe {
                        sf::AuthorizationFree(self.0, sf::kAuthorizationFlagDefaults);
                    }
                }
            }
            let _guard = AuthGuard(auth);

            // 2. Build the AuthorizationItem requesting "system.privilege.admin"
            //    with our custom prompt text shown in the dialog.
            let right_name = CString::new("system.privilege.admin").unwrap();
            let mut right_item = sf::AuthorizationItem {
                name: right_name.as_ptr(),
                valueLength: 0,
                value: std::ptr::null_mut(),
                flags: 0,
            };
            let mut rights = sf::AuthorizationRights {
                count: 1,
                items: &mut right_item,
            };

            let prompt_name = CString::new("prompt").unwrap();
            let prompt_value = CString::new(prompt).unwrap();
            let mut env_item = sf::AuthorizationItem {
                name: prompt_name.as_ptr(),
                valueLength: prompt.len(),
                value: prompt_value.as_ptr() as *mut _,
                flags: 0,
            };
            let env = sf::AuthorizationRights {
                count: 1,
                items: &mut env_item,
            };

            // 3. AuthorizationCopyRights — shows the password dialog.
            let st = sf::AuthorizationCopyRights(
                auth,
                &rights,
                &env,
                sf::kAuthorizationFlagDefaults
                    | sf::kAuthorizationFlagInteractionAllowed
                    | sf::kAuthorizationFlagPreAuthorize
                    | sf::kAuthorizationFlagExtendRights,
                std::ptr::null_mut(),
            );
            if st != 0 {
                // -60006 = errAuthorizationCanceled (user clicked Cancel)
                // -60005 = errAuthorizationDenied (wrong password too many times)
                if st == -60006 {
                    return Err(anyhow!("Du avbrøt passordvinduet."));
                }
                return Err(anyhow!("Autorisasjonen ble avvist (kode {}).", st));
            }

            // 4. AuthorizationExecuteWithPrivileges — runs `cmd args...` as root.
            let cmd_c = CString::new(cmd).unwrap();
            let arg_cs: Vec<CString> = args.iter().map(|a| CString::new(*a).unwrap()).collect();
            let mut arg_ptrs: Vec<*const c_char> = arg_cs.iter().map(|c| c.as_ptr()).collect();
            arg_ptrs.push(std::ptr::null()); // NULL-terminated

            let mut pipe: *mut libc::FILE = std::ptr::null_mut();
            let st = AuthorizationExecuteWithPrivileges(
                auth,
                cmd_c.as_ptr(),
                sf::kAuthorizationFlagDefaults,
                arg_ptrs.as_mut_ptr() as *mut *mut c_char,
                &mut pipe,
            );
            if st != 0 {
                return Err(anyhow!(
                    "AuthorizationExecuteWithPrivileges failed: {} (the command did not run)",
                    st
                ));
            }

            // 5. Read the pipe to wait for the subprocess to finish + capture
            //    its stderr/stdout for diagnostics.
            let fd = libc::fileno(pipe);
            let mut stderr_out = String::new();
            if fd >= 0 {
                // Wrap the FILE* fd in a BufReader. We do NOT close the fd via
                // Rust's File drop because that would also fclose the FILE*.
                // After reading to EOF we explicitly fclose below.
                use std::os::unix::io::FromRawFd;
                let mut f = std::fs::File::from_raw_fd(libc::dup(fd));
                let _ = f.read_to_string(&mut stderr_out);
            }
            libc::fclose(pipe);

            // No way to get the child's exit code from
            // AuthorizationExecuteWithPrivileges. If `security` failed, it
            // typically prints to stderr — surface that.
            if !stderr_out.trim().is_empty()
                && (stderr_out.contains("error")
                    || stderr_out.contains("denied")
                    || stderr_out.contains("failed"))
            {
                return Err(anyhow!("{}", stderr_out.trim()));
            }
        }
        Ok(())
    }

    // security-framework-sys 2.x doesn't re-export this deprecated function.
    // Declare it ourselves — Apple has kept the ABI stable for 15+ years.
    unsafe extern "C" {
        fn AuthorizationExecuteWithPrivileges(
            authorization: sf::AuthorizationRef,
            pathToTool: *const c_char,
            options: sf::AuthorizationFlags,
            arguments: *mut *mut c_char,
            communicationsPipe: *mut *mut libc::FILE,
        ) -> i32;
    }
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

/// Idempotent "make the wef manifest match the bundled one" called on every
/// Spell.app startup (regardless of wizard state). Without this, users who ran
/// the wizard on an OLDER Spell.app version end up with a stale manifest in
/// Word's wef folder forever — even after upgrading. The wizard considers
/// "manifest exists" as Ready and skips, so pre-existing manifests never get
/// updated.
///
/// Compares mtimes (cheap) — if bundled is newer or sizes differ, copies. No-op
/// otherwise. Failures are logged but never block app startup.
pub fn refresh_manifest_if_stale() {
    let Ok(bundled) = bundled_manifest_path() else { return };
    let Some(wef) = word_wef_dir() else { return };
    let target = wef.join(MANIFEST_FILENAME);

    // Only refresh if the user has the add-in installed (file exists). Don't
    // create it here — that's the wizard's job (it gates on user consent).
    if !target.exists() {
        return;
    }
    if !bundled.exists() {
        return;
    }

    let bundled_meta = match fs::metadata(&bundled) { Ok(m) => m, Err(_) => return };
    let target_meta = match fs::metadata(&target) { Ok(m) => m, Err(_) => return };
    let bundled_size = bundled_meta.len();
    let target_size = target_meta.len();
    let bundled_mtime = bundled_meta.modified().ok();
    let target_mtime = target_meta.modified().ok();

    let should_refresh = bundled_size != target_size
        || match (bundled_mtime, target_mtime) {
            (Some(b), Some(t)) => b > t,
            _ => false,
        };

    if should_refresh {
        if let Err(e) = fs::copy(&bundled, &target) {
            eprintln!("refresh_manifest_if_stale: copy failed: {}", e);
        } else {
            eprintln!(
                "refresh_manifest_if_stale: updated wef manifest ({} bytes → {} bytes)",
                target_size, bundled_size
            );
        }
    }
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
