//! First-launch / Settings-triggered wizards for installing Spell's optional
//! integrations into the user's environment (currently: Word for macOS).
//!
//! Designed so the same Spell.app DMG works for:
//!   - Individual users (pupil/parent runs the wizard at first launch)
//!   - School deployments (IT admin pre-installs cert + manifest centrally;
//!     the wizard detects everything is in place and skips)

#[cfg(target_os = "macos")]
pub mod word_addin_setup;
