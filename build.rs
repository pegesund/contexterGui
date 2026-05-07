fn main() {
    #[cfg(target_os = "macos")]
    {
        println!("cargo:rustc-link-lib=framework=Speech");
        println!("cargo:rustc-link-lib=framework=AppKit");
        println!("cargo:rustc-link-lib=framework=AVFoundation");
    }

    // Embed assets/Spell.ico as a Win32 icon resource so Explorer, the
    // taskbar, and Alt+Tab show our brand icon for acatts-rust.exe instead
    // of the generic blank-page placeholder Windows uses for icon-less
    // executables. Velopack picks the same icon up when generating
    // Setup.exe (matches Concentrate's ApplicationIcon flow on .NET).
    #[cfg(target_os = "windows")]
    {
        let mut res = winresource::WindowsResource::new();
        res.set_icon("assets/Spell.ico");
        if let Err(e) = res.compile() {
            // Don't fail the build if rc.exe isn't available locally — just
            // warn. CI runners always have it; dev WSL/cross builds may not.
            println!("cargo:warning=winresource: failed to embed icon: {}", e);
        }
        println!("cargo:rerun-if-changed=assets/Spell.ico");
    }
}
