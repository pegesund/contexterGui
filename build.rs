fn main() {
    #[cfg(target_os = "macos")]
    {
        println!("cargo:rustc-link-lib=framework=Speech");
        println!("cargo:rustc-link-lib=framework=AppKit");
        println!("cargo:rustc-link-lib=framework=AVFoundation");
    }

    // Embed the Spell identity in the Win32 version resource. Windows uses
    // FileDescription/ProductName in Explorer, Task Manager, error dialogs,
    // taskbar grouping, and other shell UI, so do not let it fall back to the
    // Rust target metadata.
    #[cfg(target_os = "windows")]
    {
        let mut res = winresource::WindowsResource::new();
        res.set_icon("assets/Spell.ico")
            .set("FileDescription", "Spell")
            .set("ProductName", "Spell")
            .set("InternalName", "Spell")
            .set("OriginalFilename", "Spell.exe")
            .set("CompanyName", "Cognio AS")
            .set("LegalCopyright", "Copyright (C) Cognio AS");
        if let Err(e) = res.compile() {
            // Don't fail the build if rc.exe isn't available locally - just
            // warn. CI runners always have it; dev WSL/cross builds may not.
            println!("cargo:warning=winresource: failed to embed icon: {}", e);
        }
        println!("cargo:rerun-if-changed=assets/Spell.ico");
    }
}
