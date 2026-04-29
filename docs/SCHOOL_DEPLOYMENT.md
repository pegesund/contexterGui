# Spell — school deployment guide (Mac + Word integration)

For IT administrators deploying Spell to many Mac pupils across a school district. If you're an individual end-user, just download Spell.dmg and run it — the rest of this document doesn't apply.

## What gets deployed

| Component | Where it lives | Required permission |
|---|---|---|
| Spell.app | `/Applications/Spell.app` | App install (your usual MDM does this) |
| Localhost TLS cert (root CA) | `/Library/Keychains/System.keychain` | Admin password (one-time) |
| Localhost cert + key | `~/Library/Application Support/Spell/word-addin-certs/` | User-writable, no admin needed |
| Word add-in manifest | `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/Spell-manifest.xml` | User-writable, no admin needed |

The cert install is the only step that needs admin privileges. Everything else runs in the user's home directory.

## Two deployment paths

Pick based on whether your pupils have admin password access on their Macs.

### Path A — pupils have admin (or Spell.app is installed once per Mac, not per user)

This is the simplest path. Each pupil runs Spell once and the wizard handles everything:

1. **Push Spell.dmg via your MDM** (Jamf, Intune, Mosyle, etc.) to each Mac.
2. Pupils launch Spell. The first-launch wizard appears once.
3. Pupils click "Installer integrasjon" and enter their Mac password when prompted.
4. Done — Word integration is live for that pupil.

The cert is per-machine (lives in System.keychain), so even if multiple pupils share the same Mac with separate user accounts, the wizard only needs to install the cert once. Subsequent pupils on the same Mac see the wizard skip the cert step automatically (it detects "already trusted") and only install the user-level pieces (manifest, leaf cert).

**You do NOT need to do anything in the M365 admin center for Path A.** The manifest is bundled in Spell.app and the wizard places it directly in each user's wef folder.

### Path B — pupils don't have admin (school-managed Macs)

If your MDM policy denies pupils admin password (typical for K-12), pupils can't trigger the cert install themselves. You have two options:

#### B1 — pre-install the cert via MDM (recommended)

1. **On a representative Mac**, install Spell normally (Path A). This creates a per-machine CA at `~/Library/Application Support/Spell/word-addin-certs/rootCA.pem`.
2. **Extract the rootCA.pem** and ship it via your MDM as a trusted root CA configuration profile (Jamf "Configuration Profiles → Certificate", Intune "macOS → Configuration → Certificate"). Set the trust to "Always Trust" for SSL.
3. **Push Spell.dmg** to the rest of the fleet.
4. Pupils launch Spell. The wizard detects the cert is already trusted (skips step 1) and just installs the user-level pieces (no admin prompt).

⚠️ **Caveat: if you reinstall Spell on the representative Mac, it generates a NEW CA**. Make sure to keep the same `rootCA.pem` you originally pushed, OR rotate it across the fleet. Treat the rootCA as a long-lived asset (10-year validity).

#### B2 — disable the wizard's cert step (v1.1, coming soon)

In a future version, Spell will detect when it's running on a managed Mac and offer a "skip cert install" option that defaults to HTTP-only mode (Word integration unavailable but everything else works). Until then, B1 is the recommended path.

## Manifest deployment via M365 admin center (optional)

For Path A, the wizard handles the manifest automatically. For Path B you may want to centrally deploy the manifest via M365 instead of shipping it via Spell.app:

1. Sign in to <https://admin.microsoft.com>
2. **Settings → Integrated apps → Upload custom apps**
3. Upload `manifest.xml` (extract from `Spell.app/Contents/Resources/word-addin/manifest.xml`)
4. Assign to your teachers/students AAD groups
5. Each pupil's Word will auto-load the add-in within ~6 hours of assignment

Once centrally deployed, Spell's wizard will detect the manifest is already in Word and skip the per-user manifest copy step.

## Verifying deployment on a pupil's Mac

After deployment, on a pupil's Mac:

```bash
# 1. Spell.app installed
ls /Applications/Spell.app

# 2. Cert is trusted by the system
security find-certificate -c "Spell Word Add-in Local CA" /Library/Keychains/System.keychain

# 3. Manifest is in Word's wef folder
ls ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/Spell-manifest.xml

# 4. Spell is running and serving HTTPS
curl -sk https://localhost:3000/errors
# → expect: []
```

If all four pass, Word integration is ready. Have the pupil restart Word and look under **Insert → My Add-ins → Spell**.

## Known limitations

- **Per-user manifests**: Currently each pupil's wef folder gets its own copy of `Spell-manifest.xml`. Centralized M365 deployment (Path B option above) is the cleaner path for large fleets.
- **No silent uninstall**: Spell doesn't yet provide an MDM-friendly uninstall that cleans up cert + manifest. To uninstall, IT can delete `/Applications/Spell.app`, the wef manifest, and the System.keychain CA entry manually.
- **macOS only**: This guide covers Mac. Windows uses an entirely different mechanism (Word COM/ActiveX) — Spell.exe just attaches automatically. No deployment configuration needed beyond pushing the .exe via your Windows-side MDM.

## Support

For deployment issues that aren't covered here, contact `support@cognio.no` with:
- Your MDM platform (Jamf / Intune / Mosyle / etc.)
- Estimated number of Macs in scope
- Whether pupils have admin password access
- Output of the four verification commands above (run on a representative pupil Mac)
