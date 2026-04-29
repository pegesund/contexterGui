# Spell — school deployment guide (Mac + Word integration)

For IT administrators deploying Spell to many Mac pupils across a school district. If you're an individual end-user, just download Spell.dmg and run it — the rest of this document doesn't apply.

## What gets installed where

| Component | Where it lives | Privilege required |
|---|---|---|
| `Spell.app` | `/Applications/Spell.app` | Standard `.app` install (your MDM does this) |
| TLS root CA | `/Library/Keychains/System.keychain` (system-wide) | **Admin password** — Office add-in WKWebView only honors system-domain trust anchors. Login-keychain trust is not enough. |
| TLS leaf cert + key | `~/Library/Application Support/Spell/word-addin-certs/` | User-writable, no admin |
| Word add-in manifest | `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/Spell-manifest.xml` | User-writable, no admin |

The cert install is the only step that needs admin. Everything else runs in the pupil's home directory.

## Two deployment paths

### Path A — pupils have admin (or you don't mind one password prompt per Mac)

1. Push `Spell.dmg` via your MDM to each Mac (Jamf, Intune, Mosyle, Workspace ONE, etc.).
2. Pupils launch Spell. The first-launch wizard appears.
3. Wizard prompts for admin password once (native macOS "security" dialog with friendly message: *"Spell needs your password to install a security certificate so Microsoft Word can trust the local connection."*).
4. After password is entered, the wizard finishes silently. Word integration is live.

The cert is per-machine and lives in System.keychain, so subsequent users on the same Mac (multi-user setup) won't see the password prompt — the wizard detects "Ready" and skips.

### Path B — pupils don't have admin (typical K-12 school-managed Macs)

If your MDM policy denies pupils admin password, they can't trigger the cert install themselves. Pre-install the cert via your MDM:

1. **On a representative Mac**, install Spell normally (Path A). This generates a per-machine CA at `~/Library/Application Support/Spell/word-addin-certs/rootCA.pem`.
2. **Extract `rootCA.pem`** and ship it via your MDM as a trusted root CA configuration profile:
   - **Jamf**: Configuration Profiles → Certificate → upload PEM, set Allowed Apps "All", Trust → Always Trust SSL
   - **Intune**: Devices → Configuration → macOS → Trusted certificate → upload PEM
   - **Mosyle**: Profile → Certificates → SSL Trust → Always Trust
3. Push `Spell.dmg` to the rest of the fleet.
4. Pupils launch Spell. The wizard detects the cert is already trusted (skips the cert step) and installs only the user-level pieces (manifest). **No password prompt.**

⚠️ **Caveat: if you reinstall Spell on the representative Mac, it generates a NEW CA**. Keep the same `rootCA.pem` you originally pushed, OR re-extract and re-push. Treat the rootCA as a long-lived asset (10-year validity baked in).

## Optional: centrally deploy the manifest via M365 admin center

For both paths above, the wizard handles the manifest automatically. For larger fleets you can deploy the manifest centrally instead:

1. Sign in to <https://admin.microsoft.com>
2. **Settings → Integrated apps → Upload custom apps**
3. Upload the manifest from a representative install: `/Applications/Spell.app/Contents/Resources/word-addin/manifest.xml`
4. Assign to your teacher/student AAD groups
5. Each pupil's Word will auto-load the add-in within ~6 hours of assignment

Once centrally deployed, Spell's wizard will detect the manifest is already present in Word and skip the per-user manifest copy step.

## Verifying deployment on a pupil's Mac

```bash
# 1. Spell.app installed
ls /Applications/Spell.app

# 2. Cert is trusted in System.keychain (NOT login.keychain)
security find-certificate -c "Spell Word Add-in Local CA" /Library/Keychains/System.keychain

# 3. Manifest is in Word's wef folder
ls ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/Spell-manifest.xml

# 4. Spell is running and serving HTTPS
curl -sk https://localhost:3000/errors
# → expect: []
```

If all four pass, Word integration is ready. Have the pupil restart Word and look under **Insert → My Add-ins → Spell**.

## Uninstall

Spell does not yet have a one-shot uninstaller. To remove cleanly:

```bash
sudo rm -rf /Applications/Spell.app
rm -rf ~/Library/Application\ Support/Spell
rm -f ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/Spell-manifest.xml
sudo security delete-certificate -c "Spell Word Add-in Local CA" \
  /Library/Keychains/System.keychain
```

A scriptable uninstaller is on the v1.1 roadmap.

## Why Office requires System.keychain trust

Word for Mac uses WKWebView (Apple's modern web view) for all add-in task panes. WKWebView's TLS validation only consults trust anchors in the admin (system) domain — login-keychain user-domain trust is silently ignored. We tested an earlier wizard version that installed in login.keychain (no password prompt) and Word rejected the connection with *"Add-in Error: The content is blocked because it isn't signed by a valid security certificate."*

This is a Word-specific constraint. Other apps (Safari, Chrome, curl, custom Rust apps) honor user-domain trust fine.

## macOS only

Windows uses a completely different mechanism (Word COM/ActiveX) — Spell.exe attaches to Word automatically when both are running. No deployment configuration needed beyond pushing the .exe via your Windows-side MDM. No certs, no manifests, no admin password.

## Support

For deployment issues that aren't covered here, contact `support@cognio.no` with:
- Your MDM platform (Jamf / Intune / Mosyle / etc.)
- Estimated number of Macs in scope
- Whether pupils have admin password access
- Output of the four verification commands above (run on a representative pupil Mac)
