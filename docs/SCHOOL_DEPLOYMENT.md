# Spell — school deployment guide (Mac + Word integration)

For IT administrators deploying Spell to many Mac pupils across a school district. If you're an individual end-user, just download Spell.dmg and run it — the rest of this document doesn't apply.

## Short version: it's straightforward

1. **Push `Spell.dmg` to pupils' Macs via your MDM** (Jamf, Intune, Mosyle, Workspace ONE, etc.). Standard `.app` deployment, no special configuration.
2. **Pupils launch Spell.** A wizard appears, they click "Installer integrasjon", and Word integration is set up automatically.

That's it. The wizard does NOT require admin password. Everything happens in the pupil's user account, including:
- The TLS certificate (installed in the pupil's login keychain — no system-wide changes)
- The Word add-in manifest (placed in the pupil's Word add-in folder)

This means the wizard works identically on:
- Personal Macs where the user has admin
- School-managed Macs where pupils don't have admin
- Multi-user Macs shared by several pupils (each user account gets its own setup)

## What gets installed where

| Component | Where it lives | Privilege required |
|---|---|---|
| Spell.app | `/Applications/Spell.app` | App install (your MDM does this) |
| TLS root CA | `~/Library/Keychains/login.keychain-db` (per-user) | None — login keychain is auto-unlocked at login |
| TLS leaf cert + key | `~/Library/Application Support/Spell/word-addin-certs/` | User-writable, no admin |
| Word add-in manifest | `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/Spell-manifest.xml` | User-writable, no admin |

macOS's TLS validation consults both `System.keychain` (admin domain) and `login.keychain-db` (user domain) by default, so a per-user root CA is sufficient for Word to trust Spell's local HTTPS server.

## What pupils see

1. Launch Spell from `/Applications`
2. **Language picker** → pick Bokmål
3. **Word integration wizard** → "Installer integrasjon" button
4. ~1 second — a green ✓ appears with "Ferdig! Start Microsoft Word på nytt."
5. Click "Lukk", main app starts

No password prompts. No terminal. No setup steps a pupil could fail to do.

## Optional: centrally deploy the manifest via M365 admin center

For larger fleets you may prefer to deploy the Word add-in manifest centrally rather than relying on each pupil's wizard to drop it. The wizard detects this state and skips the manifest copy step.

1. Sign in to <https://admin.microsoft.com>
2. **Settings → Integrated apps → Upload custom apps**
3. Upload the manifest from a representative install: `/Applications/Spell.app/Contents/Resources/word-addin/manifest.xml`
4. Assign to your teacher/student AAD groups
5. Each pupil's Word will auto-load the add-in within ~6 hours of assignment

Centralized deployment via M365 doesn't replace the per-pupil cert install (the cert IS per-machine because it's a localhost cert). The wizard will still run on each Mac to install the cert, but the manifest step will be skipped.

## Verifying deployment on a pupil's Mac

```bash
# 1. Spell.app installed
ls /Applications/Spell.app

# 2. Cert is trusted in login keychain
security find-certificate -c "Spell Word Add-in Local CA" \
  ~/Library/Keychains/login.keychain-db

# 3. Manifest is in Word's wef folder
ls ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/Spell-manifest.xml

# 4. Spell is running and serving HTTPS
curl -sk https://localhost:3000/errors
# → expect: []
```

If all four pass, Word integration is ready. Have the pupil restart Word and look under **Insert → My Add-ins → Spell**.

## What if I want pupils to have ZERO setup steps

Two options for the wizard auto-running silently:

**Option 1 — pre-configure the dismissed flag.** Have your MDM drop the following file before the pupil's first launch:

```json
// ~/Library/Application Support/Spell/settings.json
{ "language": "nb", "word_addin_wizard_dismissed": false }
```

This pre-selects Bokmål and lets the wizard run normally on first launch. The pupil still sees the wizard but it's a single click.

**Option 2 — pre-deploy the cert + manifest via your MDM.** Push the per-machine cert files into each pupil's home directory at MDM provisioning time (`~/Library/Application Support/Spell/word-addin-certs/`) along with a manifest in their `wef/` folder. The wizard will detect "Ready" status and skip silently.

This option is more complex and requires generating a unique cert per pupil OR sharing a cert across all pupils (the latter is technically a single point of compromise but acceptable in a school environment if the cert is short-lived and managed).

## Uninstall

Spell does not yet have an automatic uninstaller. To clean up everything an MDM-pushed install left behind:

```bash
sudo rm -rf /Applications/Spell.app
rm -rf ~/Library/Application\ Support/Spell
rm -f ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/Spell-manifest.xml
security delete-certificate -c "Spell Word Add-in Local CA" \
  ~/Library/Keychains/login.keychain-db
```

A scriptable uninstaller is on the v1.1 roadmap.

## macOS only

Windows uses a completely different mechanism (Word COM/ActiveX) — Spell.exe attaches to Word automatically when both are running. No deployment configuration needed beyond pushing the .exe via your Windows-side MDM.

## Support

For deployment issues that aren't covered here, contact `support@cognio.no` with:
- Your MDM platform (Jamf / Intune / Mosyle / etc.)
- Estimated number of Macs in scope
- Output of the four verification commands above (run on a representative pupil Mac)
- Output of `security find-identity -v -p ssl-server` (to confirm the per-user cert is properly stored)
