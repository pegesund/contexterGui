# Spell

Desktop app and release tooling for Spell.

## Build And Release

### Release feeds

Spell now supports multiple binary repos/feeds.

- Default feed: `pegesund/spell_binaries`
- Municipality-specific feed example: `pegesund/spell_binaries_nittedal`

The selected feed is controlled by `SPELL_RELEASES_REPO`.

- If unset, builds and installed apps fall back to `pegesund/spell_binaries`
- If set during build/release, the packaged app will check that repo for updates

### Release channels

Spell also supports explicit Velopack channels for auto-update targeting.

- Default Windows channel: `win`
- Example municipality auto-update channel: `win-nittedal-auto`

The selected channel is controlled by `SPELL_RELEASES_CHANNEL`.

- If unset, packaged apps use the channel baked into the Velopack package manifest
- If set during build/release, the packaged app explicitly checks that channel
- This is the right way to mark municipality-specific releases as auto-update releases

### Token handling

Do not hardcode GitHub personal access tokens in code or docs.

- Local release work should use an environment variable such as `GH_TOKEN`
- Windows CI uploads use the GitHub Actions secret `RELEASES_REPO_TOKEN`
- The token used for release uploads must have access to every binary repo we publish to

### Mac release

Publish to the default feed:

```bash
bash scripts/release-mac.sh 0.1.0
```

Publish to a dedicated feed/channel:

```bash
SPELL_RELEASES_REPO=pegesund/spell_binaries_nittedal SPELL_RELEASES_CHANNEL=osx-arm64-nittedal-auto bash scripts/release-mac.sh 0.1.0 --no-tag-push
```

Note:
- Non-default feeds should use `--no-tag-push`
- Then trigger Windows release manually with matching repo/channel values

### Windows CI release

The `Release (Windows)` workflow supports a manual `releases_repo` input.

- `version`: semver like `0.1.0`
- `releases_repo`: repo to receive release artifacts, for example `pegesund/spell_binaries_nittedal`
- `release_channel`: Velopack channel to publish, for example `win-nittedal-auto`

If `releases_repo` is not provided, the workflow uploads to `pegesund/spell_binaries`.
If `release_channel` is not provided, the workflow publishes on `win`.

### Local Windows packaging

Default feed:

```powershell
pwsh -File scripts/build-windows-local.ps1 -Version 0.1.0
```

Dedicated feed:

```powershell
pwsh -File scripts/build-windows-local.ps1 -Version 0.1.0 -ReleasesRepo pegesund/spell_binaries_nittedal -ReleaseChannel win-nittedal-auto
```

### Notes

- Keep normal customer releases on the default feed unless a deployment is meant to be isolated
- Use a dedicated binary repo/feed for municipality-specific auto-upgrade control
- Use a dedicated release channel when a municipality release should be explicitly treated as an auto-update release
