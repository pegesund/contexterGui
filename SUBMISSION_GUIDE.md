# Spell — Chrome Web Store submission guide

This guide walks through publishing both Spell extensions to the Chrome Web Store. Hand this to whoever owns the Cognio publisher account (Petter, or whoever has the dev console login).

## What gets shipped

| # | Extension | Source | Output .zip | Audience |
|---|---|---|---|---|
| 1 | **Spell — Browser companion** (native messaging) | `Spell/contexterGui/extension/` | `Spell-Browser-Extension-<v>.zip` | Mac/Win users who already installed Spell desktop |
| 2 | **Spell for Chromebook** (standalone) | `Spell/contexter_chromebook/` | `Spell-Chromebook-<v>.zip` | Chromebook users (no desktop app possible) |

These are **two separate listings** in the Chrome Web Store dev console — they share branding ("Spell") but are technically and functionally distinct.

---

## One-time setup (Cognio account)

1. **Create or designate a Google account** dedicated to Cognio's developer publisher.
   - Use something like `dev@cognio.no`. Do NOT use a personal Gmail.
   - Whoever owns this account controls every published extension.
   - Enable 2FA — Google now requires it for new publishers.

2. **Pay the $5 one-time registration fee** at <https://chrome.google.com/webstore/devconsole>.
   - Sign in with the Cognio Google account.
   - One-time fee per publisher, not per extension.

3. **Verify business identity** as `Cognio AS`.
   - Provide legal company name, address, contact info.
   - Google may ask for business documents (tax ID, registration certificate). Reply within ~48 h.
   - Goal: listing shows "Offered by Cognio AS" + "Trader" badge instead of an individual name.
   - Timeline: ~3 business days for verification.

4. **Verify the `cognio.no` domain** (publisher verification).
   - In dev console → Account → Verified domains → Add `cognio.no`.
   - Verify via Google Search Console (TXT record on the domain or HTML file upload).
   - Effect: listings show "Offered by cognio.no" badge.

5. **Publish a privacy policy at a stable URL** before submitting.
   - Required for any extension using `nativeMessaging`, `<all_urls>`, or sensitive permissions — both of ours qualify.
   - Suggested URL: `https://cognio.no/spell/privacy`
   - It must mention: what data the extension collects, how it's used, that it isn't sold or shared, and how users can contact you.

6. **(Recommended) Create an additional support page** at `https://cognio.no/spell/support` — Web Store asks for this and it improves listing trust.

---

## Per-listing build & upload

### Listing 1 — Browser companion (native messaging)

```bash
cd Spell/contexterGui/extension
bash scripts/build-extension.sh
# → produces dist/Spell-Browser-Extension-0.1.0.zip (~44 KB)
```

In the Chrome Web Store dev console:

1. **Items → Add new item** → upload `Spell-Browser-Extension-0.1.0.zip`.
2. **Store listing** tab — copy text from `Spell/contexterGui/extension/LISTING.md`:
   - Detailed description (use the Norwegian version as primary, switch to English in the dev console for the EN translation)
   - Single purpose statement
   - Category: **Productivity**
   - Language: **Norwegian** (primary), **English** (secondary)
3. **Privacy practices** tab — answer per `LISTING.md > Privacy practices form`. Paste the per-permission justifications from `LISTING.md > Permission justifications`.
4. **Distribution** tab → **Public**. Regions: select Norway and worldwide (or limit if needed).
5. Upload **5 screenshots** (1280×800 PNG, see "Screenshots" section below).
6. Upload **promotional tile** 440×280 (use `assets/Spell-1024.png` resized).
7. Submit for review.
8. **After approval** — copy the assigned **Extension ID** from the dev console. You'll need it to wire up the desktop's native-messaging host (see "Wiring up native messaging" below).

### Listing 2 — Chromebook standalone

```bash
cd Spell/contexter_chromebook
bash scripts/build-extension.sh
# → produces dist/Spell-Chromebook-0.1.0.zip (~34 MB)
```

Same dev console flow as Listing 1, but copy text from `Spell/contexter_chromebook/LISTING.md`. Permission list is different (no `nativeMessaging`; has `sidePanel`, `offscreen`, `clipboardRead`, etc.).

If Chrome Web Store warns about duplicate names, change this listing's display name to `Spell for Chromebook` (in the dev console — does NOT require rebuilding the .zip).

---

## Wiring up native messaging (Listing 1 only)

The browser extension talks to the desktop via Chrome's native messaging protocol. The desktop's installer drops a JSON file telling Chrome which extension is allowed to connect:

| OS | Path |
|---|---|
| macOS | `~/Library/Application Support/Google/Chrome/NativeMessagingHosts/com.cognio.spell.bridge.json` |
| Windows | Registry: `HKCU\Software\Google\Chrome\NativeMessagingHosts\com.cognio.spell.bridge` (points to a JSON file in the install dir) |

The JSON file must list the **extension ID assigned by Chrome Web Store** in `allowed_origins`. Until we publish, we don't know the ID.

**Workflow:**
1. Submit extension to dev console → review approves it → copy the assigned ID (32-char alphanumeric like `ikdncjkonmegknmlpepafnmpncpfhiec`).
2. Replace the placeholder ID in `Spell/contexterGui/extension/com.norsktale.bridge.json` (rename file to `com.cognio.spell.bridge.json` for consistency).
3. Cut a new desktop release that includes the updated bridge JSON.
4. End users install desktop → desktop installer registers the bridge → extension can connect.

Until step 1 is done, the extension won't work for end users. **Don't promote the extension** before the desktop release that ships the matching bridge JSON is out.

---

## Screenshots — what to capture

Chrome requires 1280×800 PNG (or 640×400). Aim for 5.

### Browser companion
1. Google Docs with a Norwegian misspelling underlined + Spell suggestion popup
2. A regular `<textarea>` (e.g. on Reddit/Twitter) showing inline corrections
3. Spell desktop window in the corner with the browser bridge connected
4. Settings/preferences view (if any)
5. Before/after writing comparison

### Chromebook standalone
1. Side panel open showing list of detected errors
2. In-Google-Docs underline of a misspelling
3. Suggestion popup with multiple alternatives
4. First-run "downloading models" screen
5. Side-by-side: Chromebook + extension active

I can capture these once the extensions are loaded into Chrome — they're easier to take after the rebrand is complete and the UI shows "Spell" not "NorskTale". Tell me when you want to do this.

---

## Optional follow-ups

- **Edge Add-ons store** (free, Microsoft Partner Center): same .zip works. Different metadata UI but mostly identical content. Worth doing once Chrome listings are approved.
- **CI auto-publish**: After first manual publish, wire up the Chrome Web Store API in CI to upload new versions on every git tag. Needs OAuth client ID + secret + refresh token from Google Cloud Console under the same Cognio Google account. I can do this once the listings are live.
- **Group publisher** (if multiple Cognio devs need access): convert the publisher to a group publisher in the dev console settings → invite team members.

---

## Estimated timeline

| Step | Wall-clock |
|---|---|
| Cognio sets up Google account + pays $5 | 30 min |
| Business identity verification | 1–3 business days |
| Domain verification | 30 min once DNS access available |
| Privacy policy live on cognio.no | depends on Cognio web team |
| First listing review (per extension) | 1–3 business days |
| Update reviews | a few hours |

So from "Petter agrees" to "live in store": ~1–2 weeks if everything goes smoothly.

---

## What I (the dev) need from Petter

Send this to him:

> 1. Confirm dev account email (e.g. `dev@cognio.no`) and pay the $5 fee at <https://chrome.google.com/webstore/devconsole>.
> 2. Verify Cognio AS as a business publisher (he provides identity docs).
> 3. Verify cognio.no domain.
> 4. Publish a privacy policy at `https://cognio.no/spell/privacy` (content suggestion: I can draft if needed).
> 5. Either share the dev-console login with me, or invite my Google account as a co-owner.
> 
> Once those are done I'll upload both extensions, paste the listing copy, and submit for review.
