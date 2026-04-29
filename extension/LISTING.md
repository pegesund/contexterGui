# Chrome Web Store listing copy — Spell (browser extension)

This is the companion browser extension that requires the Spell desktop app to be installed (Mac/Windows). It uses native messaging to send text to the desktop for analysis.

---

## Item details

- **Name**: `Spell`
- **Category**: `Productivity`
- **Language**: `Norwegian` (primary), `English` (secondary)

---

## Short description (max 132 chars)

### Norwegian
> Norsk staving og grammatikksjekk i nettleseren — virker i Google Dokumenter, tekstfelt og redigerbart innhold.

### English
> Norwegian spelling and grammar checking in your browser — works in Google Docs, text fields, and editable content.

---

## Detailed description (max 16,000 chars — keep it concise)

### Norwegian

```
Spell hjelper elever og voksne med dysleksi å skrive korrekt norsk i nettleseren. Den ser etter stavefeil og grammatikkfeil mens du skriver i Google Dokumenter, tekstfelt på nettsider, og alle slags redigerbare elementer.

✦ FUNKSJONER
• Stavekontroll for bokmål med over 633.000 oppslagsord
• Grammatikksjekk basert på morfologisk analyse
• Forslag til riktig ord når du staver feil
• Virker overalt: Google Dokumenter, e-post, sosiale medier, formularer
• Designet for personer med dysleksi — tydelige understrek, lette forslag

✦ KREVER SPELL DESKTOP
Denne utvidelsen krever at Spell-appen er installert på din Mac eller Windows-PC. Last ned fra https://cognio.no/spell

Utvidelsen sender tekst lokalt til din egen datamaskin for analyse. Ingenting sendes til skyen — alt skjer på din maskin.

✦ FOR HVEM?
• Elever i ungdomsskolen og videregående
• Studenter
• Voksne med dysleksi eller skrivevansker
• Skriveassistanse på arbeidsplassen

✦ STØTTE
Har du spørsmål? Send e-post til support@cognio.no eller besøk https://cognio.no/spell
```

### English

```
Spell helps pupils and adults with dyslexia write correct Norwegian in the browser. It checks spelling and grammar as you type in Google Docs, text fields on websites, and any editable content.

✦ FEATURES
• Bokmål spell-check with 633,000+ entries
• Grammar checking based on morphological analysis
• Word suggestions when you mistype
• Works everywhere: Google Docs, email, social media, forms
• Designed for dyslexic users — clear underlines, simple suggestions

✦ REQUIRES SPELL DESKTOP
This extension requires the Spell desktop app installed on your Mac or Windows PC. Download from https://cognio.no/spell

The extension sends text locally to your own computer for analysis. Nothing is sent to the cloud — all processing happens on your machine.

✦ WHO IS IT FOR?
• Lower- and upper-secondary pupils
• University students
• Adults with dyslexia or writing difficulties
• Workplace writing support

✦ SUPPORT
Questions? Email support@cognio.no or visit https://cognio.no/spell
```

---

## Single purpose

> Spelling and grammar checking for Norwegian (Bokmål) text inside the browser, by relaying text fields to the locally-installed Spell desktop application via native messaging.

---

## Permission justifications

Chrome Web Store requires a one-line justification per permission. Reviewers read these.

| Permission | Justification |
|---|---|
| `nativeMessaging` | The extension relays text to the locally-installed Spell desktop application for spelling and grammar analysis. No remote server is used; analysis happens on the user's own computer. |
| `alarms` | Used to schedule periodic re-checks of editable fields and clean up stale UI overlays after the user stops editing. |
| `<all_urls>` (host permission via `content_scripts`) | The extension provides spelling support in any text field on any website (Google Docs, email, social media, etc.). The text never leaves the user's computer. |

---

## Privacy practices form

When prompted by the dev console:

- **Single purpose**: Spelling/grammar checking for Norwegian text, via the local Spell desktop app.
- **What data does the extension collect?** The extension collects the contents of editable text fields the user is actively typing in. This data is sent only to the Spell desktop app running on `localhost` on the user's own computer — never to a remote server. No personal data, no analytics, no telemetry.
- **Data usage**: ☑ The data is used for the single purpose of providing spelling/grammar suggestions.
- **Data sale**: ☐ NO — data is never sold or transferred to third parties.
- **Data sharing for non-essential purposes**: ☐ NO.
- **Privacy policy URL**: `https://cognio.no/spell/privacy` (Cognio must publish this — see SUBMISSION_GUIDE.md)
