# Tax Folder

A phone app (a PWA) for the `Personal Documents/Tax/Tax27` folder in OneDrive.

- **Position**: your FY27 tax position, read live from the `OVERVIEW` sheet of `PTR Calculations 27.xlsx`. Net income, income and deductions, an indicative tax estimate, a breakdown by income stream, deductions by section, key dates (including a flag when a share sale has no capital gain recorded), receipts filed, and how far through the year you are.
- **To do + reminders** (on the Position tab): a checklist built from the year's dates (monthly receipt filing, quarterly updates for your accountant, the end-of-year sweep, lodgement, and a capital gain to work out when the workbook shows a sale with no gain recorded). Tick items off, and tap *Add reminders to my phone calendar* to get a 9 am alert on each due date.
- **Add receipt, read by Claude**: take a photo, scan, or choose a file. With a Claude API key in Settings, Claude reads the document, fills in the vendor, date and amount, guesses what it is for (claim it, company cost, part of the capital gain, personal) and asks a question where it is not sure. Say yes or edit it, and the app files the document in the right folder under Receipts (Company costs, Capital gains, Not claimable, or Receipts itself) and records your decision in the Inbox, ready to be logged in the workbook.
- **Position** adds up each income stream's own sheet in the workbook, so the figures stay correct even when the OVERVIEW sheet's saved totals are out of date.
- **Inbox**: receipts Claude finds in your Gmail, each with a guess at what it is for and a question where it is not obvious. Answer "Yes, that's right" or "Something else", add a note, and the decision is saved to `Inbox/inbox.json` in your OneDrive. Treatments: claim it, company cost, claimed back from Oakwood, personal, part of the capital gain, not a receipt. A weekly scheduled Claude task adds new receipts (see `Inbox/SYNC-PROCEDURE.md`). Gmail will not hand over attachments, so each item links to the email.
- **Ask**: ask whether you can claim something. Claude answers using your profile (`Inbox/profile.md`), your workbook numbers, your receipt decisions and the current tax news. With a Claude API key in Settings it answers instantly on the phone; without one, questions are saved to `Inbox/questions.json` and answered the next time Claude runs. Answers are general information, not tax advice.
- **For you** (inside Ask): `Inbox/briefing.json`, tax news (Budget, ATO, Queensland land tax and more) with how each item could affect you and questions to ask.
- **Files**: browse folders, search, open any file in OneDrive.
- **Scan a document** (button at the top of Add receipt): take a photo of a page, and the app finds its edges (drag the orange corners if any are off), straightens it, and cleans it up (Original, Clean, or Black & white). Add as many pages as you like and save one PDF, or a single page as a JPEG. All of it happens on the phone. On an iPhone you can also use the built-in scanner: *Choose a photo or PDF*, then *Choose File*, then the `...` menu, then *Scan Documents*.
- **Add receipt**: take a photo (or pick a photo/PDF), enter date, vendor and amount, and it is saved into `Receipts` (or any subfolder) as `2026-10-09 Bunnings $45.20.jpg`. Photos are shrunk to about 150 KB first.
- **Your accountant**: save their name and email in Settings and the Position tab gets an "Email this summary" button. It opens a ready-written email in your mail app; nothing is sent until you press Send.
- **Face ID lock** (Settings, the `...` button): asks for Face ID / fingerprint when you open the app or return after 2 minutes, and covers the screen in the app switcher.

It is plain HTML and JavaScript with no build step and no server. It signs in to Microsoft directly (OAuth with PKCE) and talks to your OneDrive through Microsoft Graph. Nothing is stored anywhere except in your phone's browser storage (sign-in tokens, settings, and the last figures read from the workbook).

## Files

| File | Purpose |
|---|---|
| `index.html` | Screens and styling |
| `app.js` | Sign-in, OneDrive calls, workbook reader, tax estimate, receipt upload, lock |
| `scan.js` | Document scanner: page finder, perspective fix, clean-up filters, PDF writer |
| `xlsx.min.js` | SheetJS 0.18.5 (Apache-2.0), reads the workbook in the browser |
| `sw.js` | Lets the app open with no signal |
| `manifest.webmanifest`, `*.png` | Home-screen name and orange `$` icons |

## One-time setup

1. **Host the folder over HTTPS** (sign-in and Face ID will not work from a plain `http://` address except on `localhost`). It is published with GitHub Pages at `https://dylanwillcocks.github.io/tax-folder/` (repo `dylanwillcocks/tax-folder`).
2. **Register the app with Microsoft**: portal.azure.com, Microsoft Entra ID, App registrations, New registration.
   - Supported account types: *Accounts in any organizational directory and personal Microsoft accounts*.
   - Redirect URI: platform **Single-page application (SPA)**, set to your hosted URL **including the trailing slash**. The app shows the exact value on its first screen.
   - API permissions: Microsoft Graph, Delegated, **Files.ReadWrite**.
   - Copy the **Application (client) ID**.
3. **Open the hosted URL on your phone**, paste the client ID, and sign in.
4. **Add to Home Screen**: iPhone: Safari, Share, Add to Home Screen. Android: Chrome menu, Install app. Then open it from the home-screen icon.
5. **Turn on the lock**: `...` button, Settings, Face ID lock, Turn on.

## How the tax position is worked out

- The workbook's income streams are columns B to G of `OVERVIEW`; the sections are `Income`, `PTR Deductions`, `CTR Deductions`, `Investment Property Deductions`, `Other` and `Key Dates`. Any row with a number in it is picked up, so new rows and new amounts appear without changing the app. Keep the section headings as they are.
- Net income = income minus PTR, CTR and investment property deductions (the same as the workbook's own Net Income line).
- The tax estimate uses resident rates (15% first bracket for 2026-27, 16% for 2025-26, 14% from 2027-28) plus the 2% Medicare levy, applied to all net income as if taxed to you personally, with no offsets, PAYG credits or capital gains. It is a guide only; trust distributions can change it a lot.
- The workbook is re-read only when OneDrive says it has changed, or when you tap Refresh.

## Notes

- Sign in from the **installed** home-screen app on iPhone, not from Safari: iOS keeps their storage separate. Set up Face ID there too.
- Microsoft expires the sign-in about every 24 hours for apps like this, so expect to tap "Sign in" again most days. It is one tap and no password.
- The lock is a gate on the app, not encryption. Your phone's own passcode is what protects the data stored in the browser.
- "Can't unlock? Sign out" on the lock screen removes the lock, the sign-in and the stored figures, so nothing is exposed.
- Changing the folder, workbook name, receipts folder or client ID: `...` button, Settings.
- The app can see your whole OneDrive (that is how Microsoft's permission works), but it only ever looks inside the folder set in Settings.

## Privacy of the Inbox, profile and Ask

- The app itself is a public page with no personal data in it. Everything personal (`inbox.json`, `profile.md`, `questions.json`, `briefing.json`) is in your private OneDrive and is read after you sign in.
- With an API key saved, each question sends your profile, workbook figures and receipt decisions to Claude (Anthropic) to answer. The key lives only in this phone's browser storage. Create a dedicated key in the Claude Console and give it a spending limit.
- Signing out removes the key, the saved conversation, the lock and the stored figures from the phone.

## If something looks out of date

Settings (the `...` button) shows the app version at the bottom. The app now re-checks its files on every open and reloads itself when a new version arrives. If a button does nothing, close the app completely and open it again.
