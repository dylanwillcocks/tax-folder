# Tax Folder

A phone app (a PWA) for the `Personal Documents/Tax/Tax27` folder in OneDrive.

- **Position**: your FY27 tax position, read live from the `OVERVIEW` sheet of `PTR Calculations 27.xlsx`. Net income, income and deductions, an indicative tax estimate, a breakdown by income stream, deductions by section, key dates (including a flag when a share sale has no capital gain recorded), receipts filed, and how far through the year you are.
- **To do + reminders** (on the Position tab): a checklist built from the year's dates (monthly receipt filing, quarterly updates for your accountant, the end-of-year sweep, lodgement, and a capital gain to work out when the workbook shows a sale with no gain recorded). Tick items off, and tap *Add reminders to my phone calendar* to get a 9 am alert on each due date.
- **Files**: browse folders, search, open any file in OneDrive.
- **Add receipt**: take a photo (or pick a photo/PDF), enter date, vendor and amount, and it is saved into `Receipts` (or any subfolder) as `2026-10-09 Bunnings $45.20.jpg`. Photos are shrunk to about 150 KB first.
- **Your accountant**: save their name and email in Settings and the Position tab gets an "Email this summary" button. It opens a ready-written email in your mail app; nothing is sent until you press Send.
- **Face ID lock** (Settings, the `...` button): asks for Face ID / fingerprint when you open the app or return after 2 minutes, and covers the screen in the app switcher.

It is plain HTML and JavaScript with no build step and no server. It signs in to Microsoft directly (OAuth with PKCE) and talks to your OneDrive through Microsoft Graph. Nothing is stored anywhere except in your phone's browser storage (sign-in tokens, settings, and the last figures read from the workbook).

## Files

| File | Purpose |
|---|---|
| `index.html` | Screens and styling |
| `app.js` | Sign-in, OneDrive calls, workbook reader, tax estimate, receipt upload, lock |
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
