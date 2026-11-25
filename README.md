# Timesheet app (no app store needed)

A phone-friendly web app. Staff log in with a PIN, record start time, end time and break, and the data lands in your Google Sheet. Works on Android and iPhone.

- `index.html`, `manifest.webmanifest`, `sw.js`, `icons/` : the phone screen, hosted free on GitHub Pages.
- `apps-script/Code.gs` : the API that reads PINs and writes to the Sheet (Google Apps Script).

Hosting the screen on GitHub Pages (instead of serving it from Apps Script) avoids the "Sorry, unable to open the file" error that happens when a phone is signed in to several Google accounts.

## 1. Google Sheet + API

1. Open your Google Sheet > **Extensions > Apps Script**.
2. Replace `Code.gs` with [apps-script/Code.gs](apps-script/Code.gs). (No HTML file is needed in Apps Script.)
3. Select `setup` in the toolbar and click **Run** (approve the permissions). This creates the `Users` and `Timesheet` tabs.
4. In `Users`, replace the sample row with one row per person: `Name | PIN | TRUE`. PINs must be unique; `FALSE` disables someone. Keep this sheet's sharing restricted.
5. **Deploy > New deployment > Web app**: Execute as **Me**, Who has access **Anyone**. Copy the URL (ends in `/exec`).

## 2. Put the URL in the screen

In [index.html](index.html), replace `PASTE_WEB_APP_URL_HERE` with the `/exec` URL.

## 3. Publish on GitHub Pages

1. Create a new repository on GitHub (e.g. `timesheet`) under riteshroy85.
2. Upload everything in this folder, or from this folder run:
   ```
   git init
   git add .
   git commit -m "Timesheet app"
   git branch -M main
   git remote add origin https://github.com/riteshroy85/timesheet.git
   git push -u origin main
   ```
3. On GitHub: **Settings > Pages > Source: Deploy from a branch > main / (root) > Save**.
4. After a minute the app is live at `https://riteshroy85.github.io/timesheet/`. Send that link to staff.

Note: a free GitHub account needs the repository to be **public** for Pages. The repo holds no secrets (PINs live in the Sheet), but the Apps Script URL will be visible in `index.html`. That is why PINs plus lockout matter.

## 4. Install on phones

- **iPhone (Safari only):** Share > Add to Home Screen
- **Android (Chrome):** menu > Install app / Add to Home screen

## Notes

- After editing `Code.gs`, use **Deploy > Manage deployments > Edit > New version** (the URL stays the same).
- The form takes Date, Start time, End time and total Break in hours (e.g. 0.75). Saving again for the same date overwrites that day's row.
- Total Hrs = End - Start, In-Hrs = Total Hrs - Break.
- 5 wrong PINs lock a device for 10 minutes; 30 wrong PINs across all devices lock everyone for 10 minutes.
- Times use the script's time zone (**Project Settings** in Apps Script). Check it is correct.
- The app needs internet to save. Until you update `sw.js`'s cache name, phones may show a cached older screen once after an update.
