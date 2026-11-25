# Timesheet app (no app store needed)

A phone-friendly web app. Staff log in with a PIN, record start time, end time and break, and the data lands in your Google Sheet. Works on Android and iPhone.

- `index.html`, `manifest.webmanifest`, `sw.js`, `icons/` : the phone screen, hosted free on GitHub Pages.
- `apps-script/Code.gs` : the API that reads PINs and writes to the Sheet (Google Apps Script).

Hosting the screen on GitHub Pages (instead of serving it from Apps Script) avoids the "Sorry, unable to open the file" error that happens when a phone is signed in to several Google accounts.

## 1. Google Sheet + API

1. Open your Google Sheet > **Extensions > Apps Script**.
2. Replace `Code.gs` with [apps-script/Code.gs](apps-script/Code.gs). (No HTML file is needed in Apps Script.)
3. Select `setup` in the toolbar and click **Run** (approve the permissions). This creates the `Users` and `Timesheet` tabs.
4. Make yourself the first admin: in the `Users` tab add a row `Your Name | your@email | (blank PIN) | TRUE | TRUE`
   (columns: Name, Email, PIN, Active, Admin). Then open the app, tap **First time? Register** and use that exact
   name and email with the code from `Settings!B1` to choose your PIN. No default user is created.
   Delete any old sample row. Keep this sheet's sharing restricted.

### Admin screen

Anyone with `TRUE` in the **Admin** column sees **Admin: manage users** after signing in. From there an admin can:

- **Add a user** (name + email). They then register with the code and choose their own PIN.
- **Reset PIN**: clears the PIN (and any lockout). The person registers again with the same name and email and picks a new PIN, and their timesheet history stays attached. Admins never see anyone's PIN.
- **Deactivate / Activate**: blocks or restores sign-in without deleting anything.
- **Delete**: removes the login. Their rows in the `Timesheet` tab stay.
- **View or change the registration code**, or clear it to switch registration off.

Admins can't reset, deactivate or delete themselves from the screen (so the last admin can't lock themselves out). To make
someone an admin, or to change your own account, edit the `Users` tab directly. The script checks the admin flag on every
admin request, so hiding the button isn't the only protection.

### Self-registration

On first use staff tap **First time? Register**, enter their name, **email**, the **registration code**, and choose their own 6 to 8 digit PIN (change `PIN_MIN` / `PIN_MAX` at the top of `Code.gs` to adjust).

- The code is in the `Settings` tab (cell B1), created by `setup`. Admins can also see and change it in the Admin screen. Clear it to switch registration off.
- Share the code only with your team. Anyone who has the code and the link can register.
- Wrong codes are counted separately from wrong PINs: 5 per phone, 30 in total, then registration pauses for 10 minutes.
- **Forgotten PIN:** clear that person's PIN cell in `Users`. They register again with the same name and email and pick a new PIN, and their old timesheet rows stay attached.
- An employee is identified by **name + email**. Email is the unique key and is stored in lower case. An email that already has a PIN can't be registered again, and if you pre-list an email, the name typed must match the one on file. Two people can share a name if their emails differ.
- Re-running `setup` upgrades an older `Users` / `Timesheet` tab by inserting the Email column. Fill in the emails for existing people. Old timesheet rows with a blank email are no longer matched to anyone, so they won't show up in the app.
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
- **Signing in:** staff enter their email and PIN. The phone remembers the email (and keeps it after log out), so next time only the keypad shows. "Not you? Use a different email" changes it. The email is saved only after a successful sign-in or registration.
- 5 wrong PINs lock that email for 10 minutes, whatever phone is used, and also lock that phone for 10 minutes. Other staff are unaffected. Wrong email and wrong PIN show the same message.
- Times use the script's time zone (**Project Settings** in Apps Script). Check it is correct.
- The app needs internet to save. Until you update `sw.js`'s cache name, phones may show a cached older screen once after an update.
