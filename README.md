# RMWA Household & Family Survey Portal

Google Apps Script web application for the RMWA household and family survey.

## Files
- `Code.gs` — Apps Script backend
- `Index.html` — public survey interface
- `Admin.html` — administrator portal

## Important security notes
- Keep the linked Google Spreadsheet private.
- Do not commit household survey records, exported reports, passwords, session tokens, or private configuration values.
- Configure secrets and spreadsheet IDs in **Apps Script → Project Settings → Script Properties**, not in source code.
- This repository contains a sanitized source copy. It does not contain your live spreadsheet ID or an initial administrator password.

## Apps Script setup
1. Create or open your Google Apps Script project.
2. Add the backend as `Code.gs` and the HTML files with the exact names `Index` and `Admin`.
3. In **Project Settings → Script Properties**, set:
   - `RMWA_INITIAL_ADMIN_PASSWORD` — a strong initial password of at least 12 characters (needed only when creating the first admin account).
   - `RMWA_SPREADSHEET_ID` — your existing private spreadsheet ID, if connecting to an existing database. If omitted on first setup, the app creates a new private spreadsheet and stores its ID automatically.
4. Run the appropriate setup function only after checking the target spreadsheet and script properties.
5. Deploy as a web app using the access settings appropriate for your survey.

## Runtime note
This application depends on Google Apps Script HtmlService and `google.script.run`. GitHub Pages cannot run the app by itself. Google Apps Script remains the application runtime and database connection.
