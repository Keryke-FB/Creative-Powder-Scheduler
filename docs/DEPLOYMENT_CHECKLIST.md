# Deployment Checklist

Use this checklist to turn the local V2 files into the live browser dashboard.

## One-Time Google Setup

1. Open Google Apps Script and create a new project.
2. Add these local files to the Apps Script project:
   - `apps-script/Code.gs`
   - `apps-script/Dashboard.html`
   - `apps-script/appsscript.json`
3. In Apps Script, enable the Advanced Google Service named `Drive`.
4. Open the linked Google Cloud project and enable the Google Drive API.
5. Add the required Script properties listed in `docs/SCRIPT_PROPERTIES.example.md`.
6. In Apps Script, run `getDashboardState()` once and approve permissions.
7. Run `installTriggers()` once.
8. Deploy as a web app.

## Optional Clasp Setup

The real `.clasp.json` file is intentionally ignored because it contains the
deployment's Apps Script project id. To work locally with clasp:

1. Copy `.clasp.example.json` to `.clasp.json`.
2. Replace `YOUR_APPS_SCRIPT_PROJECT_ID` with the Apps Script project id.
3. Run `clasp push` from the repository root.

## Recommended Web App Settings

- Execute as: `Me`
- Access: company users or specific allowed users

The dashboard runs as the deploying user so supervisors do not need direct script/API setup. Keep the Google Sheet and photo folder shared with the people who need normal visibility.

## Verification

After deployment:

1. Open the web app URL.
2. Confirm the badge under the title says `Connected to Google Apps Script`.
3. Confirm the four seed rows appear.
4. Change one row to `Run Tomorrow`.
5. Confirm the `Floor Schedule` tab updates.
6. Press `Complete / Send to Shop`.
7. Confirm the printable shop handoff opens with rows populated.
8. Confirm the selected row receives a `Sent to Shop Timestamp` and `Shop Batch`.
9. Drop one test image into the traveler photo folder.
10. Press `Digest Photos Now`.
11. Confirm a row appears in `Master Capture`.
12. Confirm the image moves into `Daily Traveler Photos/Processed`.
13. Drop one test Google Sheet version of the open-order report into `Open Order Reports`.
14. Press `Import Open Orders`.
15. Confirm report rows appear in `Master Capture`.
16. Confirm the report sheet moves into `Open Order Reports/Processed Reports`.

If the badge says `Preview mode: not connected to Google`, the dashboard is being opened as a local HTML preview rather than as the deployed Apps Script web app. In preview mode, the page can show sample rows and the print surface, but it cannot write to Google Sheets or digest Drive photos.

## Trigger Behavior

- Every 5 minutes: checks for new images and digests them.
- Every day around 3:00 PM Indiana time: runs a safety sweep before scheduling.
- Every day around 3:05 PM Indiana time: imports any new open-order reports.
- Manual button: `Digest Photos Now` runs immediately from the dashboard.
- Manual button: `Import Open Orders` imports report files immediately from the dashboard.

## Current Limits

The first reader uses Google Drive OCR and a conservative parser. It will mark rows as `Needs Review` when fields are missing or low-confidence. That is intentional: uncertain data should be visible and fixable, not silently guessed.

The likely V2.1 upgrade is adding Gemini or another vision model for more accurate field extraction from the traveler images.
