# Creative Coatings Powder Schedule V2

This project is the browser-dashboard and automatic intake layer for the Creative Coatings powder schedule.

## What It Does

- Reads traveler photos from the Google Drive `Daily Traveler Photos` folder.
- Imports Plate Shop open-order reports from native Google Sheets placed in the Google Drive `Open Order Reports` folder.
- Uses Gemini Vision Flash for traveler extraction when a Gemini API key is configured.
- Falls back to Google Drive OCR if the Gemini key is missing or a vision call fails.
- Parses schedule fields into the `Master Capture` Google Sheet.
- Keeps the `Floor Schedule` tab generated from `Run Tomorrow` and `Hot List`.
- Moves captured photos into a `Processed` subfolder.
- Provides a browser dashboard with:
  - `Import Open Orders`
  - `Digest Photos Now`
  - status updates for `On Hold`, `Run Tomorrow`, and `Hot List`
  - `Complete / Send to Shop` for the printable supervisor handoff
  - simple review filtering for low-confidence photo reads
  - a visible reader badge showing whether Gemini Vision or Drive OCR fallback is active

The staff-facing browser table is intentionally narrow:

- Status
- Customer
- Part Number
- Quantity
- Powder
- Masked

Powder and masked values are filled by matching the schedule part number against the configured powder master and masked parts spreadsheets.

The shop handoff sorts by paint color, then hot priority, then customer so supervisors can keep the line in one color as long as possible while still seeing urgent jobs first.

## V2 Report Intake

The normal intake path is now the open-order report dump:

1. Convert the current Plate Shop open-order export into a Google Sheet and place that Google Sheet in `Creative Coatings Powder Schedule / Open Order Reports`.
2. Press `Import Open Orders`, or let the 3:05 PM Indiana sweep pick them up.
3. Imported report sheets move into `Open Order Reports / Processed Reports`.
4. Traveler photos remain available as backup and enrichment when a report line needs verification.

The report importer treats `JobNumber + Part Number` as the duplicate key because one job can contain multiple part lines.

The importer also accepts the current powder schedule database export shape:

- `Part Received Date`
- `JobNumber`
- `PO#`
- `PartNumber`
- `Parts Received by Dept`
- `Parts Leaving Dept`
- `Department Captured`
- `Customer`
- `Departements Used...`

For that feed, set `WEEKLY_OPEN_ORDERS_SHEET_NAME` to `Powder Schedule Database`.

## Runtime Configuration

The public repository does not hardcode company Google Drive or Sheet IDs.
Configure the deployment through Apps Script properties instead.

See `docs/SCRIPT_PROPERTIES.example.md` for the required property names.

## Deploy As Google Apps Script

1. Create a Google Apps Script project.
2. Add the files in `apps-script/`:
   - `Code.gs`
   - `Dashboard.html`
   - `appsscript.json`
3. Enable the Advanced Google Service named `Drive`.
4. Make sure the Google Cloud project behind the script has the Drive API enabled.
5. Set the required script properties from `docs/SCRIPT_PROPERTIES.example.md`.
6. Run `installTriggers()` once from the Apps Script editor.
7. Deploy as a web app.

See `docs/DEPLOYMENT_CHECKLIST.md` for the fuller rollout checklist.
See `docs/OPERATOR_GUIDE.md` for the daily shop-floor workflow.

Recommended access setting for a company dashboard:

- Execute as: `Me`
- Who has access: your company/domain, or specific users if preferred

## Operating Model

The intended V2 workflow is hybrid:

- New photos are checked every 5 minutes.
- A daily sweep runs at 3:00 PM Indiana time.
- A scheduler can press `Digest Photos Now` in the dashboard.

That gives the shop near-real-time capture during the day and a reliable pre-schedule sweep before supervisors make the next-day floor sheet.

## Gemini Vision Setup

The dashboard is wired for Google's stable Flash vision-capable API model, `gemini-2.5-flash`.

To turn on the stronger OCR/photo reader:

1. Create a Gemini API key in Google AI Studio.
2. Open the Apps Script project.
3. Go to Project Settings.
4. Add a script property named `GEMINI_API_KEY`.
5. Paste the API key as the value.
6. Refresh the dashboard and confirm the reader badge says `Gemini Vision`.

The Google AI Ultra plan gives you access to Gemini products, but the Apps Script automation still needs an API key for server-to-server calls.

## Notes

The reader is conservative. If it cannot confidently identify required fields, it still appends the row with `Needs Review` so the item is visible instead of silently disappearing.

Existing low-confidence rows from the old Drive OCR pass should be repaired or reprocessed after Gemini Vision is configured.
