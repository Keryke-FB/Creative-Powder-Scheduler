# Operator Guide

## Daily Flow

1. Put native Google Sheets versions of the open-order reports in the `Open Order Reports` Google Drive folder.
2. Put traveler photos in the `Daily Traveler Photos` Google Drive folder when photos are needed as backup.
3. Open the deployed dashboard and confirm the badge says `Connected to Google Apps Script`.
4. Confirm the reader badge says `Gemini Vision` for the stronger photo reader.
5. Let the automatic intake run, or press `Import Open Orders` / `Digest Photos Now` in the dashboard.
6. Review any rows marked `Needs Review`.
7. Correct fields directly in the dashboard row and press `Save`.
8. Choose `Run Tomorrow`, `Hot List`, or `On Hold`.
9. Press `Complete / Send to Shop`.
10. Print from the popup template for supervisors.

## Status Rules

- `On Hold`: stays in Master Capture only.
- `Run Tomorrow`: appears on Floor Schedule with a blank first column.
- `Hot List`: appears on Floor Schedule with `Hot` in the first column.

`On Hold` is the backlog. It keeps the traveler row available for tomorrow or any later day without putting it on the current printable floor schedule.

## Dashboard Buttons

- `Import Open Orders`: imports new Plate Shop open-order Google Sheets and moves processed report sheets into `Open Order Reports / Processed Reports`.
- `Digest Photos Now`: reads new traveler photos immediately instead of waiting for the automatic trigger.
- `Save`: commits edits made directly in that dashboard row back to Master Capture.
- `OK`: clears a `Needs Review` flag after the row has been checked.
- `Complete / Send to Shop`: stamps every current `Run Tomorrow` and `Hot List` row with a shop batch, updates the printable Floor Schedule handoff block, and leaves `On Hold` rows untouched.

When `Complete / Send to Shop` finishes, the dashboard opens a printable shop-floor template. Press `Print` from that surface.

## Print Order

The shop handoff is ordered for line efficiency:

1. Paint color
2. Hot list priority within that color
3. Customer within that priority group

The reason is practical: staying in one color reduces color-change downtime, hot jobs should be pulled first inside that color, and customer grouping helps supervisors understand shipping flow.

## Review Rules

Rows marked `Needs Review` mean the photo reader could not confidently find one or more required fields. The source photo link is available in the row so the scheduler can compare the image and fix the row quickly.

After fixing the row, press `OK` to clear the review flag.

## Reader Rules

The best reader is Gemini Vision on `gemini-2.5-flash`. If the dashboard says `Drive OCR fallback`, the system will still capture what it can, but the row quality will be much weaker and more rows will need review.

## Duplicate Rules

The digest skips a photo if the source photo ID was already captured. It also skips a new photo if it can read a control/work order number that already exists in Master Capture.

That keeps repeated photos from duplicating the schedule backlog.

The open-order report importer uses `JobNumber + Part Number` as the duplicate key, because the same job number can appear on several part lines.
