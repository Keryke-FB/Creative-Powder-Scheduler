# Apps Script Properties

Set these in Apps Script under **Project Settings > Script properties**.

Required:

| Property | Purpose |
| --- | --- |
| `SPREADSHEET_ID` | Main Creative Coatings schedule workbook. |
| `PHOTO_FOLDER_ID` | Google Drive folder where new traveler photos are dropped. |
| `WEEKLY_OPEN_ORDERS_SPREADSHEET_ID` | Source Google Sheet for the current Plate Shop open-order report. |
| `POWDER_MASTER_SPREADSHEET_ID` | Powder master lookup workbook. |
| `MASKED_PARTS_SPREADSHEET_ID` | Masked parts lookup workbook. |

Optional:

| Property | Default |
| --- | --- |
| `MASTER_SHEET_NAME` | `Master Capture` |
| `FLOOR_SHEET_NAME` | `Floor Schedule` |
| `WEEKLY_OPEN_ORDERS_SHEET_NAME` | First sheet in the open-order workbook |
| `POWDER_MASTER_SHEET_NAME` | `Powder Master` |
| `MASKED_PARTS_SHEET_NAME` | `Sheet1` |
| `MASKED_PARTS_FOLDER_ID` | Blank |
| `PROCESSED_FOLDER_NAME` | `Processed` |
| `REPORT_FOLDER_NAME` | `Open Order Reports` |
| `PROCESSED_REPORT_FOLDER_NAME` | `Processed Reports` |
| `FORMER_SCHEDULES_FOLDER_NAME` | `Former Schedules` |
| `REVIEW_STATUS` | `Needs Review` |
| `DEFAULT_STATUS` | `On Hold` |
| `DIGEST_LIMIT` | `50` |
| `REPORT_IMPORT_LIMIT` | `10` |
| `GEMINI_MODEL` | `gemini-2.5-flash` |
| `TIMEZONE` | `America/Indiana/Indianapolis` |

Vision OCR:

| Property | Purpose |
| --- | --- |
| `GEMINI_API_KEY` | Enables Gemini Vision traveler extraction. If omitted, Drive OCR fallback is used. |
